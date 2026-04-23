import Tesseract from 'tesseract.js'
import { RecognizeScaleTextService } from './generated/services/RecognizeScaleTextService'

export type ScaleOcrStatus = 'idle' | 'reading' | 'done' | 'error'

export interface ParsedWeightResult {
  displayValue: number | null
  suggestion: string
  rawNumber: string
  unit: 'kg' | 'lb' | 't' | 'unknown' | null
}

export interface OcrResult {
  rawText: string
  parsed: ParsedWeightResult
  reliable: boolean
  requiresConfirmation: boolean
  confidence: number
  candidates: OcrCandidate[]
  reasons: string[]
}

export interface CropRect {
  x: number
  y: number
  width: number
  height: number
}

export interface OcrPreprocessResult {
  imageDataUrls: string[]
  candidateTexts: string[]
  cropPreviewUrl: string
  cropRect: CropRect
}

export type OcrSource = 'sevenSegment' | 'tesseract' | 'remote' | 'merged'

export interface OcrCandidate {
  source: OcrSource
  text: string
  parsed: ParsedWeightResult
  confidence: number
  score: number
  reliable: boolean
  reasons: string[]
}

const OCR_CHAR_WHITELIST = '0123456789.,kgKGTtlbLBsS '
const OCR_NUMBER_WHITELIST = '0123456789., '
const MIN_OCR_WIDTH = 1800
const MAX_OCR_WIDTH = 2400
const CONTRAST_FACTOR = 1.55
const THRESHOLD = 145
const WHITE_BORDER = 28
const MIN_PLAUSIBLE_DISPLAY_VALUE = 0
const MAX_PLAUSIBLE_DISPLAY_VALUE = 100000
const MIN_RELIABLE_SCORE = 120
const MIN_RELIABLE_CONFIDENCE = 70
const ENABLE_LOCAL_OCR = false

interface Bbox {
  left: number
  top: number
  right: number
  bottom: number
}

interface ImageVariantOptions {
  mode: 'gray' | 'binary' | 'invertedBinary'
  dilateDark?: boolean
  whiteBorder?: boolean
}

let workerPromise: Promise<Tesseract.Worker> | null = null
let ocrQueue: Promise<unknown> = Promise.resolve()

const getWeightOcrWorker = async () => {
  workerPromise ??= Tesseract.createWorker(
    'eng',
    Tesseract.OEM.LSTM_ONLY,
    {},
    {
      load_system_dawg: '0',
      load_freq_dawg: '0',
      load_punc_dawg: '0',
      load_number_dawg: '0',
    },
  ).then(async (worker) => {
    await worker.setParameters({
      tessedit_char_whitelist: OCR_CHAR_WHITELIST,
      tessedit_pageseg_mode: Tesseract.PSM.SPARSE_TEXT,
      preserve_interword_spaces: '1',
      user_defined_dpi: '150',
    })
    return worker
  })

  return workerPromise
}

const blobToDataUrl = (file: Blob) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => {
      if (typeof reader.result === 'string') {
        resolve(reader.result)
        return
      }
      reject(new Error('Could not prepare the scale photo for OCR.'))
    }
    reader.onerror = () => reject(reader.error ?? new Error('Could not prepare the scale photo for OCR.'))
    reader.readAsDataURL(file)
  })

const loadImageUrl = (src: string, errorMessage: string) =>
  new Promise<HTMLImageElement>((resolve, reject) => {
    const image = new Image()
    image.onload = () => resolve(image)
    image.onerror = () => reject(new Error(errorMessage))
    image.src = src
  })

const loadImageFile = async (file: File) => {
  const dataUrl = await blobToDataUrl(file)
  return loadImageUrl(dataUrl, `Could not load ${file.name}.`)
}

const loadImageSource = (source: File | string) => {
  if (typeof source === 'string') {
    return loadImageUrl(source, 'Could not load the scale photo.')
  }
  return loadImageFile(source)
}

const cloneImageData = (imageData: ImageData) => new ImageData(new Uint8ClampedArray(imageData.data), imageData.width, imageData.height)

const getLuma = (data: Uint8ClampedArray, index: number) => data[index] * 0.299 + data[index + 1] * 0.587 + data[index + 2] * 0.114

const expandBbox = (bbox: Bbox, width: number, height: number, padding: number): Bbox => ({
  left: Math.max(0, bbox.left - padding),
  top: Math.max(0, bbox.top - padding),
  right: Math.min(width - 1, bbox.right + padding),
  bottom: Math.min(height - 1, bbox.bottom + padding),
})

const bboxToCropRect = (bbox: Bbox, width: number, height: number): CropRect => ({
  x: bbox.left / width,
  y: bbox.top / height,
  width: (bbox.right - bbox.left + 1) / width,
  height: (bbox.bottom - bbox.top + 1) / height,
})

const cropRectToBbox = (cropRect: CropRect, width: number, height: number): Bbox => {
  const left = Math.round(Math.max(0, Math.min(0.98, cropRect.x)) * width)
  const top = Math.round(Math.max(0, Math.min(0.98, cropRect.y)) * height)
  const right = Math.round(Math.max(left + 1, Math.min(1, cropRect.x + cropRect.width) * width)) - 1
  const bottom = Math.round(Math.max(top + 1, Math.min(1, cropRect.y + cropRect.height) * height)) - 1
  return {
    left: Math.max(0, Math.min(width - 2, left)),
    top: Math.max(0, Math.min(height - 2, top)),
    right: Math.max(1, Math.min(width - 1, right)),
    bottom: Math.max(1, Math.min(height - 1, bottom)),
  }
}

const cropImageData = (imageData: ImageData, bbox: Bbox) => {
  const canvas = document.createElement('canvas')
  const width = bbox.right - bbox.left + 1
  const height = bbox.bottom - bbox.top + 1
  canvas.width = width
  canvas.height = height
  const ctx = canvas.getContext('2d', { willReadFrequently: true })
  if (!ctx) throw new Error('Could not prepare the scale photo for OCR.')
  const fullCanvas = document.createElement('canvas')
  fullCanvas.width = imageData.width
  fullCanvas.height = imageData.height
  const fullCtx = fullCanvas.getContext('2d')
  if (!fullCtx) throw new Error('Could not prepare the scale photo for OCR.')
  fullCtx.putImageData(imageData, 0, 0)
  ctx.drawImage(fullCanvas, bbox.left, bbox.top, width, height, 0, 0, width, height)
  return ctx.getImageData(0, 0, width, height)
}

const imageDataToPreviewUrl = (imageData: ImageData) => {
  const canvas = document.createElement('canvas')
  canvas.width = imageData.width
  canvas.height = imageData.height
  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error('Could not prepare the crop preview.')
  ctx.putImageData(imageData, 0, 0)
  return canvas.toDataURL('image/jpeg', 0.86)
}

const countInkInRegion = (
  ink: Uint8Array,
  width: number,
  leftRatio: number,
  topRatio: number,
  rightRatio: number,
  bottomRatio: number,
) => {
  const height = Math.floor(ink.length / width)
  const left = Math.max(0, Math.floor(leftRatio * width))
  const right = Math.min(width - 1, Math.ceil(rightRatio * width))
  const top = Math.max(0, Math.floor(topRatio * height))
  const bottom = Math.min(height - 1, Math.ceil(bottomRatio * height))
  let count = 0
  let area = 0
  for (let y = top; y <= bottom; y += 1) {
    for (let x = left; x <= right; x += 1) {
      area += 1
      count += ink[y * width + x]
    }
  }
  return area > 0 ? count / area : 0
}

const recognizeSevenSegmentDigit = (imageData: ImageData, bbox: Bbox, inkMode: 'dark' | 'light') => {
  const width = bbox.right - bbox.left + 1
  const height = bbox.bottom - bbox.top + 1
  if (width < 8 || height < 16) return null
  const ink = new Uint8Array(width * height)
  const { data } = imageData
  let lumaTotal = 0
  let lumaCount = 0
  for (let y = bbox.top; y <= bbox.bottom; y += 1) {
    for (let x = bbox.left; x <= bbox.right; x += 1) {
      lumaTotal += getLuma(data, (y * imageData.width + x) * 4)
      lumaCount += 1
    }
  }
  const mean = lumaTotal / Math.max(1, lumaCount)
  const threshold = inkMode === 'dark' ? mean - 18 : mean + 22
  for (let y = bbox.top; y <= bbox.bottom; y += 1) {
    for (let x = bbox.left; x <= bbox.right; x += 1) {
      const localIndex = (y - bbox.top) * width + (x - bbox.left)
      const luma = getLuma(data, (y * imageData.width + x) * 4)
      ink[localIndex] = inkMode === 'dark' ? Number(luma < threshold) : Number(luma > threshold)
    }
  }

  const segmentRatios = [
    countInkInRegion(ink, width, 0.2, 0.02, 0.8, 0.22),
    countInkInRegion(ink, width, 0.62, 0.12, 0.98, 0.46),
    countInkInRegion(ink, width, 0.62, 0.54, 0.98, 0.88),
    countInkInRegion(ink, width, 0.2, 0.78, 0.8, 0.98),
    countInkInRegion(ink, width, 0.02, 0.54, 0.38, 0.88),
    countInkInRegion(ink, width, 0.02, 0.12, 0.38, 0.46),
    countInkInRegion(ink, width, 0.2, 0.4, 0.8, 0.62),
  ]
  const active = segmentRatios.map((ratio) => ratio > 0.1)
  const digitSegments: Record<string, boolean[]> = {
    '0': [true, true, true, true, true, true, false],
    '1': [false, true, true, false, false, false, false],
    '2': [true, true, false, true, true, false, true],
    '3': [true, true, true, true, false, false, true],
    '4': [false, true, true, false, false, true, true],
    '5': [true, false, true, true, false, true, true],
    '6': [true, false, true, true, true, true, true],
    '7': [true, true, true, false, false, false, false],
    '8': [true, true, true, true, true, true, true],
    '9': [true, true, true, true, false, true, true],
  }
  let bestDigit = ''
  let bestDistance = Number.POSITIVE_INFINITY
  for (const [digit, segments] of Object.entries(digitSegments)) {
    const distance = segments.reduce((sum, expected, index) => sum + Number(expected !== active[index]), 0)
    if (distance < bestDistance) {
      bestDistance = distance
      bestDigit = digit
    }
  }
  return bestDistance <= 2 ? bestDigit : null
}

const getInkBbox = (imageData: ImageData, inkMode: 'dark' | 'light') => {
  const { data, width, height } = imageData
  const lumas: number[] = []
  for (let i = 0; i < data.length; i += 4) lumas.push(getLuma(data, i))
  const mean = lumas.reduce((sum, value) => sum + value, 0) / Math.max(1, lumas.length)
  const threshold = inkMode === 'dark' ? mean - 22 : mean + 28
  let left = width
  let top = height
  let right = -1
  let bottom = -1
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const luma = getLuma(data, (y * width + x) * 4)
      const isInk = inkMode === 'dark' ? luma < threshold : luma > threshold
      if (!isInk) continue
      left = Math.min(left, x)
      top = Math.min(top, y)
      right = Math.max(right, x)
      bottom = Math.max(bottom, y)
    }
  }
  if (right <= left || bottom <= top) return null
  return { left, top, right, bottom }
}

const recognizeSevenSegmentText = (imageData: ImageData) => {
  const attempts = (['dark', 'light'] as const).map((inkMode) => {
    const inkBbox = getInkBbox(imageData, inkMode)
    if (!inkBbox) return ''
    const { data, width } = imageData
    const meanInkWidth = inkBbox.right - inkBbox.left + 1
    const colCounts: number[] = []
    let lumaTotal = 0
    let lumaCount = 0
    for (let y = inkBbox.top; y <= inkBbox.bottom; y += 1) {
      for (let x = inkBbox.left; x <= inkBbox.right; x += 1) {
        lumaTotal += getLuma(data, (y * width + x) * 4)
        lumaCount += 1
      }
    }
    const mean = lumaTotal / Math.max(1, lumaCount)
    const threshold = inkMode === 'dark' ? mean - 18 : mean + 22
    for (let x = inkBbox.left; x <= inkBbox.right; x += 1) {
      let count = 0
      for (let y = inkBbox.top; y <= inkBbox.bottom; y += 1) {
        const luma = getLuma(data, (y * width + x) * 4)
        if (inkMode === 'dark' ? luma < threshold : luma > threshold) count += 1
      }
      colCounts.push(count)
    }

    const minColumnInk = Math.max(2, Math.round((inkBbox.bottom - inkBbox.top + 1) * 0.08))
    const runs: Bbox[] = []
    let runStart: number | null = null
    for (let i = 0; i <= colCounts.length; i += 1) {
      const active = i < colCounts.length && colCounts[i] >= minColumnInk
      if (active && runStart === null) runStart = i
      if ((!active || i === colCounts.length) && runStart !== null) {
        const runEnd = i - 1
        if (runEnd - runStart + 1 >= Math.max(5, meanInkWidth * 0.035)) {
          runs.push({
            left: inkBbox.left + runStart,
            right: inkBbox.left + runEnd,
            top: inkBbox.top,
            bottom: inkBbox.bottom,
          })
        }
        runStart = null
      }
    }

    return runs
      .map((run) => recognizeSevenSegmentDigit(imageData, run, inkMode))
      .filter((digit): digit is string => Boolean(digit))
      .join('')
  })

  return attempts
    .filter((value) => value.length >= 2 && value.length <= 6)
    .sort((a, b) => b.length - a.length)[0] ?? ''
}

const findDarkRegionBbox = (imageData: ImageData): Bbox | null => {
  const { data, width, height } = imageData
  let left = width
  let top = height
  let right = -1
  let bottom = -1

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const index = (y * width + x) * 4
      if (getLuma(data, index) > 86) continue
      left = Math.min(left, x)
      top = Math.min(top, y)
      right = Math.max(right, x)
      bottom = Math.max(bottom, y)
    }
  }

  if (right < left || bottom < top) return null
  const bboxWidth = right - left + 1
  const bboxHeight = bottom - top + 1
  if (bboxWidth < width * 0.08 || bboxHeight < height * 0.08) return null
  return expandBbox({ left, top, right, bottom }, width, height, 16)
}

const findLitTextBbox = (imageData: ImageData, searchBbox: Bbox): Bbox | null => {
  const { data, width, height } = imageData
  let maxLuma = 0
  for (let y = searchBbox.top; y <= searchBbox.bottom; y += 1) {
    for (let x = searchBbox.left; x <= searchBbox.right; x += 1) {
      const index = (y * width + x) * 4
      maxLuma = Math.max(maxLuma, getLuma(data, index))
    }
  }

  const threshold = Math.max(118, maxLuma - 72)
  let left = width
  let top = height
  let right = -1
  let bottom = -1
  let count = 0

  for (let y = searchBbox.top; y <= searchBbox.bottom; y += 1) {
    for (let x = searchBbox.left; x <= searchBbox.right; x += 1) {
      const index = (y * width + x) * 4
      const luma = getLuma(data, index)
      if (luma < threshold) continue
      left = Math.min(left, x)
      top = Math.min(top, y)
      right = Math.max(right, x)
      bottom = Math.max(bottom, y)
      count += 1
    }
  }

  if (count < 20 || right < left || bottom < top) return null
  const bboxWidth = right - left + 1
  const bboxHeight = bottom - top + 1
  if (bboxWidth < 18 || bboxHeight < 12) return null
  return expandBbox({ left, top, right, bottom }, width, height, Math.round(Math.max(bboxWidth, bboxHeight) * 0.25))
}

const dilateDarkPixels = (imageData: ImageData, radius = 1) => {
  const source = new Uint8ClampedArray(imageData.data)
  const { data, width, height } = imageData
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      let makeDark = false
      for (let dy = -radius; dy <= radius && !makeDark; dy += 1) {
        for (let dx = -radius; dx <= radius; dx += 1) {
          const nx = x + dx
          const ny = y + dy
          if (nx < 0 || nx >= width || ny < 0 || ny >= height) continue
          const sourceIndex = (ny * width + nx) * 4
          if (source[sourceIndex] < 128) {
            makeDark = true
            break
          }
        }
      }
      if (!makeDark) continue
      const targetIndex = (y * width + x) * 4
      data[targetIndex] = 0
      data[targetIndex + 1] = 0
      data[targetIndex + 2] = 0
    }
  }
}

const imageDataToUrl = (imageData: ImageData, options: ImageVariantOptions) => {
  const canvas = document.createElement('canvas')
  const border = options.whiteBorder ? WHITE_BORDER : 0
  canvas.width = imageData.width + border * 2
  canvas.height = imageData.height + border * 2
  const ctx = canvas.getContext('2d', { willReadFrequently: true })
  if (!ctx) throw new Error('Could not prepare the scale photo for OCR.')

  const nextImageData = cloneImageData(imageData)
  const { data } = nextImageData
  for (let i = 0; i < data.length; i += 4) {
    const gray = getLuma(data, i)
    const contrasted = Math.max(0, Math.min(255, (gray - 128) * CONTRAST_FACTOR + 128))
    const binaryValue = contrasted > THRESHOLD ? 255 : 0
    const value = options.mode === 'gray'
      ? contrasted
      : options.mode === 'invertedBinary'
        ? 255 - binaryValue
        : binaryValue
    data[i] = value
    data[i + 1] = value
    data[i + 2] = value
  }

  if (options.dilateDark) dilateDarkPixels(nextImageData, 1)

  ctx.fillStyle = '#fff'
  ctx.fillRect(0, 0, canvas.width, canvas.height)
  ctx.putImageData(nextImageData, border, border)
  return canvas.toDataURL('image/png')
}

export const preprocessImageForOcr = async (source: File | string, cropRect?: CropRect): Promise<OcrPreprocessResult> => {
  const image = await loadImageSource(source)
  const upscale = Math.max(1, MIN_OCR_WIDTH / image.naturalWidth)
  const downscale = MAX_OCR_WIDTH / image.naturalWidth
  const scale = Math.min(upscale, downscale)
  const width = Math.max(1, Math.round(image.naturalWidth * scale))
  const height = Math.max(1, Math.round(image.naturalHeight * scale))
  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height

  const ctx = canvas.getContext('2d', { willReadFrequently: true })
  if (!ctx) throw new Error('Could not prepare the scale photo for OCR.')

  ctx.imageSmoothingEnabled = false
  ctx.drawImage(image, 0, 0, width, height)
  const imageData = ctx.getImageData(0, 0, width, height)
  const fullImageBbox = { left: 0, top: 0, right: width - 1, bottom: height - 1 }
  let selectedBbox = cropRect ? cropRectToBbox(cropRect, width, height) : fullImageBbox

  if (!cropRect) {
    const darkRegionBbox = findDarkRegionBbox(imageData)
    const litTextBbox = darkRegionBbox ? findLitTextBbox(imageData, darkRegionBbox) : null
    selectedBbox = litTextBbox ?? darkRegionBbox ?? fullImageBbox
  }

  const selectedCrop = cropImageData(imageData, selectedBbox)
  const sevenSegmentText = ENABLE_LOCAL_OCR ? recognizeSevenSegmentText(selectedCrop) : ''
  const imageDataVariants = cropRect
    ? [selectedCrop]
    : [
      selectedCrop,
      imageData,
    ]

  return {
    candidateTexts: sevenSegmentText ? [sevenSegmentText] : [],
    imageDataUrls: imageDataVariants.flatMap((variant) => [
    imageDataToUrl(variant, { mode: 'invertedBinary', dilateDark: true, whiteBorder: true }),
    imageDataToUrl(variant, { mode: 'invertedBinary', whiteBorder: true }),
    imageDataToUrl(variant, { mode: 'binary', whiteBorder: true }),
    imageDataToUrl(variant, { mode: 'gray', whiteBorder: true }),
    ]),
    cropPreviewUrl: imageDataToPreviewUrl(selectedCrop),
    cropRect: bboxToCropRect(selectedBbox, width, height),
  }
}

const normalizeNumericText = (value: string) => {
  const cleaned = value.trim().replace(/\s/g, '')
  const hasComma = cleaned.includes(',')
  const hasDot = cleaned.includes('.')

  if (hasComma && hasDot) {
    const lastComma = cleaned.lastIndexOf(',')
    const lastDot = cleaned.lastIndexOf('.')
    if (lastComma > lastDot) return cleaned.replace(/\./g, '').replace(',', '.')
    return cleaned.replace(/,/g, '')
  }

  if (hasComma) return cleaned.replace(',', '.')
  return cleaned
}

const formatDisplayValue = (value: number) => {
  const rounded = Math.round(value * 1000) / 1000
  return Number.isInteger(rounded) ? String(rounded) : String(rounded).replace(/0+$/, '').replace(/\.$/, '')
}

const isPlausibleDisplayValue = (value: number | null) =>
  value !== null && Number.isFinite(value) && value > MIN_PLAUSIBLE_DISPLAY_VALUE && value <= MAX_PLAUSIBLE_DISPLAY_VALUE

const parseUnit = (value: string): ParsedWeightResult['unit'] => {
  const lower = value.toLowerCase()
  if (/\bkg\b|kilogram/.test(lower)) return 'kg'
  if (/\blbs?\b|pound/.test(lower)) return 'lb'
  if (/\bt\b|ton|tonne/.test(lower)) return 't'
  return 'unknown'
}

export const parseWeightFromOcrText = (text: string): ParsedWeightResult => {
  const normalizedText = text.replace(/\s+/g, ' ').trim()
  const numberPattern = /\d+(?:[.,]\d+)?/g
  const matches = Array.from(normalizedText.matchAll(numberPattern))
  const candidates = matches
    .map((match) => {
      const rawNumber = match[0]
      const start = match.index ?? 0
      const end = start + rawNumber.length
      const context = normalizedText.slice(Math.max(0, start - 12), Math.min(normalizedText.length, end + 12))
      const numericValue = Number(normalizeNumericText(rawNumber))
      const unit = parseUnit(context)
      const unitScore = unit === 'kg' || unit === 'lb' || unit === 't' ? 2 : 0
      const decimalScore = /[.,]/.test(rawNumber) ? 1 : 0
      return { rawNumber, numericValue, unit, score: unitScore + decimalScore }
    })
    .filter((candidate) => Number.isFinite(candidate.numericValue) && candidate.numericValue >= 0)

  const unitIndex = normalizedText.search(/\b(?:kg|KG|Kg|kG|lb|LB|Lb|lbs|LBS|t|T|ton|tonne)\b/)
  if (unitIndex >= 0) {
    const priorIntegerMatches = matches.filter((match) => {
      const start = match.index ?? 0
      return start < unitIndex && !/[.,]/.test(match[0])
    })
    const joinedDigits = priorIntegerMatches.map((match) => match[0]).join('')
    if (joinedDigits.length >= 2 && joinedDigits.length <= 6) {
      const start = priorIntegerMatches[0]?.index ?? 0
      const context = normalizedText.slice(start, Math.min(normalizedText.length, unitIndex + 8))
      const numericValue = Number(joinedDigits)
      const unit = parseUnit(context)
      candidates.push({
        rawNumber: joinedDigits,
        numericValue,
        unit,
        score: (unit === 'kg' || unit === 'lb' || unit === 't' ? 4 : 0) + 2,
      })
    }
  }

  candidates
    .sort((a, b) => b.score - a.score || b.numericValue - a.numericValue)

  const best = candidates[0]
  if (!best) {
    return { displayValue: null, suggestion: '', rawNumber: '', unit: null }
  }

  return {
    displayValue: best.numericValue,
    suggestion: formatDisplayValue(best.numericValue),
    rawNumber: best.rawNumber,
    unit: best.unit,
  }
}

const hasKgUnit = (text: string) => /\bkg\b|kilogram/i.test(text)
const hasPoundUnit = (text: string) => /\blbs?\b|pound/i.test(text)
const hasTonneUnit = (text: string) => /\bt\b|ton|tonne/i.test(text)

const mergeNumericAndUnitText = (numericText: string, unitText: string) => {
  const numberMatch = numericText.match(/\d+(?:[.,]\d+)?/)
  if (!numberMatch) return ''
  if (hasKgUnit(unitText)) return `${numberMatch[0]} kg`
  if (hasPoundUnit(unitText)) return `${numberMatch[0]} lb`
  if (hasTonneUnit(unitText)) return `${numberMatch[0]} t`
  return numberMatch[0]
}

const getAgreementCount = (candidate: OcrCandidate, candidates: OcrCandidate[]) => {
  const value = candidate.parsed.displayValue
  if (value === null) return 0
  return candidates.filter((other) => (
    other !== candidate &&
    other.parsed.displayValue !== null &&
    Math.abs(other.parsed.displayValue - value) <= Math.max(0.005, value * 0.01)
  )).length
}

const isCleanUnitlessNumber = (candidate: OcrCandidate) => {
  if (candidate.parsed.unit !== 'unknown' || !candidate.parsed.rawNumber) return false
  const compactText = candidate.text.replace(/\s+/g, '')
  return compactText === candidate.parsed.rawNumber && isPlausibleDisplayValue(candidate.parsed.displayValue)
}

const isMicrosoftOnlyCandidate = (candidate: OcrCandidate) =>
  !ENABLE_LOCAL_OCR &&
  candidate.source === 'remote' &&
  isPlausibleDisplayValue(candidate.parsed.displayValue) &&
  candidate.confidence >= MIN_RELIABLE_CONFIDENCE

const createOcrCandidate = (
  text: string,
  source: OcrSource,
  confidence: number,
  sourceBoost = 0,
): OcrCandidate => {
  const parsed = parseWeightFromOcrText(text)
  const reasons: string[] = []
  if (!parsed.suggestion || parsed.displayValue === null) reasons.push('No parseable number')
  const hasUnit = parsed.unit === 'kg' || parsed.unit === 'lb' || parsed.unit === 't'
  const plausible = isPlausibleDisplayValue(parsed.displayValue)
  if (!hasUnit && parsed.suggestion) reasons.push('Unit was not recognized')
  if (parsed.displayValue !== null && !plausible) {
    reasons.push(`Outside ${MIN_PLAUSIBLE_DISPLAY_VALUE}-${MAX_PLAUSIBLE_DISPLAY_VALUE} display range`)
  }

  const unitScore = hasUnit ? 32 : -16
  const decimalScore = parsed.rawNumber.includes('.') || parsed.rawNumber.includes(',') ? 6 : 0
  const plausibleScore = plausible ? 24 : -80
  const textScore = Math.min(8, text.trim().length)
  const sourceScore = source === 'remote' ? 14 : source === 'merged' ? 10 : source === 'sevenSegment' ? 6 : 0
  const microsoftOnlyScore = !ENABLE_LOCAL_OCR && source === 'remote' && plausible ? 42 : 0
  const score = parsed.suggestion
    ? confidence + unitScore + decimalScore + plausibleScore + textScore + sourceScore + sourceBoost + microsoftOnlyScore
    : 0
  const reliable = (
    score >= MIN_RELIABLE_SCORE &&
    confidence >= MIN_RELIABLE_CONFIDENCE &&
    plausible &&
    (hasUnit || (!ENABLE_LOCAL_OCR && source === 'remote'))
  )

  return { source, text, parsed, confidence, score, reliable, reasons }
}

const selectBestOcrResult = (candidates: OcrCandidate[], rawTexts: string[]): OcrResult => {
  const usableCandidates = candidates
    .filter((candidate) => candidate.parsed.suggestion && isPlausibleDisplayValue(candidate.parsed.displayValue))
    .map((candidate) => {
      const agreementCount = getAgreementCount(candidate, candidates)
      const agreementBonusCount = Math.min(agreementCount, 3)
      const cleanUnitlessWithAgreement = agreementCount >= 1 && isCleanUnitlessNumber(candidate)
      return {
        ...candidate,
        score: candidate.score + agreementBonusCount * 18,
        reliable: candidate.reliable || isMicrosoftOnlyCandidate(candidate) || (
          candidate.score + agreementBonusCount * 18 >= MIN_RELIABLE_SCORE &&
          candidate.confidence >= MIN_RELIABLE_CONFIDENCE &&
          (candidate.parsed.unit !== 'unknown' || cleanUnitlessWithAgreement)
        ),
        reasons: agreementCount > 0 ? [...candidate.reasons, 'Agrees across OCR passes'] : candidate.reasons,
      }
    })
    .sort((a, b) => b.score - a.score)

  const best = usableCandidates[0]
  if (best) {
    const rawText = [best.text, ...rawTexts.filter(Boolean)]
      .filter(Boolean)
      .join('\n---\n')
    return {
      rawText,
      parsed: best.parsed,
      reliable: best.reliable,
      requiresConfirmation: true,
      confidence: Math.max(0, Math.min(100, best.confidence)),
      candidates: usableCandidates,
      reasons: best.reasons,
    }
  }

  const rawText = rawTexts.filter(Boolean).join('\n---\n')
  const parsed = parseWeightFromOcrText(rawTexts.join('\n'))
  return {
    rawText,
    parsed: isPlausibleDisplayValue(parsed.displayValue) ? parsed : { displayValue: null, suggestion: '', rawNumber: '', unit: null },
    reliable: false,
    requiresConfirmation: Boolean(parsed.suggestion),
    confidence: 0,
    candidates,
    reasons: rawText ? ['No OCR candidate passed plausibility checks'] : ['No OCR text was returned'],
  }
}

const dataUrlToBlob = async (dataUrl: string) => {
  const response = await fetch(dataUrl)
  return response.blob()
}

const dataUrlToFlowFile = (dataUrl: string) => {
  const match = dataUrl.match(/^data:([^;,]+);base64,(.*)$/)
  if (!match) return { contentBytes: dataUrl, name: 'scale-crop.png' }
  const [, mimeType, contentBytes] = match
  const extension = mimeType === 'image/jpeg' ? 'jpg' : mimeType === 'image/bmp' ? 'bmp' : 'png'
  return { contentBytes, name: `scale-crop.${extension}` }
}

const getRemoteOcrEndpoint = () => String(import.meta.env.VITE_WEIGHT_OCR_ENDPOINT ?? '').trim()

const getRemoteTextItems = (payload: unknown): { text: string; confidence: number }[] => {
  const items: { text: string; confidence: number }[] = []
  const pushText = (text: unknown, confidence: unknown) => {
    if (typeof text !== 'string' || !text.trim()) return
    const numericConfidence = typeof confidence === 'number' && Number.isFinite(confidence)
      ? confidence <= 1 ? confidence * 100 : confidence
      : 82
    items.push({ text: text.trim(), confidence: Math.max(0, Math.min(100, numericConfidence)) })
  }

  if (payload && typeof payload === 'object') {
    const objectPayload = payload as Record<string, unknown>
    pushText(objectPayload.text, objectPayload.confidence)
    const lines = Array.isArray(objectPayload.lines) ? objectPayload.lines : []
    for (const line of lines) {
      if (line && typeof line === 'object') {
        const linePayload = line as Record<string, unknown>
        pushText(linePayload.text ?? linePayload.content, linePayload.confidence)
      }
    }
  }

  return items
}

const runRemoteWeightOcr = async (imageDataUrl?: string) => {
  const candidates: OcrCandidate[] = []

  if (imageDataUrl) {
    try {
      const result = await RecognizeScaleTextService.Run({
        file: dataUrlToFlowFile(imageDataUrl),
      })
      candidates.push(...getRemoteTextItems(result.data).map((item) => createOcrCandidate(item.text, 'remote', item.confidence)))
    } catch {
      // Keep local/browser OCR usable when the Power Apps host or flow connection is unavailable.
    }
  }

  const endpoint = getRemoteOcrEndpoint()
  if (!endpoint || !imageDataUrl) return candidates

  try {
    const body = new FormData()
    body.append('image', await dataUrlToBlob(imageDataUrl), 'scale-crop.jpg')
    const response = await fetch(endpoint, { method: 'POST', body })
    if (!response.ok) return candidates
    const payload: unknown = await response.json()
    candidates.push(...getRemoteTextItems(payload).map((item) => createOcrCandidate(item.text, 'remote', item.confidence)))
  } catch {
    return candidates
  }

  return candidates
}

export const runWeightOcr = async (
  imageDataUrls: string | string[],
  candidateTexts: string[] = [],
  options: { remoteImageDataUrl?: string } = {},
): Promise<OcrResult> => {
  const run = async () => {
    const imageVariants = Array.isArray(imageDataUrls) ? imageDataUrls : [imageDataUrls]
    const candidates: OcrCandidate[] = []
    const rawTexts: string[] = []

    candidates.push(...await runRemoteWeightOcr(options.remoteImageDataUrl ?? imageVariants[0]))

    if (ENABLE_LOCAL_OCR) {
      const worker = await getWeightOcrWorker()
      const pageSegModes = [Tesseract.PSM.SINGLE_LINE, Tesseract.PSM.SPARSE_TEXT]
      const unitTexts: string[] = []
      const numericTexts: string[] = [...candidateTexts]

      for (const candidateText of candidateTexts) {
        candidates.push(createOcrCandidate(candidateText, 'sevenSegment', 82))
      }

      for (const imageDataUrl of imageVariants) {
        for (const pageSegMode of pageSegModes) {
          for (const whitelist of [OCR_CHAR_WHITELIST, OCR_NUMBER_WHITELIST]) {
            await worker.setParameters({
              tessedit_char_whitelist: whitelist,
              tessedit_pageseg_mode: pageSegMode,
            })
            const result = await worker.recognize(imageDataUrl)
            const rawText = result.data.text
            rawTexts.push(rawText)
            if (whitelist === OCR_CHAR_WHITELIST) unitTexts.push(rawText)
            else numericTexts.push(rawText)
            candidates.push(createOcrCandidate(rawText, 'tesseract', result.data.confidence))
          }
        }
      }

      for (const numericText of numericTexts) {
        for (const unitText of unitTexts) {
          const mergedText = mergeNumericAndUnitText(numericText, unitText)
          if (!mergedText) continue
          candidates.push(createOcrCandidate(mergedText, 'merged', 80, 4))
        }
      }
    }

    return selectBestOcrResult(candidates, rawTexts)
  }

  const nextRun = ocrQueue.then(run, run)
  ocrQueue = nextRun.catch(() => undefined)
  return nextRun
}
