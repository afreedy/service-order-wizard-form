import { beforeEach, describe, expect, it, vi } from 'vitest'

const mockedRun = vi.fn()
const mockedCreateWorker = vi.fn()

vi.mock('./generated/services/RecognizeScaleTextService', () => ({
  RecognizeScaleTextService: {
    Run: mockedRun,
  },
}))

vi.mock('tesseract.js', () => {
  const module = {
    createWorker: mockedCreateWorker,
    OEM: {
      LSTM_ONLY: 1,
    },
    PSM: {
      SINGLE_LINE: 7,
      SPARSE_TEXT: 11,
    },
  }

  return {
    default: module,
    ...module,
  }
})

const loadWeightOcrModule = async () => import('./weightOcr')

const makeWorker = (texts: string[]) => {
  let index = 0
  return {
    setParameters: vi.fn().mockResolvedValue(undefined),
    recognize: vi.fn().mockImplementation(async () => {
      const text = texts[Math.min(index, texts.length - 1)] ?? ''
      index += 1
      return {
        data: {
          text,
          confidence: 92,
        },
      }
    }),
  }
}

describe('runWeightOcr fallback pipeline', () => {
  beforeEach(() => {
    vi.resetModules()
    vi.clearAllMocks()
  })

  it('returns the Microsoft OCR result without running local OCR when remote output is usable', async () => {
    mockedRun.mockResolvedValue({
      data: {
        text: '1234 kg',
        confidence: 95,
      },
    })
    mockedCreateWorker.mockResolvedValue(makeWorker(['']))

    const { runWeightOcr } = await loadWeightOcrModule()
    const result = await runWeightOcr(['variant-a'], ['1234'], {
      remoteImageDataUrl: 'data:image/png;base64,AAA',
    })

    expect(result.parsed.suggestion).toBe('1234')
    expect(result.reliable).toBe(true)
    expect(result.reasons).not.toContain('Fell back to local OCR')
    expect(mockedCreateWorker).not.toHaveBeenCalled()
  })

  it('runs local OCR fallback when Microsoft OCR is low confidence', async () => {
    mockedRun.mockResolvedValue({
      data: {
        text: '1234 kg',
        confidence: 35,
      },
    })
    mockedCreateWorker.mockResolvedValue(
      makeWorker(['1234 kg', '1234', '1234 kg', '1234']),
    )

    const { runWeightOcr } = await loadWeightOcrModule()
    const result = await runWeightOcr(['variant-a'], ['1234'], {
      remoteImageDataUrl: 'data:image/png;base64,AAA',
    })

    expect(mockedCreateWorker).toHaveBeenCalledTimes(1)
    expect(result.parsed.suggestion).toBe('1234')
    expect(result.reasons).toContain('Microsoft OCR confidence too low')
    expect(result.reasons).toContain('Fell back to local OCR')
  })

  it('runs local OCR fallback when Microsoft OCR returns no text', async () => {
    mockedRun.mockResolvedValue({
      data: {
        lines: [],
      },
    })
    mockedCreateWorker.mockResolvedValue(
      makeWorker(['987 kg', '987', '987 kg', '987']),
    )

    const { runWeightOcr } = await loadWeightOcrModule()
    const result = await runWeightOcr(['variant-a'], ['987'], {
      remoteImageDataUrl: 'data:image/png;base64,AAA',
    })

    expect(mockedCreateWorker).toHaveBeenCalledTimes(1)
    expect(result.parsed.suggestion).toBe('987')
    expect(result.reasons).toContain('Microsoft OCR returned no text')
    expect(result.reasons).toContain('Fell back to local OCR')
  })
})

describe('parseWeightFromOcrText', () => {
  beforeEach(() => {
    vi.resetModules()
    vi.clearAllMocks()
  })

  it('parses decimal weights and recognized units', async () => {
    const { parseWeightFromOcrText } = await loadWeightOcrModule()

    expect(parseWeightFromOcrText('Gross weight 12.50 kg')).toEqual({
      displayValue: 12.5,
      suggestion: '12.5',
      rawNumber: '12.50',
      unit: 'kg',
    })
  })

  it('joins separated digits before a unit when OCR splits the number', async () => {
    const { parseWeightFromOcrText } = await loadWeightOcrModule()

    expect(parseWeightFromOcrText('1 234 kg')).toEqual({
      displayValue: 1234,
      suggestion: '1234',
      rawNumber: '1234',
      unit: 'kg',
    })
  })

  it('returns an empty suggestion when no number is present', async () => {
    const { parseWeightFromOcrText } = await loadWeightOcrModule()

    expect(parseWeightFromOcrText('stable reading unavailable')).toEqual({
      displayValue: null,
      suggestion: '',
      rawNumber: '',
      unit: null,
    })
  })
})
