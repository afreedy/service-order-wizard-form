/* ═══════════════════════════════════════════════════════
   CORA Environment — Service Order Wizard
   Single-file architecture: ~18 inline icons, 4-step
   wizard, CRUD via PowerApps SharePoint connector.
   ═══════════════════════════════════════════════════════ */

import { useEffect, useState, useCallback, useMemo, useRef } from 'react'
import type { ChangeEvent, DragEvent, PointerEvent, ReactNode } from 'react'
import { getContext } from '@microsoft/power-apps/app'
import type { Drivers1Read } from './generated/models/Drivers1Model'
import type { Customer_area_data_clean_finalRead } from './generated/models/Customer_area_data_clean_finalModel'
import type { VehiclesRead, VehiclesWrite } from './generated/models/VehiclesModel'
import type { Waste_CategoriesRead, Waste_CategoriesWrite } from './generated/models/Waste_CategoriesModel'
import type { ServiceOrderWasteItemsRead, ServiceOrderWasteItemsWrite } from './generated/models/ServiceOrderWasteItemsModel'
import type { ServiceOrderProofQueueRead, ServiceOrderProofQueueWrite } from './generated/models/ServiceOrderProofQueueModel'
import type { ServiceOrdersRead, ServiceOrdersWrite } from './generated/models/ServiceOrdersModel'
import { Drivers1Service } from './generated/services/Drivers1Service'
import { Customer_area_data_clean_finalService } from './generated/services/Customer_area_data_clean_finalService'
import { VehiclesService } from './generated/services/VehiclesService'
import { Waste_CategoriesService } from './generated/services/Waste_CategoriesService'
import { ServiceOrderWasteItemsService } from './generated/services/ServiceOrderWasteItemsService'
import { ServiceOrderProofQueueService } from './generated/services/ServiceOrderProofQueueService'
import { ServiceOrdersService } from './generated/services/ServiceOrdersService'
import { preprocessImageForOcr, runWeightOcr } from './weightOcr'
import type { CropRect } from './weightOcr'
import type { ScaleOcrStatus } from './weightOcr'
import logoImage from './assets/logo.png'
import './App.css'

/* ═══════════════════════════════════════════════════════
   Inline SVG Icons (Lucide-style, ~18 icons)
   ═══════════════════════════════════════════════════════ */
const p = { fill: 'none', stroke: 'currentColor', strokeWidth: 1.8, strokeLinecap: 'round' as const, strokeLinejoin: 'round' as const }

const I = {
  gear:     <svg viewBox="0 0 24 24" {...p}><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>,
  plus:     <svg viewBox="0 0 24 24" {...p}><path d="M12 5v14M5 12h14"/></svg>,
  clipList: <svg viewBox="0 0 24 24" {...p}><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="8" y="2" width="8" height="4" rx="1"/><path d="M9 12h6M9 16h6"/></svg>,
  user:     <svg viewBox="0 0 24 24" {...p}><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4 3.5-7 8-7s8 3 8 7"/></svg>,
  truck:    <svg viewBox="0 0 24 24" {...p}><path d="M1 3h15v13H1z"/><path d="M16 8h4l3 4v4h-7V8z"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg>,
  recycle:  <svg viewBox="0 0 24 24" {...p}><path d="M7 19H4.815a1.83 1.83 0 0 1-1.57-.881 1.785 1.785 0 0 1-.004-1.784L7.196 9.5"/><path d="M11 19h8.203a1.83 1.83 0 0 0 1.556-.89 1.784 1.784 0 0 0 0-1.775l-1.226-2.12"/><path d="M14 16l3 3 3-3"/><path d="m8.293 13.596-4.6-8.1a1.786 1.786 0 0 1 .004-1.785A1.83 1.83 0 0 1 5.267 3h2.466"/><path d="m12.5 5.5 2-3.5 2 3.5"/><path d="M15.4 8.9 9.8 19.5"/></svg>,
  mapPin:   <svg viewBox="0 0 24 24" {...p}><path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0Z"/><circle cx="12" cy="10" r="3"/></svg>,
  building: <svg viewBox="0 0 24 24" {...p}><rect x="4" y="2" width="16" height="20" rx="2"/><path d="M9 22v-4h6v4M8 6h.01M16 6h.01M12 6h.01M12 10h.01M12 14h.01M16 10h.01M16 14h.01M8 10h.01M8 14h.01"/></svg>,
  calendar: <svg viewBox="0 0 24 24" {...p}><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/></svg>,
  fileTxt:  <svg viewBox="0 0 24 24" {...p}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M9 13h6M9 17h3"/></svg>,
  check:    <svg viewBox="0 0 24 24" {...p}><polyline points="20 6 9 17 4 12"/></svg>,
  chevR:    <svg viewBox="0 0 24 24" {...p}><polyline points="9 18 15 12 9 6"/></svg>,
  chevL:    <svg viewBox="0 0 24 24" {...p}><polyline points="15 18 9 12 15 6"/></svg>,
  send:     <svg viewBox="0 0 24 24" {...p}><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>,
  refresh:  <svg viewBox="0 0 24 24" {...p}><path d="M1 4v6h6"/><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/></svg>,
  menu:     <svg viewBox="0 0 24 24" {...p}><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>,
  x:        <svg viewBox="0 0 24 24" {...p}><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>,
  trash:    <svg viewBox="0 0 24 24" {...p}><path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>,
  leaf:     <svg viewBox="0 0 24 24" {...p}><path d="M17 8C8 10 5.9 16.17 3.82 21.34l1.89-.82 .18-.39"/><path d="M20.18 2.82A5.56 5.56 0 0 0 17 2c-3 0-7 2-7 7 0 3.5 1.9 6.5 5 8 3.5-2 5.56-5.5 5.56-9.5a5.56 5.56 0 0 0-.38-2.68Z"/></svg>,
}

/* ═══════════════════════════════════════════════════════
   Types
   ═══════════════════════════════════════════════════════ */
type Load = 'loading' | 'loaded' | 'error'
type View = 'form' | 'list' | 'admin'
type AdminTab = 'drivers' | 'vehicles' | 'waste'
type ProtectedDestination = { view: 'list' } | { view: 'admin'; tab: AdminTab }

interface Toast { type: 'success' | 'error'; text: string }

interface WasteLine {
  id: string
  WasteCategory: string
  Tonnage: string
  scalePhotoFile?: File
  scalePhotoPreviewUrl?: string
  scaleOcrCropPreviewUrl?: string
  scaleOcrCrop?: CropRect
  scaleOcrStatus?: ScaleOcrStatus
  scaleOcrText?: string
  scaleOcrSuggestion?: string
  scaleOcrConfidence?: number
  scaleOcrReasons?: string[]
  scaleOcrError?: string
  scaleOcrRequestId?: string
}

interface Form {
  Title: string; Customer: string; CustomerName: string
  CustomerLocation: string; CustomerTenant: string
  CustomerLevel4: string; CustomerLevel5: string; CustomerLevel6: string; CustomerLevel7: string
  CustomerLevel8: string; CustomerLevel9: string; CustomerLevel10: string
  IsAdhocCustomer: boolean; DriverName: string
  VehicleNumber: string; DateOfCollection: string
  WasteItems: WasteLine[]
  Notes: string
}

interface ProofUploadRequest {
  serviceOrderId: number
  orderTitle: string
  signatureFileName: string
  signatureBase64: string
  beforePhotoFileName?: string
  beforePhotoBase64?: string
  afterPhotoFileName?: string
  afterPhotoBase64?: string
}

interface ProofUploadResponse {
  queueItem?: ServiceOrderProofQueueRead
  queued: boolean
  signatureUrl?: string
  beforePhotoUrl?: string
  afterPhotoUrl?: string
}

type DriverOptionSource = Drivers1Read
type CustomerAreaOption = Customer_area_data_clean_finalRead
type CustomerLevelField =
  | 'CustomerName'
  | 'CustomerLocation'
  | 'CustomerTenant'
  | 'CustomerLevel4'
  | 'CustomerLevel5'
  | 'CustomerLevel6'
  | 'CustomerLevel7'
  | 'CustomerLevel8'
  | 'CustomerLevel9'
  | 'CustomerLevel10'

const CUSTOMER_LEVELS: { key: CustomerLevelField; label: string; placeholder: string; searchPlaceholder: string }[] = [
  { key: 'CustomerName', label: 'Customer', placeholder: 'Select customer…', searchPlaceholder: 'Search customers...' },
  { key: 'CustomerLocation', label: 'Location', placeholder: 'Select location…', searchPlaceholder: 'Search locations...' },
  { key: 'CustomerTenant', label: 'Sub-location', placeholder: 'Select sub-location…', searchPlaceholder: 'Search sub-locations...' },
  { key: 'CustomerLevel4', label: 'Level 4', placeholder: 'Select level 4…', searchPlaceholder: 'Search level 4...' },
  { key: 'CustomerLevel5', label: 'Level 5', placeholder: 'Select level 5…', searchPlaceholder: 'Search level 5...' },
  { key: 'CustomerLevel6', label: 'Level 6', placeholder: 'Select level 6…', searchPlaceholder: 'Search level 6...' },
  { key: 'CustomerLevel7', label: 'Level 7', placeholder: 'Select level 7…', searchPlaceholder: 'Search level 7...' },
  { key: 'CustomerLevel8', label: 'Level 8', placeholder: 'Select level 8…', searchPlaceholder: 'Search level 8...' },
  { key: 'CustomerLevel9', label: 'Level 9', placeholder: 'Select level 9…', searchPlaceholder: 'Search level 9...' },
  { key: 'CustomerLevel10', label: 'Level 10', placeholder: 'Select level 10…', searchPlaceholder: 'Search level 10...' },
]

const ADMIN_EMAILS = [
  'bizdev@cora-environment.com',
  // Add the second admin email here.
].map((email) => email.toLowerCase())
const ADMIN_PASSCODE = 'CORA2026'
const ADMIN_UNLOCK_STORAGE_KEY = 'cora-admin-unlocked'

const normalize = (value: string) => value.trim().toLowerCase()

const getDriverName = (row: DriverOptionSource) => {
  const dynamicRow = row as Record<string, unknown>
  const candidate = dynamicRow.DriverName
  return (typeof candidate === 'string' ? candidate.trim() : '') || row.Title?.trim() || ''
}

const uniqueValues = (values: string[]) => Array.from(new Set(values.filter(Boolean)))
const readString = (value: unknown) => (typeof value === 'string' ? value.trim() : '')

const getCustomerName = (row: CustomerAreaOption) => row.Title?.trim() || ''
const getCustomerArea = (row: CustomerAreaOption) => {
  const dynamicRow = row as Record<string, unknown>
  const candidate = dynamicRow.Area ?? dynamicRow.field_1
  return readString(candidate)
}
const getCustomerSubLocationRaw = (row: CustomerAreaOption) => {
  const dynamicRow = row as Record<string, unknown>
  const candidate =
    dynamicRow.SubLocation ??
    dynamicRow.field_2 ??
    dynamicRow.Sub_x002d_Location ??
    dynamicRow.Sub_x0020_Location ??
    dynamicRow.Sub_x0020_location ??
    dynamicRow.SubCategory ??
    dynamicRow.Sub_x002d_Category ??
    dynamicRow.Sub_x0020_Category ??
    dynamicRow.Tower ??
    dynamicRow.LocationLevel3 ??
    dynamicRow.Level3 ??
    dynamicRow.CustomerTenant
  return readString(candidate)
}
const getCustomerSubLocationParts = (row: CustomerAreaOption) =>
  getCustomerSubLocationRaw(row)
    .split('/')
    .map((part) => part.trim())
    .filter(Boolean)

const getCustomerLevelValue = (row: CustomerAreaOption, levelIndex: number) => {
  const dynamicRow = row as Record<string, unknown>
  if (levelIndex === 0) return getCustomerName(row)
  if (levelIndex === 1) return getCustomerArea(row)
  const splitSubLocation = getCustomerSubLocationParts(row)
  if (levelIndex === 2) return splitSubLocation[0] ?? ''
  const levelNumber = levelIndex + 1
  const candidate =
    dynamicRow[`Level${levelNumber}`] ??
    dynamicRow[`LocationLevel${levelNumber}`] ??
    dynamicRow[`CustomerLevel${levelNumber}`] ??
    dynamicRow[`field_${levelIndex}`]
  return readString(candidate) || splitSubLocation[levelIndex - 2] || ''
}

const getFormCustomerLevelValue = (form: Form, levelIndex: number) =>
  String(form[CUSTOMER_LEVELS[levelIndex].key] ?? '').trim()

const rowMatchesCustomerPath = (row: CustomerAreaOption, form: Form, upToLevelIndex: number) => {
  for (let i = 0; i < upToLevelIndex; i += 1) {
    const selected = getFormCustomerLevelValue(form, i)
    if (!selected || getCustomerLevelValue(row, i) !== selected) return false
  }
  return true
}

const getCustomerLevelOptions = (rows: CustomerAreaOption[], form: Form, levelIndex: number) => {
  if (levelIndex > 0) {
    for (let i = 0; i < levelIndex; i += 1) {
      if (!getFormCustomerLevelValue(form, i)) return []
    }
  }
  return uniqueValues(
    rows
      .filter((row) => rowMatchesCustomerPath(row, form, levelIndex))
      .map((row) => getCustomerLevelValue(row, levelIndex)),
  ).sort((a, b) => a.localeCompare(b))
}

const pad = (value: number) => String(value).padStart(2, '0')

const formatDateInputValue = (date: Date) =>
  `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`

const generateOrderTitle = () => {
  const now = new Date()
  const randomSuffix = Math.random().toString(36).slice(2, 6).toUpperCase()
  return `SO-${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}-${String(now.getMilliseconds()).padStart(3, '0')}-${randomSuffix}`
}

const createWasteLine = (): WasteLine => ({
  id: `waste-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
  WasteCategory: '',
  Tonnage: '',
  scaleOcrStatus: 'idle',
})

const createBlankForm = (): Form => ({
  Title: generateOrderTitle(), Customer: '', CustomerName: '', CustomerLocation: '',
  CustomerTenant: '', CustomerLevel4: '', CustomerLevel5: '', CustomerLevel6: '', CustomerLevel7: '',
  CustomerLevel8: '', CustomerLevel9: '', CustomerLevel10: '', IsAdhocCustomer: false, DriverName: '',
  VehicleNumber: '', DateOfCollection: formatDateInputValue(new Date()), WasteItems: [createWasteLine()],
  Notes: '',
}
)

const hasWasteLineValue = (line: WasteLine) =>
  Boolean(line.WasteCategory.trim() || line.Tonnage.trim())

const isWasteLineComplete = (line: WasteLine) => {
  const tonnage = Number(line.Tonnage)
  return Boolean(line.WasteCategory.trim()) && line.Tonnage.trim() !== '' && Number.isFinite(tonnage) && tonnage >= 0
}

const getCompleteWasteLines = (lines: WasteLine[]) => lines.filter(isWasteLineComplete)

const hasIncompleteWasteLine = (lines: WasteLine[]) =>
  lines.some((line) => hasWasteLineValue(line) && !isWasteLineComplete(line))

const revokeScalePhotoPreviewUrls = (lines: WasteLine[]) => {
  for (const line of lines) {
    if (line.scalePhotoPreviewUrl) URL.revokeObjectURL(line.scalePhotoPreviewUrl)
  }
}

const STEPS = [
  { label: 'Basic Info', desc: 'Title & date', icon: I.fileTxt },
  { label: 'Customer', desc: 'Customer details', icon: I.building },
  { label: 'Assignment', desc: 'Driver & vehicle', icon: I.truck },
  { label: 'Proof', desc: 'Signature & photos', icon: I.clipList },
  { label: 'Review', desc: 'Confirm & submit', icon: I.check },
] as const

const LAST_STEP = STEPS.length - 1

const uploadServiceOrderProof = async (request: ProofUploadRequest): Promise<ProofUploadResponse> => {
  const queueItem: Partial<Omit<ServiceOrderProofQueueWrite, 'ID'>> = {
    Title: `${request.orderTitle}-proof`,
    ServiceOrderId: String(request.serviceOrderId),
    OrderTitle: request.orderTitle,
    SignatureFileName: request.signatureFileName,
    SignatureBase64: request.signatureBase64,
    Processed: false,
  }

  if (request.beforePhotoFileName && request.beforePhotoBase64) {
    queueItem.BeforePhotoFileName = request.beforePhotoFileName
    queueItem.BeforePhotoBase64 = request.beforePhotoBase64
  }
  if (request.afterPhotoFileName && request.afterPhotoBase64) {
    queueItem.AfterPhotoFileName = request.afterPhotoFileName
    queueItem.AfterPhotoBase64 = request.afterPhotoBase64
  }

  const result = await ServiceOrderProofQueueService.create(queueItem as Omit<ServiceOrderProofQueueWrite, 'ID'>)

  const confirmedQueueItem = result.data?.ID
    ? result.data
    : await findProofQueueByTitle(queueItem.Title ?? '')
  if (!confirmedQueueItem?.ID) {
    throw new Error('Could not confirm the proof media queue item.')
  }

  return { queueItem: confirmedQueueItem, queued: true }
}

const dataUrlToBase64 = (dataUrl: string) => dataUrl.split(',')[1] ?? ''

const loadImage = (file: File) =>
  new Promise<HTMLImageElement>((resolve, reject) => {
    const url = URL.createObjectURL(file)
    const image = new Image()
    image.onload = () => {
      URL.revokeObjectURL(url)
      resolve(image)
    }
    image.onerror = () => {
      URL.revokeObjectURL(url)
      reject(new Error(`Could not load ${file.name}.`))
    }
    image.src = url
  })

const fileToBase64 = async (file: File) => {
  const image = await loadImage(file)
  const canvas = document.createElement('canvas')
  const maxSide = 720
  const scale = Math.min(1, maxSide / Math.max(image.naturalWidth, image.naturalHeight))
  canvas.width = Math.max(1, Math.round(image.naturalWidth * scale))
  canvas.height = Math.max(1, Math.round(image.naturalHeight * scale))

  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error('Could not prepare the photo for upload.')
  ctx.drawImage(image, 0, 0, canvas.width, canvas.height)

  for (const quality of [0.72, 0.62, 0.52, 0.42, 0.32]) {
    const base64 = dataUrlToBase64(canvas.toDataURL('image/jpeg', quality))
    if (base64.length <= 60000) return base64
  }

  throw new Error(`${file.name} is too large to queue. Try a smaller photo.`)
}

const getPhotoFileName = (orderTitle: string, label: 'before' | 'after', file: File) =>
  `${orderTitle}-${label}-${file.name.replace(/[^a-z0-9]/gi, '-').slice(0, 24)}.jpg`

const delay = (ms: number) => new Promise((resolve) => window.setTimeout(resolve, ms))

const escapeODataString = (value: string) => value.replace(/'/g, "''")

const findServiceOrdersByTitle = async (title: string, expectedCount = 1) => {
  const filter = `Title eq '${escapeODataString(title)}'`
  for (let attempt = 0; attempt < 5; attempt += 1) {
    const result = await ServiceOrdersService.getAll({ filter, top: Math.max(expectedCount, 50) })
    const rows = (result.data ?? []).filter((order) => order.Title === title)
    if (rows.length >= expectedCount || attempt === 4) {
      return rows.sort((a, b) => Number(a.ID ?? 0) - Number(b.ID ?? 0))
    }
    await delay(400)
  }
  return []
}

const findProofQueueByTitle = async (title: string) => {
  if (!title) return undefined
  const filter = `Title eq '${escapeODataString(title)}'`
  for (let attempt = 0; attempt < 5; attempt += 1) {
    const result = await ServiceOrderProofQueueService.getAll({ filter, top: 1 })
    const row = (result.data ?? []).find((item) => item.Title === title)
    if (row?.ID || attempt === 4) return row
    await delay(400)
  }
  return undefined
}

/* ═══════════════════════════════════════════════════════
   Sub-components
   ═══════════════════════════════════════════════════════ */

function CustomSelect({
  id, value, onChange, disabled, options, placeholder, searchable = false, searchPlaceholder = 'Search...',
}: {
  id: string, value: string, onChange: (val: string) => void, disabled?: boolean,
  options: { value: string, label: string }[], placeholder?: string, searchable?: boolean, searchPlaceholder?: string
}) {
  const [isOpen, setIsOpen] = useState(false)
  const [searchTerm, setSearchTerm] = useState('')
  const containerRef = useRef<HTMLDivElement>(null)
  const searchInputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false)
        setSearchTerm('')
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [])

  useEffect(() => {
    if (searchable) window.setTimeout(() => searchInputRef.current?.focus(), 0)
  }, [isOpen, searchable])

  const selectedOption = options.find(o => o.value === value)
  const normalizedSearch = searchTerm.trim().toLowerCase()
  const filteredOptions = normalizedSearch
    ? options.filter((option) => option.label.toLowerCase().includes(normalizedSearch))
    : options

  return (
    <div className={`cora-custom-select ${disabled ? 'disabled' : ''} ${isOpen ? 'open' : ''}`} ref={containerRef}>
      <button 
        type="button" 
        id={id}
        className={`cora-custom-select-trigger ${!selectedOption && placeholder ? 'placeholder' : ''}`}
        onClick={() => {
          if (disabled) return
          setIsOpen((current) => {
            if (current) setSearchTerm('')
            return !current
          })
        }}
        aria-haspopup="listbox"
        aria-expanded={isOpen}
      >
        <span className="cora-custom-select-value">
          {selectedOption ? selectedOption.label : (placeholder || 'Select...')}
        </span>
        <svg className="cora-custom-select-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <polyline points="6 9 12 15 18 9"/>
        </svg>
      </button>
      {isOpen && (
        <ul className="cora-custom-select-menu" role="listbox">
          {searchable && (
            <li className="cora-custom-select-search-wrap">
              <input
                ref={searchInputRef}
                className="cora-custom-select-search"
                value={searchTerm}
                onChange={(event) => setSearchTerm(event.target.value)}
                onKeyDown={(event) => event.stopPropagation()}
                placeholder={searchPlaceholder}
                aria-label={searchPlaceholder}
              />
            </li>
          )}
          {filteredOptions.length === 0 && <li className="cora-custom-select-empty">No matching options</li>}
          {filteredOptions.map(opt => (
            <li 
              key={opt.value} 
              role="option" 
              aria-selected={value === opt.value}
              className={`cora-custom-select-option ${value === opt.value ? 'selected' : ''}`}
              onClick={() => { onChange(opt.value); setIsOpen(false); setSearchTerm('') }}
            >
              {opt.label}
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}

function CustomDatePicker({ id, value, onChange, disabled }: { id: string, value: string, onChange: (val: string) => void, disabled?: boolean }) {
  const [isOpen, setIsOpen] = useState(false)
  const [viewDate, setViewDate] = useState(() => value ? new Date(value) : new Date())
  const containerRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [])

  const currentYear = viewDate.getFullYear()
  const currentMonth = viewDate.getMonth()

  const startOfMonth = new Date(currentYear, currentMonth, 1)
  const endOfMonth = new Date(currentYear, currentMonth + 1, 0)
  
  const startDayOfWeek = startOfMonth.getDay()
  const daysInMonth = endOfMonth.getDate()

  const days = []
  for (let i = 0; i < startDayOfWeek; i++) {
    days.push(null)
  }
  for (let i = 1; i <= daysInMonth; i++) {
    days.push(new Date(currentYear, currentMonth, i))
  }

  const handlePrevMonth = () => setViewDate(new Date(currentYear, currentMonth - 1, 1))
  const handleNextMonth = () => setViewDate(new Date(currentYear, currentMonth + 1, 1))

  const handleSelect = (d: Date) => {
    const y = d.getFullYear()
    const m = String(d.getMonth() + 1).padStart(2, '0')
    const day = String(d.getDate()).padStart(2, '0')
    onChange(`${y}-${m}-${day}`)
    setIsOpen(false)
  }

  const formatDisplay = (val: string) => {
    if (!val) return ''
    const [y, m, d] = val.split('-')
    if (!y || !m || !d) return ''
    return `${d}/${m}/${y}`
  }

  return (
    <div className={`cora-custom-select ${disabled ? 'disabled' : ''} ${isOpen ? 'open' : ''}`} ref={containerRef}>
      <button 
        type="button" 
        id={id}
        className={`cora-custom-select-trigger ${!value ? 'placeholder' : ''}`}
        onClick={() => !disabled && setIsOpen(!isOpen)}
        aria-haspopup="dialog"
        aria-expanded={isOpen}
      >
        <span className="cora-custom-select-value">
          {formatDisplay(value) || 'Select date...'}
        </span>
        <svg className="cora-custom-select-icon" style={{ transform: 'none' }} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
          <line x1="16" y1="2" x2="16" y2="6" />
          <line x1="8" y1="2" x2="8" y2="6" />
          <line x1="3" y1="10" x2="21" y2="10" />
        </svg>
      </button>
      {isOpen && (
        <div className="cora-date-picker-menu">
          <div className="cora-date-picker-head">
            <button type="button" onClick={handlePrevMonth} className="cora-date-picker-nav">{I.chevL}</button>
            <div className="cora-date-picker-title">
              {viewDate.toLocaleString('default', { month: 'short', year: 'numeric' })}
            </div>
            <button type="button" onClick={handleNextMonth} className="cora-date-picker-nav">{I.chevR}</button>
          </div>
          <div className="cora-date-picker-grid">
            {['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'].map(w => (
              <div key={w} className="cora-date-picker-dayname">{w}</div>
            ))}
            {days.map((d, i) => (
              d ? (
                <button
                  key={i}
                  type="button"
                  className={`cora-date-picker-day ${value === `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}` ? 'selected' : ''}`}
                  onClick={() => handleSelect(d)}
                >
                  {d.getDate()}
                </button>
              ) : (
                <div key={i} />
              )
            ))}
          </div>
        </div>
      )}
    </div>
  )
}

function Sidebar({
  view, onNav, onAdminNav, open, collapsed, onClose, isAdmin,
}: {
  view: View
  onNav: (v: View) => void
  onAdminNav: (tab: AdminTab) => void
  open: boolean
  collapsed: boolean
  onClose: () => void
  isAdmin: boolean
}) {
  return (
    <>
      <div className={`cora-sidebar-backdrop ${open ? 'visible' : ''}`} onClick={onClose} />
      <aside
        className={`cora-sidebar ${open ? 'open' : ''} ${collapsed ? 'collapsed' : ''}`}
      >
        <div className="cora-sidebar-brand">
          <img src={logoImage} alt="CORA Environment" className="cora-sidebar-brand-img" />
          {/* Close button — visible only on mobile when sidebar is open */}
          <button className="cora-sidebar-close" onClick={onClose} aria-label="Close menu">
            {I.x}
          </button>
        </div>

        <div className="cora-sidebar-section">
          <div className="cora-sidebar-section-title">Operations</div>
          <button className={`cora-nav-item ${view === 'form' ? 'active' : ''}`} onClick={() => { onNav('form'); onClose() }} id="nav-new-order">
            {I.plus} New Order
          </button>
          <button
            className={`cora-nav-item ${view === 'list' ? 'active' : ''}`}
            onClick={() => { onNav('list'); onClose() }}
            disabled={!isAdmin}
            id="nav-orders"
          >
            {I.clipList} All Orders
          </button>
        </div>

        <div className="cora-sidebar-section">
          <div className="cora-sidebar-section-title">Reference</div>
          <button
            className={`cora-nav-item ${view === 'admin' ? 'active' : ''}`}
            onClick={() => { onAdminNav('drivers'); onClose() }}
            disabled={!isAdmin}
          >
            {I.user} Drivers
          </button>
          <button
            className={`cora-nav-item ${view === 'admin' ? 'active' : ''}`}
            onClick={() => { onAdminNav('vehicles'); onClose() }}
            disabled={!isAdmin}
          >
            {I.truck} Vehicles
          </button>
          <button
            className={`cora-nav-item ${view === 'admin' ? 'active' : ''}`}
            onClick={() => { onAdminNav('waste'); onClose() }}
            disabled={!isAdmin}
          >
            {I.recycle} Waste Categories
          </button>
        </div>

        <div className="cora-sidebar-footer">v1.0 · CORA Environment</div>
      </aside>
    </>
  )
}

/* ── Toast ── */
function ToastNotif({ toast, onClose }: { toast: Toast; onClose: () => void }) {
  useEffect(() => { const t = setTimeout(onClose, 4200); return () => clearTimeout(t) }, [onClose])
  return (
    <div className="cora-toast-wrap">
      <div className={`cora-toast ${toast.type}`} role="alert">
        <span>{toast.type === 'success' ? '✓' : '✕'}</span>
        <span>{toast.text}</span>
        <button className="cora-toast-close" onClick={onClose} aria-label="Dismiss">×</button>
      </div>
    </div>
  )
}

/* ── Orders Table ── */
function OrdersTable({ orders, loadState }: { orders: ServiceOrdersRead[]; loadState: Load }) {
  const [search, setSearch] = useState('')

  const cols = ['ID', 'Title', 'CustomerName', 'DriverName', 'VehicleNumber', 'DateOfCollection'] as const
  const colLabels: Record<(typeof cols)[number], string> = {
    ID: 'ID',
    Title: 'Title',
    CustomerName: 'Customer',
    DriverName: 'Driver',
    VehicleNumber: 'Vehicle',
    DateOfCollection: 'Collection Date',
  }

  const normalizedSearch = search.trim().toLowerCase()
  const filtered = normalizedSearch
    ? orders.filter((o) =>
        cols.some((c) => String(o[c] ?? '').toLowerCase().includes(normalizedSearch))
      )
    : orders

  const formatCell = (col: (typeof cols)[number], value: unknown) => {
    const str = String(value ?? '—')
    if (col === 'ID') return <span className="cora-table-id">{str}</span>
    if (col === 'DateOfCollection' && str !== '—') {
      const parts = str.split('-')
      if (parts.length === 3) return `${parts[2]}/${parts[1]}/${parts[0]}`
    }
    return str
  }

  return (
    <div className="cora-card cora-orders-card">
      <div className="cora-card-head">
        <span className="cora-card-title" style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <span style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: 30, height: 30, borderRadius: '50%', background: 'var(--cora-green-muted)', color: 'var(--cora-green-hover)' }}>{I.clipList}</span>
          All Service Orders
        </span>
        <span className="cora-card-badge">{orders.length} records</span>
      </div>

      {loadState === 'loaded' && orders.length > 0 && (
        <div className="cora-table-toolbar">
          <div className="cora-table-search-wrap">
            <svg className="cora-table-search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" />
            </svg>
            <input
              className="cora-table-search"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Search orders..."
              aria-label="Search orders"
            />
            {search && (
              <button className="cora-table-search-clear" onClick={() => setSearch('')} aria-label="Clear search">
                {I.x}
              </button>
            )}
          </div>
          {normalizedSearch && (
            <span className="cora-table-filter-badge">
              {filtered.length} of {orders.length} shown
            </span>
          )}
        </div>
      )}

      {loadState === 'loading' && (
        <div className="cora-card-body">
          {[1, 2, 3, 4, 5].map((i) => <div key={i} className="cora-skeleton" style={{ marginBottom: 10, height: i === 1 ? 32 : 44, opacity: 1 - i * 0.12 }} />)}
        </div>
      )}
      {loadState === 'error' && (
        <div className="cora-table-empty cora-table-empty-error">
          <span className="cora-table-empty-icon" style={{ background: 'var(--error-bg)', color: 'var(--error)' }}>{I.x}</span>
          <span>Failed to load orders. Please try again.</span>
        </div>
      )}
      {loadState === 'loaded' && orders.length === 0 && (
        <div className="cora-table-empty">
          <span className="cora-table-empty-icon">{I.clipList}</span>
          <span>No service orders yet. Create one to get started.</span>
        </div>
      )}
      {loadState === 'loaded' && orders.length > 0 && filtered.length === 0 && (
        <div className="cora-table-empty">
          <span className="cora-table-empty-icon">{I.clipList}</span>
          <span>No orders match "{search}"</span>
        </div>
      )}
      {loadState === 'loaded' && filtered.length > 0 && (
        <div className="cora-table-wrap">
          <table>
            <thead><tr>{cols.map((c) => <th key={c}>{colLabels[c]}</th>)}</tr></thead>
            <tbody>
              {filtered.map((o, i) => (
                <tr key={o.ID ?? i} style={{ animationDelay: `${Math.min(i * 30, 300)}ms` }}>
                  {cols.map((c) => (
                    <td key={c} data-label={colLabels[c]}>{formatCell(c, o[c])}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}

function AdminReferencePage({
  activeTab, onTabChange, drivers, vehicles, categories, busy, currentUserEmail,
  onAddDriver, onAddVehicle, onAddWasteCategory,
  onUpdateDriver, onUpdateVehicle, onUpdateWasteCategory,
  onDeleteDriver, onDeleteVehicle, onDeleteWasteCategory,
}: {
  activeTab: AdminTab
  onTabChange: (tab: AdminTab) => void
  drivers: Drivers1Read[]
  vehicles: VehiclesRead[]
  categories: Waste_CategoriesRead[]
  busy: boolean
  currentUserEmail: string
  onAddDriver: (name: string) => Promise<void>
  onAddVehicle: (vehicleNumber: string) => Promise<void>
  onAddWasteCategory: (category: string) => Promise<void>
  onUpdateDriver: (id: number, name: string) => Promise<void>
  onUpdateVehicle: (id: number, vehicleNumber: string) => Promise<void>
  onUpdateWasteCategory: (id: number, category: string) => Promise<void>
  onDeleteDriver: (id: number, name: string) => Promise<void>
  onDeleteVehicle: (id: number, vehicleNumber: string) => Promise<void>
  onDeleteWasteCategory: (id: number, category: string) => Promise<void>
}) {
  const [driverName, setDriverName] = useState('')
  const [vehicleNumber, setVehicleNumber] = useState('')
  const [wasteCategory, setWasteCategory] = useState('')
  const [saving, setSaving] = useState<AdminTab | null>(null)
  const [editing, setEditing] = useState<{ tab: AdminTab; id: number; value: string } | null>(null)
  const [busyItem, setBusyItem] = useState<string | null>(null)

  const submit = async (tab: AdminTab) => {
    const values = {
      drivers: driverName,
      vehicles: vehicleNumber,
      waste: wasteCategory,
    }
    const value = values[tab].trim()
    if (!value) return

    setSaving(tab)
    try {
      if (tab === 'drivers') {
        await onAddDriver(value)
        setDriverName('')
      } else if (tab === 'vehicles') {
        await onAddVehicle(value)
        setVehicleNumber('')
      } else {
        await onAddWasteCategory(value)
        setWasteCategory('')
      }
    } finally {
      setSaving(null)
    }
  }

  const tabs: { key: AdminTab; label: string; count: number; icon: ReactNode }[] = [
    { key: 'drivers', label: 'Drivers', count: drivers.length, icon: I.user },
    { key: 'vehicles', label: 'Vehicles', count: vehicles.length, icon: I.truck },
    { key: 'waste', label: 'Waste Categories', count: categories.length, icon: I.recycle },
  ]
  const activeItems = activeTab === 'drivers'
    ? drivers.map((driver) => ({ id: driver.ID, value: getDriverName(driver) }))
    : activeTab === 'vehicles'
      ? vehicles.map((vehicle) => ({ id: vehicle.ID, value: vehicle.Title?.trim() || '' }))
      : categories.map((category) => ({ id: category.ID, value: category.Title?.trim() || '' }))

  const activeConfig = {
    drivers: {
      title: 'Add Driver',
      label: 'Driver Name',
      placeholder: 'Enter driver name',
      value: driverName,
      onChange: setDriverName,
    },
    vehicles: {
      title: 'Add Vehicle',
      label: 'Vehicle Number',
      placeholder: 'Enter vehicle number',
      value: vehicleNumber,
      onChange: setVehicleNumber,
    },
    waste: {
      title: 'Add Waste Category',
      label: 'Waste Category',
      placeholder: 'Enter waste category',
      value: wasteCategory,
      onChange: setWasteCategory,
    },
  }[activeTab]

  const updateItem = async (id: number, value: string) => {
    const cleaned = value.trim()
    if (!cleaned) return
    const key = `update-${activeTab}-${id}`
    setBusyItem(key)
    try {
      if (activeTab === 'drivers') await onUpdateDriver(id, cleaned)
      else if (activeTab === 'vehicles') await onUpdateVehicle(id, cleaned)
      else await onUpdateWasteCategory(id, cleaned)
      setEditing(null)
    } finally {
      setBusyItem(null)
    }
  }

  const deleteItem = async (id: number, value: string) => {
    if (!window.confirm(`Delete "${value}"? This cannot be undone.`)) return
    const key = `delete-${activeTab}-${id}`
    setBusyItem(key)
    try {
      if (activeTab === 'drivers') await onDeleteDriver(id, value)
      else if (activeTab === 'vehicles') await onDeleteVehicle(id, value)
      else await onDeleteWasteCategory(id, value)
      if (editing?.id === id && editing.tab === activeTab) setEditing(null)
    } finally {
      setBusyItem(null)
    }
  }

  return (
    <div className="cora-card">
      <div className="cora-card-head">
        <span className="cora-card-title" style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <span style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: 30, height: 30, borderRadius: '50%', background: 'var(--cora-green-muted)', color: 'var(--cora-green-hover)' }}>{I.gear}</span>
          Admin Reference Data
        </span>
        <span className="cora-card-badge">{currentUserEmail || 'Admin'}</span>
      </div>
      <div className="cora-card-body">
        <div className="cora-admin-tabs" role="tablist" aria-label="Admin reference sections">
          {tabs.map((tab) => (
            <button
              key={tab.key}
              type="button"
              className={`cora-admin-tab ${activeTab === tab.key ? 'active' : ''}`}
              onClick={() => {
                setEditing(null)
                onTabChange(tab.key)
              }}
              disabled={busy || Boolean(saving) || Boolean(busyItem)}
            >
              {tab.icon} {tab.label} <span>{tab.count}</span>
            </button>
          ))}
        </div>

        <div className="cora-form-grid cora-admin-form">
          <div className="cora-section-divider"><span>{activeConfig.title}</span></div>
          <div className="cora-field">
            <label className="cora-label" htmlFor={`admin-${activeTab}`}>{activeConfig.label}</label>
            <input
              id={`admin-${activeTab}`}
              className="cora-input"
              value={activeConfig.value}
              onChange={(event) => activeConfig.onChange(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') {
                  event.preventDefault()
                  void submit(activeTab)
                }
              }}
              placeholder={activeConfig.placeholder}
              disabled={busy || Boolean(saving)}
            />
          </div>
          <div className="cora-field cora-admin-submit-field">
            <label className="cora-label" aria-hidden="true">&nbsp;</label>
            <button
              type="button"
              className="cora-btn cora-btn-primary"
              onClick={() => void submit(activeTab)}
              disabled={busy || Boolean(saving) || !activeConfig.value.trim()}
            >
              {saving === activeTab ? <><span className="cora-spinner" /> Saving</> : <>{I.plus} Add</>}
            </button>
          </div>
        </div>

        <div className="cora-admin-list">
          <div className="cora-section-divider"><span>Current {tabs.find((tab) => tab.key === activeTab)?.label}</span></div>
          {activeItems.filter((item) => item.id && item.value).length === 0 ? (
            <div className="cora-table-empty">No {tabs.find((tab) => tab.key === activeTab)?.label.toLowerCase()} have been added yet. Use the form above to add one.</div>
          ) : (
            <div className="cora-admin-crud-list">
              {activeItems.filter((item): item is { id: number; value: string } => Boolean(item.id && item.value)).map((item) => (
                <div className="cora-admin-crud-item" key={`${activeTab}-${item.id}`}>
                  {editing?.tab === activeTab && editing.id === item.id ? (
                    <input
                      className="cora-input cora-admin-edit-input"
                      value={editing.value}
                      onChange={(event) => setEditing({ ...editing, value: event.target.value })}
                      onKeyDown={(event) => {
                        if (event.key === 'Enter') {
                          event.preventDefault()
                          void updateItem(item.id, editing.value)
                        }
                        if (event.key === 'Escape') setEditing(null)
                      }}
                      disabled={Boolean(busyItem)}
                      autoFocus
                    />
                  ) : (
                    <span className="cora-admin-crud-name">{item.value}</span>
                  )}
                  <div className="cora-admin-crud-actions">
                    {editing?.tab === activeTab && editing.id === item.id ? (
                      <>
                        <button
                          type="button"
                          className="cora-btn cora-btn-primary"
                          onClick={() => void updateItem(item.id, editing.value)}
                          disabled={Boolean(busyItem) || !editing.value.trim()}
                        >
                          {busyItem === `update-${activeTab}-${item.id}` ? <><span className="cora-spinner" /> Saving</> : 'Save'}
                        </button>
                        <button type="button" className="cora-btn cora-btn-outline" onClick={() => setEditing(null)} disabled={Boolean(busyItem)}>
                          Cancel
                        </button>
                      </>
                    ) : (
                      <>
                        <button
                          type="button"
                          className="cora-btn cora-btn-outline"
                          onClick={() => setEditing({ tab: activeTab, id: item.id, value: item.value })}
                          disabled={Boolean(busyItem)}
                        >
                          Edit
                        </button>
                        <button
                          type="button"
                          className="cora-btn cora-btn-danger"
                          onClick={() => void deleteItem(item.id, item.value)}
                          disabled={Boolean(busyItem)}
                        >
                          {busyItem === `delete-${activeTab}-${item.id}` ? <><span className="cora-spinner" /> Deleting</> : 'Delete'}
                        </button>
                      </>
                    )}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

/* ═══════════════════════════════════════════════════════
   Wizard Step Content
   ═══════════════════════════════════════════════════════ */

function StepBasic({ form, set, busy }: { form: Form; set: <K extends keyof Form>(k: K, v: Form[K]) => void; busy: boolean }) {
  return (
    <div className="cora-form-grid" key="step-basic">
      <div className="cora-section-divider"><span>Order Identification</span></div>
      <div className="cora-field">
        <label className="cora-label" htmlFor="f-title">Title <span className="req">*</span></label>
        <input id="f-title" className="cora-input" value={form.Title} readOnly disabled={busy} />
        <p className="cora-field-hint">This order title is auto-generated and read-only.</p>
      </div>
      <div className="cora-field">
        <label className="cora-label" htmlFor="f-date">Collection Date</label>
        <CustomDatePicker
          id="f-date"
          value={form.DateOfCollection}
          onChange={(val) => set('DateOfCollection', val)}
          disabled={busy}
        />
      </div>
    </div>
  )
}

function StepCustomer({
  form, set, busy, customerAreas, onOpenAdhocModal,
}: {
  form: Form
  set: <K extends keyof Form>(k: K, v: Form[K]) => void
  busy: boolean
  customerAreas: CustomerAreaOption[]
  onOpenAdhocModal: () => void
}) {
  const levelOptions = CUSTOMER_LEVELS.map((_, levelIndex) => getCustomerLevelOptions(customerAreas, form, levelIndex))
  const visibleLevels = CUSTOMER_LEVELS.filter((_, levelIndex) => levelIndex === 0 || levelOptions[levelIndex].length > 0)

  const chooseLevel = (levelIndex: number, value: string) => {
    const level = CUSTOMER_LEVELS[levelIndex]
    set(level.key, value)
    if (levelIndex === 0) {
      set('IsAdhocCustomer', false)
      set('Customer', value)
    }
    CUSTOMER_LEVELS.slice(levelIndex + 1).forEach((childLevel) => {
      set(childLevel.key, '')
    })
  }

  return (
    <div className="cora-form-grid" key="step-customer">
      <div className="cora-section-divider"><span>Customer Details</span></div>
      {visibleLevels.map((level) => {
        const levelIndex = CUSTOMER_LEVELS.indexOf(level)
        const options = levelOptions[levelIndex]
        const selectedValue = options.includes(form[level.key]) ? form[level.key] : ''
        return (
          <div className="cora-field" key={level.key}>
            <label className="cora-label" htmlFor={`f-customer-level-${levelIndex + 1}`}>{level.label}</label>
            <CustomSelect
              id={`f-customer-level-${levelIndex + 1}`}
              value={selectedValue}
              onChange={(val) => chooseLevel(levelIndex, val)}
              disabled={busy}
              placeholder={level.placeholder}
              searchable
              searchPlaceholder={level.searchPlaceholder}
              options={options.map((option) => ({ value: option, label: option }))}
            />
          </div>
        )
      })}
      <div className="cora-field">
        <label className="cora-label">Ad-hoc Customer</label>
        <div className="cora-switch-row">
          <div
            className={`cora-switch ${form.IsAdhocCustomer ? 'on' : ''}`}
            role="switch" aria-checked={form.IsAdhocCustomer} tabIndex={0}
            onClick={() => {
              if (busy) return
              if (form.IsAdhocCustomer) {
                set('IsAdhocCustomer', false)
                return
              }
              onOpenAdhocModal()
            }}
            onKeyDown={(e) => {
              if (busy || (e.key !== ' ' && e.key !== 'Enter')) return
              e.preventDefault()
              if (form.IsAdhocCustomer) {
                set('IsAdhocCustomer', false)
                return
              }
              onOpenAdhocModal()
            }}
            id="switch-adhoc"
          />
          <span className="cora-switch-label">{form.IsAdhocCustomer ? 'Yes' : 'No'}</span>
        </div>
      </div>
    </div>
  )
}

function AdhocCustomerModal({
  initialValues, onCancel, onSave, busy,
}: {
  initialValues: Pick<Form, 'CustomerName' | 'CustomerLocation' | 'CustomerTenant'>
  onCancel: () => void
  onSave: (values: Pick<Form, 'CustomerName' | 'CustomerLocation' | 'CustomerTenant'>) => void
  busy: boolean
}) {
  const [values, setValues] = useState(initialValues)
  const customerName = values.CustomerName.trim()
  const canSave = Boolean(customerName)

  useEffect(() => {
    setValues(initialValues)
  }, [initialValues])

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') onCancel()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onCancel])

  const update = (field: keyof typeof values, value: string) => {
    setValues((current) => ({ ...current, [field]: value }))
  }

  const save = () => {
    if (!canSave) return
    onSave({
      CustomerName: customerName,
      CustomerLocation: values.CustomerLocation.trim(),
      CustomerTenant: values.CustomerTenant.trim(),
    })
  }

  return (
    <div className="cora-modal-backdrop" role="presentation" onMouseDown={onCancel}>
      <div className="cora-modal" role="dialog" aria-modal="true" aria-labelledby="adhoc-customer-title" onMouseDown={(event) => event.stopPropagation()}>
        <div className="cora-modal-head">
          <div>
            <h2 id="adhoc-customer-title">Ad-hoc Customer Setup</h2>
            <p>Enter the custom customer information for this special order.</p>
          </div>
          <button className="cora-modal-close" type="button" onClick={onCancel} aria-label="Close ad-hoc customer details">
            {I.x}
          </button>
        </div>
        <div className="cora-modal-body">
          <div className="cora-floating-field">
            <input
              id="adhoc-customer-name"
              className="cora-floating-input"
              placeholder=" "
              value={values.CustomerName}
              onChange={(event) => update('CustomerName', event.target.value)}
              autoFocus
              disabled={busy}
            />
            <label className="cora-floating-label" htmlFor="adhoc-customer-name">Customer Name *</label>
          </div>
          <div className="cora-floating-field">
            <input
              id="adhoc-customer-location"
              className="cora-floating-input"
              placeholder=" "
              value={values.CustomerLocation}
              onChange={(event) => update('CustomerLocation', event.target.value)}
              disabled={busy}
            />
            <label className="cora-floating-label" htmlFor="adhoc-customer-location">Location / Building</label>
          </div>
          <div className="cora-floating-field">
            <input
              id="adhoc-customer-tenant"
              className="cora-floating-input"
              placeholder=" "
              value={values.CustomerTenant}
              onChange={(event) => update('CustomerTenant', event.target.value)}
              disabled={busy}
            />
            <label className="cora-floating-label" htmlFor="adhoc-customer-tenant">Sub-location / Tower</label>
          </div>
        </div>
        <div className="cora-modal-foot">
          <button className="cora-btn cora-btn-outline" type="button" onClick={onCancel} disabled={busy}>Cancel</button>
          <button className="cora-btn cora-btn-primary" type="button" onClick={save} disabled={busy || !canSave}>
            Confirm Customer
          </button>
        </div>
      </div>
    </div>
  )
}

function AdminPasscodeModal({
  onCancel, onUnlock,
}: {
  onCancel: () => void
  onUnlock: () => void
}) {
  const [passcode, setPasscode] = useState('')
  const [error, setError] = useState('')

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') onCancel()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onCancel])

  const submit = () => {
    if (passcode === ADMIN_PASSCODE) {
      setError('')
      onUnlock()
      return
    }
    setError('Incorrect admin passcode.')
  }

  return (
    <div className="cora-modal-backdrop" role="presentation" onMouseDown={onCancel}>
      <div className="cora-modal" role="dialog" aria-modal="true" aria-labelledby="admin-passcode-title" onMouseDown={(event) => event.stopPropagation()}>
        <div className="cora-modal-head">
          <div>
            <h2 id="admin-passcode-title">Admin Access</h2>
            <p>Enter the admin passcode to access admin-only pages.</p>
          </div>
          <button className="cora-modal-close" type="button" onClick={onCancel} aria-label="Close admin access">
            {I.x}
          </button>
        </div>
        <div className="cora-modal-body">
          <div className="cora-floating-field">
            <input
              id="admin-passcode"
              className="cora-floating-input"
              type="password"
              placeholder=" "
              value={passcode}
              onChange={(event) => {
                setPasscode(event.target.value)
                setError('')
              }}
              onKeyDown={(event) => {
                if (event.key === 'Enter') {
                  event.preventDefault()
                  submit()
                }
              }}
              autoFocus
            />
            <label className="cora-floating-label" htmlFor="admin-passcode">Admin Passcode</label>
          </div>
          {error && <p className="cora-modal-error">{error}</p>}
        </div>
        <div className="cora-modal-foot">
          <button className="cora-btn cora-btn-outline" type="button" onClick={onCancel}>Cancel</button>
          <button className="cora-btn cora-btn-primary" type="button" onClick={submit} disabled={!passcode.trim()}>
            Unlock Admin
          </button>
        </div>
      </div>
    </div>
  )
}

function ConfirmResetModal({
  onCancel, onConfirm,
}: {
  onCancel: () => void
  onConfirm: () => void
}) {
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') onCancel()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onCancel])

  return (
    <div className="cora-modal-backdrop" role="presentation" onMouseDown={onCancel}>
      <div className="cora-modal" role="dialog" aria-modal="true" aria-labelledby="confirm-reset-title" onMouseDown={(event) => event.stopPropagation()} style={{ width: 'min(400px, 100%)' }}>
        <div className="cora-modal-head">
          <div>
            <h2 id="confirm-reset-title" style={{ color: 'var(--error)' }}>Start Over?</h2>
            <p>Are you sure you want to discard your progress and start over? This cannot be undone.</p>
          </div>
          <button className="cora-modal-close" type="button" onClick={onCancel} aria-label="Cancel">
            {I.x}
          </button>
        </div>
        <div className="cora-modal-foot">
          <button className="cora-btn cora-btn-outline" type="button" onClick={onCancel}>Cancel</button>
          <button className="cora-btn cora-btn-danger" type="button" onClick={onConfirm} autoFocus>
            Start Over
          </button>
        </div>
      </div>
    </div>
  )
}

const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(max, value))

function CropAdjustModal({
  imageUrl, initialCrop, onCancel, onApply,
}: {
  imageUrl: string
  initialCrop?: CropRect
  onCancel: () => void
  onApply: (crop: CropRect) => void
}) {
  const [crop, setCrop] = useState<CropRect>(initialCrop ?? { x: 0.2, y: 0.25, width: 0.6, height: 0.35 })
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const imageRef = useRef<HTMLImageElement | null>(null)
  const interactionRef = useRef<{
    mode: 'draw' | 'move' | 'resize'
    startX: number
    startY: number
    startCrop: CropRect
  } | null>(null)

  const drawCropCanvas = useCallback(() => {
    const canvas = canvasRef.current
    const image = imageRef.current
    const ctx = canvas?.getContext('2d')
    if (!canvas || !image || !ctx || !image.naturalWidth || !image.naturalHeight) return

    const maxWidth = Math.min(900, Math.max(320, Math.round(window.innerWidth * 0.78)))
    const maxHeight = Math.min(620, Math.max(260, Math.round(window.innerHeight * 0.58)))
    const scale = Math.min(maxWidth / image.naturalWidth, maxHeight / image.naturalHeight, 1)
    const width = Math.max(1, Math.round(image.naturalWidth * scale))
    const height = Math.max(1, Math.round(image.naturalHeight * scale))

    canvas.width = width
    canvas.height = height
    ctx.clearRect(0, 0, width, height)
    ctx.drawImage(image, 0, 0, width, height)

    const cropX = crop.x * width
    const cropY = crop.y * height
    const cropWidth = crop.width * width
    const cropHeight = crop.height * height

    ctx.fillStyle = 'rgba(0, 0, 0, 0.42)'
    ctx.fillRect(0, 0, width, cropY)
    ctx.fillRect(0, cropY + cropHeight, width, height - cropY - cropHeight)
    ctx.fillRect(0, cropY, cropX, cropHeight)
    ctx.fillRect(cropX + cropWidth, cropY, width - cropX - cropWidth, cropHeight)
    ctx.strokeStyle = '#6BBF59'
    ctx.lineWidth = 2
    ctx.strokeRect(cropX, cropY, cropWidth, cropHeight)
    ctx.fillStyle = '#6BBF59'
    ctx.beginPath()
    ctx.arc(cropX + cropWidth, cropY + cropHeight, 10, 0, Math.PI * 2)
    ctx.fill()
    ctx.strokeStyle = '#fff'
    ctx.lineWidth = 2
    ctx.stroke()
  }, [crop])

  useEffect(() => {
    const image = new Image()
    image.onload = () => {
      imageRef.current = image
      drawCropCanvas()
    }
    image.src = imageUrl
  }, [drawCropCanvas, imageUrl])

  useEffect(() => {
    drawCropCanvas()
  }, [drawCropCanvas])

  useEffect(() => {
    window.addEventListener('resize', drawCropCanvas)
    return () => window.removeEventListener('resize', drawCropCanvas)
  }, [drawCropCanvas])

  const pointerToFrame = (event: PointerEvent<HTMLElement>) => {
    const canvas = canvasRef.current
    if (!canvas) return null
    const rect = canvas.getBoundingClientRect()
    return {
      x: (event.clientX - rect.left) / rect.width,
      y: (event.clientY - rect.top) / rect.height,
    }
  }

  const startInteraction = (event: PointerEvent<HTMLElement>, mode: 'move' | 'resize') => {
    event.preventDefault()
    event.currentTarget.setPointerCapture(event.pointerId)
    const point = pointerToFrame(event)
    if (!point) return
    interactionRef.current = {
      mode,
      startX: point.x,
      startY: point.y,
      startCrop: crop,
    }
  }

  const startDraw = (event: PointerEvent<HTMLDivElement>) => {
    if (!(event.target instanceof HTMLCanvasElement)) return
    event.preventDefault()
    event.target.setPointerCapture(event.pointerId)
    const point = pointerToFrame(event)
    if (!point) return
    const startCrop = { x: clamp(point.x, 0, 1), y: clamp(point.y, 0, 1), width: 0.01, height: 0.01 }
    setCrop(startCrop)
    interactionRef.current = {
      mode: 'draw',
      startX: startCrop.x,
      startY: startCrop.y,
      startCrop,
    }
  }

  const moveInteraction = (event: PointerEvent<HTMLDivElement>) => {
    const interaction = interactionRef.current
    const point = pointerToFrame(event)
    if (!interaction || !point) return
    event.preventDefault()
    const dx = point.x - interaction.startX
    const dy = point.y - interaction.startY

    if (interaction.mode === 'draw') {
      const x = clamp(Math.min(interaction.startX, point.x), 0, 0.99)
      const y = clamp(Math.min(interaction.startY, point.y), 0, 0.99)
      const width = clamp(Math.abs(point.x - interaction.startX), 0.08, 1 - x)
      const height = clamp(Math.abs(point.y - interaction.startY), 0.08, 1 - y)
      setCrop({ x, y, width, height })
      return
    }

    if (interaction.mode === 'move') {
      setCrop({
        ...interaction.startCrop,
        x: clamp(interaction.startCrop.x + dx, 0, 1 - interaction.startCrop.width),
        y: clamp(interaction.startCrop.y + dy, 0, 1 - interaction.startCrop.height),
      })
      return
    }

    setCrop({
      ...interaction.startCrop,
      width: clamp(interaction.startCrop.width + dx, 0.08, 1 - interaction.startCrop.x),
      height: clamp(interaction.startCrop.height + dy, 0.08, 1 - interaction.startCrop.y),
    })
  }

  const endInteraction = (event: PointerEvent<HTMLDivElement>) => {
    if (interactionRef.current && event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId)
    }
    interactionRef.current = null
  }

  return (
    <div className="cora-modal-backdrop" role="dialog" aria-modal="true" aria-labelledby="crop-adjust-title">
      <div className="cora-modal cora-crop-modal">
        <div className="cora-modal-head">
          <div>
            <h2 id="crop-adjust-title">Adjust OCR crop</h2>
            <p>Drag the box around only the scale display, then run OCR again.</p>
          </div>
          <button className="cora-modal-close" type="button" onClick={onCancel} aria-label="Close crop editor">
            {I.x}
          </button>
        </div>
        <div className="cora-modal-body">
          <div
            className="cora-crop-frame"
            onPointerDown={startDraw}
            onPointerMove={moveInteraction}
            onPointerUp={endInteraction}
            onPointerCancel={endInteraction}
          >
            <canvas ref={canvasRef} aria-label="Scale crop source" />
            <div
              className="cora-crop-box"
              style={{
                left: `${crop.x * 100}%`,
                top: `${crop.y * 100}%`,
                width: `${crop.width * 100}%`,
                height: `${crop.height * 100}%`,
              }}
              onPointerDown={(event) => {
                event.stopPropagation()
                startInteraction(event, 'move')
              }}
            >
              <span
                className="cora-crop-handle"
                onPointerDown={(event) => {
                  event.stopPropagation()
                  startInteraction(event, 'resize')
                }}
              />
            </div>
          </div>
          <p className="cora-field-hint">Drag on the image to draw a new crop. Include the digits and unit, but keep the foot, floor, and empty background out.</p>
        </div>
        <div className="cora-modal-foot">
          <button className="cora-btn cora-btn-outline" type="button" onClick={onCancel}>Cancel</button>
          <button className="cora-btn cora-btn-primary" type="button" onClick={() => onApply(crop)}>
            Use crop
          </button>
        </div>
      </div>
    </div>
  )
}

function StepAssignment({
  form, set, updateForm, busy, driverOptions, vehicles, categories,
}: {
  form: Form; set: <K extends keyof Form>(k: K, v: Form[K]) => void; busy: boolean
  updateForm: (updater: (current: Form) => Form) => void
  driverOptions: string[]; vehicles: VehiclesRead[]; categories: Waste_CategoriesRead[]
}) {
  const [cropEditingLineId, setCropEditingLineId] = useState<string | null>(null)
  const cropEditingLine = form.WasteItems.find((line) => line.id === cropEditingLineId)

  const updateWasteLine = (id: string, changes: Partial<Omit<WasteLine, 'id'>>) => {
    updateForm((current) => ({
      ...current,
      WasteItems: current.WasteItems.map((line) => (
        line.id === id ? { ...line, ...changes } : line
      )),
    }))
  }

  const addWasteLine = () => {
    set('WasteItems', [...form.WasteItems, createWasteLine()])
  }

  const removeWasteLine = (id: string) => {
    const removedLine = form.WasteItems.find((line) => line.id === id)
    if (removedLine?.scalePhotoPreviewUrl) URL.revokeObjectURL(removedLine.scalePhotoPreviewUrl)
    const nextLines = form.WasteItems.filter((line) => line.id !== id)
    set('WasteItems', nextLines.length > 0 ? nextLines : [createWasteLine()])
  }

  const clearScalePhoto = (id: string) => {
    updateForm((current) => ({
      ...current,
      WasteItems: current.WasteItems.map((line) => {
        if (line.id !== id) return line
        if (line.scalePhotoPreviewUrl) URL.revokeObjectURL(line.scalePhotoPreviewUrl)
        return {
          ...line,
          scalePhotoPreviewUrl: undefined,
          scalePhotoFile: undefined,
          scaleOcrCropPreviewUrl: undefined,
          scaleOcrCrop: undefined,
          scaleOcrStatus: 'idle',
          scaleOcrText: undefined,
          scaleOcrSuggestion: undefined,
          scaleOcrConfidence: undefined,
          scaleOcrReasons: undefined,
          scaleOcrError: undefined,
          scaleOcrRequestId: undefined,
        }
      }),
    }))
  }

  const runOcrForLine = async (id: string, file: File, requestId: string, crop?: CropRect) => {
    try {
      const preprocessed = await preprocessImageForOcr(file, crop)
      const result = await runWeightOcr(preprocessed.imageDataUrls, preprocessed.candidateTexts, {
        remoteImageDataUrl: preprocessed.cropPreviewUrl,
      })
      const suggestion = result.reliable ? result.parsed.suggestion : ''
      updateForm((current) => ({
        ...current,
        WasteItems: current.WasteItems.map((line) => {
          if (line.id !== id || line.scaleOcrRequestId !== requestId) return line
          return {
            ...line,
            scaleOcrStatus: suggestion ? 'done' : 'error',
            scaleOcrText: result.rawText,
            scaleOcrSuggestion: suggestion || undefined,
            scaleOcrConfidence: suggestion ? result.confidence : undefined,
            scaleOcrReasons: result.reasons,
            scaleOcrError: suggestion ? undefined : 'Could not confidently read weight. Enter manually.',
            scaleOcrCropPreviewUrl: preprocessed.cropPreviewUrl,
            scaleOcrCrop: preprocessed.cropRect,
          }
        }),
      }))
    } catch (error) {
      updateForm((current) => ({
        ...current,
        WasteItems: current.WasteItems.map((line) => {
          if (line.id !== id || line.scaleOcrRequestId !== requestId) return line
          return {
            ...line,
            scaleOcrStatus: 'error',
            scaleOcrConfidence: undefined,
            scaleOcrReasons: undefined,
            scaleOcrError: error instanceof Error ? error.message : 'Could not read weight. Enter manually.',
          }
        }),
      }))
    }
  }

  const handleScalePhotoChange = async (id: string, event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0] ?? null
    event.target.value = ''
    if (!file) return
    if (!file.type.startsWith('image/')) {
      updateWasteLine(id, {
        scaleOcrStatus: 'error',
        scaleOcrText: undefined,
        scaleOcrSuggestion: undefined,
        scaleOcrConfidence: undefined,
        scaleOcrReasons: undefined,
        scaleOcrError: 'Choose an image file.',
      })
      return
    }

    const requestId = `scale-ocr-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
    const previewUrl = URL.createObjectURL(file)
    updateForm((current) => ({
      ...current,
      WasteItems: current.WasteItems.map((line) => {
        if (line.id !== id) return line
        if (line.scalePhotoPreviewUrl) URL.revokeObjectURL(line.scalePhotoPreviewUrl)
        return {
          ...line,
          scalePhotoFile: file,
          scalePhotoPreviewUrl: previewUrl,
          scaleOcrStatus: 'idle',
          scaleOcrText: undefined,
          scaleOcrSuggestion: undefined,
          scaleOcrConfidence: undefined,
          scaleOcrReasons: undefined,
          scaleOcrError: undefined,
          scaleOcrCropPreviewUrl: undefined,
          scaleOcrCrop: undefined,
          scaleOcrRequestId: requestId,
        }
      }),
    }))
    setCropEditingLineId(id)
  }

  const applyScaleCrop = (id: string, crop: CropRect) => {
    const line = form.WasteItems.find((item) => item.id === id)
    if (!line?.scalePhotoFile) return
    const requestId = `scale-ocr-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
    setCropEditingLineId(null)
    updateWasteLine(id, {
      scaleOcrStatus: 'reading',
      scaleOcrText: undefined,
      scaleOcrSuggestion: undefined,
      scaleOcrConfidence: undefined,
      scaleOcrReasons: undefined,
      scaleOcrError: undefined,
      scaleOcrRequestId: requestId,
      scaleOcrCrop: crop,
    })
    void runOcrForLine(id, line.scalePhotoFile, requestId, crop)
  }

  const applyScaleSuggestion = (id: string) => {
    const line = form.WasteItems.find((item) => item.id === id)
    if (!line?.scaleOcrSuggestion) return
    updateWasteLine(id, {
      Tonnage: line.scaleOcrSuggestion,
      scaleOcrStatus: 'idle',
      scaleOcrSuggestion: undefined,
      scaleOcrConfidence: undefined,
      scaleOcrReasons: undefined,
      scaleOcrError: undefined,
    })
  }

  return (
    <div className="cora-form-grid" key="step-assign">
      {cropEditingLine?.scalePhotoPreviewUrl && (
        <CropAdjustModal
          imageUrl={cropEditingLine.scalePhotoPreviewUrl}
          initialCrop={cropEditingLine.scaleOcrCrop}
          onCancel={() => setCropEditingLineId(null)}
          onApply={(crop) => applyScaleCrop(cropEditingLine.id, crop)}
        />
      )}
      <div className="cora-section-divider"><span>Resource Assignment</span></div>
      <div className="cora-field">
        <label className="cora-label" htmlFor="f-driver">Driver</label>
        <CustomSelect
          id="f-driver"
          value={form.DriverName}
          onChange={(val) => set('DriverName', val)}
          disabled={busy}
          placeholder="Select driver…"
          options={driverOptions.map((name) => ({ value: name, label: name }))}
        />
      </div>
      <div className="cora-field">
        <label className="cora-label" htmlFor="f-vehicle">Vehicle</label>
        <CustomSelect
          id="f-vehicle"
          value={form.VehicleNumber}
          onChange={(val) => set('VehicleNumber', val)}
          disabled={busy}
          placeholder="Select vehicle…"
          options={vehicles.map((v) => ({ value: v.Title ?? '', label: v.Title ?? `Vehicle ${v.ID}` }))}
        />
      </div>

      <div className="cora-section-divider"><span>Waste Details</span></div>
      <div className="cora-field span">
        <div className="cora-waste-lines">
          {form.WasteItems.map((line, index) => (
            <div className="cora-waste-line" key={line.id}>
              <div className="cora-field">
                <label className="cora-label" htmlFor={`f-waste-category-${line.id}`}>Waste Category</label>
                <CustomSelect
                  id={`f-waste-category-${line.id}`}
                  value={line.WasteCategory}
                  onChange={(val) => updateWasteLine(line.id, { WasteCategory: val })}
                  disabled={busy}
                  placeholder="Select waste category..."
                  options={categories.map((c) => ({ value: c.Title ?? '', label: c.Title ?? `Waste category ${c.ID}` }))}
                />
              </div>
              <div className="cora-field">
                <label className="cora-label" htmlFor={`f-tonnage-${line.id}`}>Tonnage</label>
                <input
                  id={`f-tonnage-${line.id}`}
                  className="cora-input"
                  type="number"
                  min="0"
                  step="0.01"
                  placeholder="e.g. 1.5"
                  value={line.Tonnage}
                  onChange={(e) => updateWasteLine(line.id, { Tonnage: e.target.value })}
                  disabled={busy}
                />
              </div>
              <div className="cora-field cora-scale-ocr">
                <label className="cora-label" htmlFor={`f-scale-photo-${line.id}`}>Scale photo</label>
                {line.scalePhotoPreviewUrl && (
                  <img className="cora-scale-preview" src={line.scalePhotoPreviewUrl} alt={`Scale photo ${index + 1} preview`} />
                )}
                {line.scaleOcrCropPreviewUrl && (
                  <div className="cora-scale-crop-preview">
                    <span>OCR crop preview</span>
                    <img src={line.scaleOcrCropPreviewUrl} alt={`OCR crop preview ${index + 1}`} />
                  </div>
                )}
                <div className="cora-scale-actions">
                  <label className={`cora-btn cora-btn-secondary ${busy ? 'disabled' : ''}`} htmlFor={`f-scale-photo-${line.id}`}>
                    {line.scalePhotoPreviewUrl ? 'Replace scale photo' : 'Scan scale'}
                  </label>
                  {line.scalePhotoPreviewUrl && (
                    <button className="cora-btn cora-btn-outline" type="button" onClick={() => setCropEditingLineId(line.id)} disabled={busy}>
                      Adjust crop
                    </button>
                  )}
                  {line.scalePhotoPreviewUrl && (
                    <button className="cora-btn cora-btn-outline" type="button" onClick={() => clearScalePhoto(line.id)} disabled={busy}>
                      Clear
                    </button>
                  )}
                </div>
                <input
                  id={`f-scale-photo-${line.id}`}
                  className="cora-file-input"
                  type="file"
                  accept="image/*"
                  capture="environment"
                  onChange={(event) => void handleScalePhotoChange(line.id, event)}
                  disabled={busy}
                />
                {line.scaleOcrStatus === 'reading' && <p className="cora-scale-status reading">Reading scale...</p>}
                {line.scalePhotoPreviewUrl && line.scaleOcrStatus === 'idle' && (
                  <p className="cora-scale-status reading">Adjust the crop to read the scale.</p>
                )}
                {line.scaleOcrStatus === 'done' && line.scaleOcrSuggestion && (
                  <div className="cora-scale-suggestion">
                    <p className="cora-scale-status done">
                      Detected: {line.scaleOcrSuggestion}
                      {typeof line.scaleOcrConfidence === 'number' ? ` (${Math.round(line.scaleOcrConfidence)}% confidence)` : ''}
                    </p>
                    {line.scaleOcrReasons && line.scaleOcrReasons.length > 0 && (
                      <p className="cora-scale-status reading">{line.scaleOcrReasons.join('. ')}.</p>
                    )}
                    <button
                      className="cora-btn cora-btn-primary"
                      type="button"
                      onClick={() => applyScaleSuggestion(line.id)}
                      disabled={busy || line.Tonnage === line.scaleOcrSuggestion}
                    >
                      Use detected value
                    </button>
                  </div>
                )}
                {line.scaleOcrStatus === 'error' && (
                  <p className="cora-scale-status error">{line.scaleOcrError || 'Could not read weight. Enter manually.'}</p>
                )}
                {line.scaleOcrText && (
                  <details className="cora-scale-raw">
                    <summary>OCR text</summary>
                    <pre>{line.scaleOcrText}</pre>
                  </details>
                )}
              </div>
              <button
                className="cora-btn cora-btn-danger cora-waste-remove"
                type="button"
                onClick={() => removeWasteLine(line.id)}
                disabled={busy || (form.WasteItems.length === 1 && !hasWasteLineValue(line))}
                aria-label={`Remove waste item ${index + 1}`}
              >
                {I.trash} Remove
              </button>
            </div>
          ))}
          <button className="cora-btn cora-btn-secondary cora-waste-add" type="button" onClick={() => addWasteLine()} disabled={busy}>
            {I.plus} Add Waste Item
          </button>
        </div>
      </div>

      <div className="cora-section-divider"><span>Notes</span></div>
      <div className="cora-field span">
        <label className="cora-label" htmlFor="f-notes">Additional Notes</label>
        <textarea id="f-notes" className="cora-textarea" placeholder="Instructions or remarks…"
          value={form.Notes} onChange={(e) => set('Notes', e.target.value)} disabled={busy} />
      </div>

    </div>
  )
}

function SignaturePad({
  value, onChange, disabled,
}: {
  value: string
  onChange: (value: string) => void
  disabled: boolean
}) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const drawingRef = useRef(false)
  const lastPointRef = useRef<{ x: number; y: number } | null>(null)
  const hasDrawnRef = useRef(false)
  const strokeDistanceRef = useRef(0)

  const prepareCanvas = useCallback((restoreValue = value) => {
    const canvas = canvasRef.current
    if (!canvas) return

    const rect = canvas.getBoundingClientRect()
    const dpr = window.devicePixelRatio || 1
    const width = Math.max(300, Math.floor(rect.width))
    const height = 180
    canvas.width = Math.floor(width * dpr)
    canvas.height = Math.floor(height * dpr)

    const ctx = canvas.getContext('2d')
    if (!ctx) return
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, width, height)
    ctx.lineCap = 'round'
    ctx.lineJoin = 'round'
    ctx.lineWidth = 2.4
    ctx.strokeStyle = '#0e1f1d'

    if (!restoreValue) return

    const img = new Image()
    img.onload = () => ctx.drawImage(img, 0, 0, width, height)
    img.src = restoreValue
  }, [value])

  useEffect(() => {
    prepareCanvas()
    const handleResize = () => prepareCanvas()
    window.addEventListener('resize', handleResize)
    return () => window.removeEventListener('resize', handleResize)
  }, [prepareCanvas])

  const pointFromEvent = (event: PointerEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current
    if (!canvas) return null
    const rect = canvas.getBoundingClientRect()
    return { x: event.clientX - rect.left, y: event.clientY - rect.top }
  }

  const beginDrawing = (event: PointerEvent<HTMLCanvasElement>) => {
    if (disabled) return
    const point = pointFromEvent(event)
    if (!point) return
    event.currentTarget.setPointerCapture(event.pointerId)
    drawingRef.current = true
    hasDrawnRef.current = false
    strokeDistanceRef.current = 0
    lastPointRef.current = point
  }

  const draw = (event: PointerEvent<HTMLCanvasElement>) => {
    if (!drawingRef.current || disabled) return
    const canvas = canvasRef.current
    const ctx = canvas?.getContext('2d')
    const point = pointFromEvent(event)
    const lastPoint = lastPointRef.current
    if (!canvas || !ctx || !point || !lastPoint) return

    const segmentDistance = Math.hypot(point.x - lastPoint.x, point.y - lastPoint.y)
    strokeDistanceRef.current += segmentDistance
    if (strokeDistanceRef.current >= 8) hasDrawnRef.current = true

    ctx.beginPath()
    ctx.moveTo(lastPoint.x, lastPoint.y)
    ctx.lineTo(point.x, point.y)
    ctx.stroke()
    lastPointRef.current = point
  }

  const finishDrawing = () => {
    if (!drawingRef.current) return
    drawingRef.current = false
    lastPointRef.current = null
    if (hasDrawnRef.current) onChange(canvasRef.current?.toDataURL('image/png') ?? '')
    else prepareCanvas(value)
  }

  const clearSignature = () => {
    if (disabled) return
    hasDrawnRef.current = false
    strokeDistanceRef.current = 0
    prepareCanvas('')
    onChange('')
  }

  return (
    <div className="cora-signature-wrap">
      <canvas
        ref={canvasRef}
        className="cora-signature-canvas"
        aria-label="Customer signature box"
        onPointerDown={beginDrawing}
        onPointerMove={draw}
        onPointerUp={finishDrawing}
        onPointerCancel={finishDrawing}
        onPointerLeave={finishDrawing}
      />
      <div className="cora-signature-actions">
        <span>{value ? 'Signature captured.' : 'Ask the customer to sign inside the box.'}</span>
        <button className="cora-btn cora-btn-outline" type="button" onClick={clearSignature} disabled={disabled || !value}>
          Clear signature
        </button>
      </div>
    </div>
  )
}

function PhotoUpload({
  id, label, file, previewUrl, onChange, disabled,
}: {
  id: string
  label: string
  file: File | null
  previewUrl: string
  onChange: (file: File | null) => void
  disabled: boolean
}) {
  const [isDragging, setIsDragging] = useState(false)

  const chooseFile = (selected: File | null) => {
    if (selected && !selected.type.startsWith('image/')) return
    onChange(selected)
  }

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const selected = event.target.files?.[0] ?? null
    chooseFile(selected)
    event.target.value = ''
  }

  const handleDragOver = (event: DragEvent<HTMLLabelElement>) => {
    if (disabled) return
    event.preventDefault()
    setIsDragging(true)
  }

  const handleDragLeave = (event: DragEvent<HTMLLabelElement>) => {
    if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
      setIsDragging(false)
    }
  }

  const handleDrop = (event: DragEvent<HTMLLabelElement>) => {
    if (disabled) return
    event.preventDefault()
    setIsDragging(false)
    chooseFile(event.dataTransfer.files?.[0] ?? null)
  }

  return (
    <div className="cora-photo-picker">
      <label className="cora-label" htmlFor={id}>{label}</label>
      {previewUrl ? (
        <img className="cora-photo-preview" src={previewUrl} alt={`${label} preview`} />
      ) : (
        <label
          className={`cora-photo-empty ${isDragging ? 'dragging' : ''} ${disabled ? 'disabled' : ''}`}
          htmlFor={id}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
        >
          <span className="cora-photo-dropzone-icon" aria-hidden="true">
            <svg width="36" height="36" viewBox="0 0 24 24" fill="none">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </span>
          <span className="cora-photo-dropzone-label">Drop image here or choose photo</span>
          <span className="cora-photo-dropzone-hint">Camera and image files supported</span>
        </label>
      )}
      {file && <span className="cora-photo-name">{file.name}</span>}
      <div className="cora-photo-actions">
        <label className={`cora-btn cora-btn-secondary ${disabled ? 'disabled' : ''}`} htmlFor={id}>
          {file ? 'Replace photo' : 'Choose photo'}
        </label>
        {file && (
          <button className="cora-btn cora-btn-outline" type="button" onClick={() => onChange(null)} disabled={disabled}>
            Remove
          </button>
        )}
      </div>
      <input
        id={id}
        className="cora-file-input"
        type="file"
        accept="image/*"
        capture="environment"
        onChange={handleFileChange}
        disabled={disabled}
      />
    </div>
  )
}

function StepProof({
  signatureDataUrl, setSignatureDataUrl, beforePhotoFile, setBeforePhotoFile,
  afterPhotoFile, setAfterPhotoFile, beforePhotoPreviewUrl, afterPhotoPreviewUrl, busy,
}: {
  signatureDataUrl: string
  setSignatureDataUrl: (value: string) => void
  beforePhotoFile: File | null
  setBeforePhotoFile: (value: File | null) => void
  afterPhotoFile: File | null
  setAfterPhotoFile: (value: File | null) => void
  beforePhotoPreviewUrl: string
  afterPhotoPreviewUrl: string
  busy: boolean
}) {
  return (
    <div className="cora-form-grid" key="step-proof">
      <div className="cora-section-divider"><span>Customer Signature</span></div>
      <div className="cora-field span">
        <label className="cora-label">Signature <span className="req">*</span></label>
        <SignaturePad value={signatureDataUrl} onChange={setSignatureDataUrl} disabled={busy} />
      </div>

      <div className="cora-section-divider"><span>Collection Photos</span></div>
      <div className="cora-field">
        <PhotoUpload
          id="f-before-photo"
          label="Before collection photo"
          file={beforePhotoFile}
          previewUrl={beforePhotoPreviewUrl}
          onChange={setBeforePhotoFile}
          disabled={busy}
        />
      </div>
      <div className="cora-field">
        <PhotoUpload
          id="f-after-photo"
          label="After collection photo"
          file={afterPhotoFile}
          previewUrl={afterPhotoPreviewUrl}
          onChange={setAfterPhotoFile}
          disabled={busy}
        />
      </div>
    </div>
  )
}

function StepReview({
  form, signatureDataUrl, beforePhotoPreviewUrl, afterPhotoPreviewUrl,
}: {
  form: Form
  signatureDataUrl: string
  beforePhotoPreviewUrl: string
  afterPhotoPreviewUrl: string
}) {
  const completeWasteLines = getCompleteWasteLines(form.WasteItems)
  const totalTonnage = completeWasteLines.reduce((sum, line) => sum + Number(line.Tonnage), 0)
  const customerLevelRows: [string, string][] = CUSTOMER_LEVELS
    .map((level, levelIndex): [string, string] => [
      levelIndex === 0 ? 'Customer Name' : level.label,
      form[level.key] || '—',
    ])
    .filter(([, value], levelIndex) => levelIndex < 3 || value !== '—')
  const rows: [string, string][] = [
    ['Title', form.Title || '—'],
    ['Collection Date', form.DateOfCollection || '—'],
    ...customerLevelRows,
    ['Ad-hoc', form.IsAdhocCustomer ? 'Yes' : 'No'],
    ['Driver', form.DriverName || '—'],
    ['Vehicle', form.VehicleNumber || '—'],
    ['Waste Items', completeWasteLines.length > 0 ? String(completeWasteLines.length) : '—'],
    ['Total Tonnage', completeWasteLines.length > 0 ? totalTonnage.toString() : '—'],
    ['Signature', signatureDataUrl ? 'Captured' : 'Missing'],
    ['Before Photo', beforePhotoPreviewUrl ? 'Selected' : '—'],
    ['After Photo', afterPhotoPreviewUrl ? 'Selected' : '—'],
    ['Notes', form.Notes || '—'],
  ]
  return (
    <div className="cora-form-grid" key="step-review">
      <div className="cora-section-divider"><span>Review Your Order</span></div>
      <div className="cora-field span">
        <div className="cora-review-table">
          <table>
            <tbody>
              {rows.map(([k, v]) => (
                <tr key={k}>
                  <td style={{ fontWeight: 600, color: 'var(--tx-secondary)', width: 160 }}>{k}</td>
                  <td>{v}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      {completeWasteLines.length > 0 && (
        <>
          <div className="cora-section-divider"><span>Waste Summary</span></div>
          <div className="cora-field span">
            <div className="cora-review-table">
              <table>
                <thead>
                  <tr>
                    <th>Waste Category</th>
                    <th>Tonnage</th>
                  </tr>
                </thead>
                <tbody>
                  {completeWasteLines.map((line) => (
                    <tr key={line.id}>
                      <td>{line.WasteCategory}</td>
                      <td>{line.Tonnage}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
      <div className="cora-section-divider"><span>Proof Preview</span></div>
      <div className="cora-field span">
        <div className="cora-proof-preview">
          {signatureDataUrl && (
            <div>
              <span className="cora-proof-preview-label">Signature</span>
              <img src={signatureDataUrl} alt="Customer signature preview" />
            </div>
          )}
          {beforePhotoPreviewUrl && (
            <div>
              <span className="cora-proof-preview-label">Before collection</span>
              <img src={beforePhotoPreviewUrl} alt="Before collection preview" />
            </div>
          )}
          {afterPhotoPreviewUrl && (
            <div>
              <span className="cora-proof-preview-label">After collection</span>
              <img src={afterPhotoPreviewUrl} alt="After collection preview" />
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

/* ═══════════════════════════════════════════════════════
   App
   ═══════════════════════════════════════════════════════ */
function App() {
  /* ── State ── */
  const [view, setView] = useState<View>('form')
  const [adminTab, setAdminTab] = useState<AdminTab>('drivers')
  const [currentUserEmail, setCurrentUserEmail] = useState('')
  const [adminPasscodeOpen, setAdminPasscodeOpen] = useState(false)
  const [pendingProtectedDestination, setPendingProtectedDestination] = useState<ProtectedDestination>({ view: 'admin', tab: 'drivers' })
  const [adminUnlocked, setAdminUnlocked] = useState(() => sessionStorage.getItem(ADMIN_UNLOCK_STORAGE_KEY) === 'true')
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false)

  const [customerAreas, setCustomerAreas] = useState<CustomerAreaOption[]>([])
  const [drivers, setDrivers] = useState<Drivers1Read[]>([])
  const [vehicles, setVehicles] = useState<VehiclesRead[]>([])
  const [categories, setCategories] = useState<Waste_CategoriesRead[]>([])
  const [orders, setOrders] = useState<ServiceOrdersRead[]>([])
  const [loadState, setLoadState] = useState<Load>('loading')
  const [loadError, setLoadError] = useState('')

  const [form, setForm] = useState<Form>(() => createBlankForm())
  const [step, setStep] = useState(0)
  const [submitting, setSubmitting] = useState(false)
  const [submitted, setSubmitted] = useState(false)
  const [toast, setToast] = useState<Toast | null>(null)
  const [adhocModalOpen, setAdhocModalOpen] = useState(false)
  const [resetModalOpen, setResetModalOpen] = useState(false)
  const [signatureDataUrl, setSignatureDataUrl] = useState('')
  const [beforePhotoFile, setBeforePhotoFile] = useState<File | null>(null)
  const [afterPhotoFile, setAfterPhotoFile] = useState<File | null>(null)
  const scalePhotoPreviewUrlsRef = useRef<string[]>([])

  const loadReferenceData = useCallback(async () => {
    setLoadState('loading')
    setLoadError('')

    const timeout = new Promise<never>((_, reject) => {
      window.setTimeout(() => reject(new Error('Reference data load timed out.')), 10000)
    })

    try {
      const loadCustomers = async () => {
        const clean = await Customer_area_data_clean_finalService.getAll({ top: 5000 })
        if (!clean.data || !Array.isArray(clean.data) || clean.data.length === 0) {
          throw new Error('The customer-area-data-clean-final list returned no customer rows.')
        }
        return clean.data
      }

      const [c, d, v, w, o] = await Promise.race([
        Promise.all([
          loadCustomers(), Drivers1Service.getAll(), VehiclesService.getAll(),
          Waste_CategoriesService.getAll(), ServiceOrdersService.getAll(),
        ]),
        timeout,
      ])
      setCustomerAreas(c)
      setDrivers(d.data && Array.isArray(d.data) ? d.data : [])
      setVehicles(v.data && Array.isArray(v.data) ? v.data : [])
      setCategories(w.data && Array.isArray(w.data) ? w.data : [])
      setOrders(o.data && Array.isArray(o.data) ? o.data : [])
      setLoadState('loaded')
    } catch (error) {
      setLoadError(error instanceof Error ? error.message : 'Could not load reference data.')
      setLoadState('error')
    }
  }, [])

  useEffect(() => {
    void loadReferenceData()
  }, [loadReferenceData])

  useEffect(() => {
    let cancelled = false
    getContext()
      .then((context) => {
        if (!cancelled) setCurrentUserEmail(context.user.userPrincipalName?.toLowerCase() || '')
      })
      .catch(() => {
        if (!cancelled) setCurrentUserEmail('')
      })
    return () => { cancelled = true }
  }, [])

  useEffect(() => {
    scalePhotoPreviewUrlsRef.current = form.WasteItems
      .map((line) => line.scalePhotoPreviewUrl)
      .filter((value): value is string => Boolean(value))
  }, [form.WasteItems])

  useEffect(() => () => {
    for (const previewUrl of scalePhotoPreviewUrlsRef.current) {
      URL.revokeObjectURL(previewUrl)
    }
  }, [])

  /* ── Helpers ── */
  const set = useCallback(<K extends keyof Form>(k: K, v: Form[K]) => setForm((p) => ({ ...p, [k]: v })), [])

  const isAdmin = useMemo(
    () => Boolean(currentUserEmail && ADMIN_EMAILS.includes(currentUserEmail)),
    [currentUserEmail],
  )

  const showAdmin = useCallback((tab: AdminTab) => {
    setAdminTab(tab)
    setView('admin')
  }, [])

  const openAdmin = useCallback((tab: AdminTab) => {
    if (!isAdmin) {
      setToast({ type: 'error', text: 'You do not have access to admin pages.' })
      return
    }
    if (adminUnlocked) {
      showAdmin(tab)
      return
    }
    setPendingProtectedDestination({ view: 'admin', tab })
    setAdminPasscodeOpen(true)
  }, [adminUnlocked, isAdmin, showAdmin])

  const navigate = useCallback((nextView: View) => {
    if (nextView === 'list') {
      if (!isAdmin) {
        setToast({ type: 'error', text: 'You do not have access to service orders.' })
        setView('form')
        return
      }
      if (!adminUnlocked) {
        setPendingProtectedDestination({ view: 'list' })
        setAdminPasscodeOpen(true)
        return
      }
    }
    if (nextView === 'admin' && (!isAdmin || !adminUnlocked)) {
      setToast({ type: 'error', text: 'You do not have access to admin pages.' })
      setView('form')
      return
    }
    setView(nextView)
  }, [adminUnlocked, isAdmin])

  useEffect(() => {
    if (view === 'admin' && (!isAdmin || !adminUnlocked)) setView('form')
    if (view === 'list' && (!isAdmin || !adminUnlocked)) setView('form')
  }, [adminUnlocked, isAdmin, view])

  const unlockAdmin = useCallback(() => {
    sessionStorage.setItem(ADMIN_UNLOCK_STORAGE_KEY, 'true')
    setAdminUnlocked(true)
    setAdminPasscodeOpen(false)
    if (pendingProtectedDestination.view === 'list') {
      setView('list')
      return
    }
    showAdmin(pendingProtectedDestination.tab)
  }, [pendingProtectedDestination, showAdmin])

  const driverOptions = useMemo(
    () => uniqueValues(drivers.map(getDriverName)),
    [drivers],
  )

  const beforePhotoPreviewUrl = useMemo(
    () => (beforePhotoFile ? URL.createObjectURL(beforePhotoFile) : ''),
    [beforePhotoFile],
  )

  const afterPhotoPreviewUrl = useMemo(
    () => (afterPhotoFile ? URL.createObjectURL(afterPhotoFile) : ''),
    [afterPhotoFile],
  )

  useEffect(() => () => {
    if (beforePhotoPreviewUrl) URL.revokeObjectURL(beforePhotoPreviewUrl)
  }, [beforePhotoPreviewUrl])

  useEffect(() => () => {
    if (afterPhotoPreviewUrl) URL.revokeObjectURL(afterPhotoPreviewUrl)
  }, [afterPhotoPreviewUrl])

  const reset = useCallback(() => {
    revokeScalePhotoPreviewUrls(form.WasteItems)
    setForm(createBlankForm())
    setSignatureDataUrl('')
    setBeforePhotoFile(null)
    setAfterPhotoFile(null)
    setAdhocModalOpen(false)
    setStep(0)
    setSubmitted(false)
    setSubmitting(false)
  }, [form.WasteItems])

  const clearCurrentStepFields = useCallback(() => {
    if (submitting || submitted) return

    if (step === 0) {
      setForm((current) => ({
        ...current,
        DateOfCollection: formatDateInputValue(new Date()),
      }))
      setToast({ type: 'success', text: 'Basic info fields cleared.' })
      return
    }

    if (step === 1) {
      setForm((current) => ({
        ...current,
        Customer: '',
        CustomerName: '',
        CustomerLocation: '',
        CustomerTenant: '',
        CustomerLevel4: '',
        CustomerLevel5: '',
        CustomerLevel6: '',
        CustomerLevel7: '',
        CustomerLevel8: '',
        CustomerLevel9: '',
        CustomerLevel10: '',
        IsAdhocCustomer: false,
      }))
      setAdhocModalOpen(false)
      setToast({ type: 'success', text: 'Customer fields cleared.' })
      return
    }

    if (step === 2) {
      setForm((current) => ({
        ...current,
        DriverName: '',
        VehicleNumber: '',
        WasteItems: [createWasteLine()],
        Notes: '',
      }))
      revokeScalePhotoPreviewUrls(form.WasteItems)
      setToast({ type: 'success', text: 'Assignment fields cleared.' })
      return
    }

    if (step === 3) {
      setSignatureDataUrl('')
      setBeforePhotoFile(null)
      setAfterPhotoFile(null)
      setToast({ type: 'success', text: 'Proof fields cleared.' })
      return
    }

    setToast({ type: 'error', text: 'There are no editable fields to clear on the review page.' })
  }, [form.WasteItems, step, submitted, submitting])

  const addDriver = useCallback(async (name: string) => {
    const cleaned = name.trim()
    if (!cleaned) return
    if (driverOptions.some((driver) => normalize(driver) === normalize(cleaned))) {
      setToast({ type: 'error', text: `"${cleaned}" already exists in drivers.` })
      return
    }
    try {
      const result = await Drivers1Service.create({ Title: cleaned })
      if (!result.data?.ID) throw new Error('Could not add the driver.')
      setDrivers((current) => [...current, result.data as Drivers1Read])
      setToast({ type: 'success', text: `Driver "${cleaned}" added.` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not add the driver.' })
    }
  }, [driverOptions])

  const addVehicle = useCallback(async (vehicleNumber: string) => {
    const cleaned = vehicleNumber.trim()
    if (!cleaned) return
    if (vehicles.some((vehicle) => normalize(vehicle.Title || '') === normalize(cleaned))) {
      setToast({ type: 'error', text: `"${cleaned}" already exists in vehicles.` })
      return
    }
    try {
      const result = await VehiclesService.create({ Title: cleaned } as Omit<VehiclesWrite, 'ID'>)
      if (!result.data?.ID) throw new Error('Could not add the vehicle.')
      setVehicles((current) => [...current, result.data as VehiclesRead])
      setToast({ type: 'success', text: `Vehicle "${cleaned}" added.` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not add the vehicle.' })
    }
  }, [vehicles])

  const addWasteCategory = useCallback(async (category: string) => {
    const cleaned = category.trim()
    if (!cleaned) return
    if (categories.some((item) => normalize(item.Title || '') === normalize(cleaned))) {
      setToast({ type: 'error', text: `"${cleaned}" already exists in waste categories.` })
      return
    }
    try {
      const result = await Waste_CategoriesService.create({ Title: cleaned } as Omit<Waste_CategoriesWrite, 'ID'>)
      if (!result.data?.ID) throw new Error('Could not add the waste category.')
      setCategories((current) => [...current, result.data as Waste_CategoriesRead])
      setToast({ type: 'success', text: `Waste category "${cleaned}" added.` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not add the waste category.' })
    }
  }, [categories])

  const updateDriver = useCallback(async (id: number, name: string) => {
    const cleaned = name.trim()
    if (!cleaned) return
    if (drivers.some((driver) => driver.ID !== id && normalize(getDriverName(driver)) === normalize(cleaned))) {
      setToast({ type: 'error', text: `"${cleaned}" already exists in drivers.` })
      return
    }
    try {
      const result = await Drivers1Service.update(String(id), { Title: cleaned })
      setDrivers((current) => current.map((driver) => (
        driver.ID === id ? (result.data ?? { ...driver, Title: cleaned }) as Drivers1Read : driver
      )))
      setToast({ type: 'success', text: `Driver updated to "${cleaned}".` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not update the driver.' })
    }
  }, [drivers])

  const updateVehicle = useCallback(async (id: number, vehicleNumber: string) => {
    const cleaned = vehicleNumber.trim()
    if (!cleaned) return
    if (vehicles.some((vehicle) => vehicle.ID !== id && normalize(vehicle.Title || '') === normalize(cleaned))) {
      setToast({ type: 'error', text: `"${cleaned}" already exists in vehicles.` })
      return
    }
    try {
      const result = await VehiclesService.update(String(id), { Title: cleaned })
      setVehicles((current) => current.map((vehicle) => (
        vehicle.ID === id ? (result.data ?? { ...vehicle, Title: cleaned }) as VehiclesRead : vehicle
      )))
      setToast({ type: 'success', text: `Vehicle updated to "${cleaned}".` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not update the vehicle.' })
    }
  }, [vehicles])

  const updateWasteCategory = useCallback(async (id: number, category: string) => {
    const cleaned = category.trim()
    if (!cleaned) return
    if (categories.some((item) => item.ID !== id && normalize(item.Title || '') === normalize(cleaned))) {
      setToast({ type: 'error', text: `"${cleaned}" already exists in waste categories.` })
      return
    }
    try {
      const result = await Waste_CategoriesService.update(String(id), { Title: cleaned })
      setCategories((current) => current.map((item) => (
        item.ID === id ? (result.data ?? { ...item, Title: cleaned }) as Waste_CategoriesRead : item
      )))
      setToast({ type: 'success', text: `Waste category updated to "${cleaned}".` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not update the waste category.' })
    }
  }, [categories])

  const deleteDriver = useCallback(async (id: number, name: string) => {
    try {
      await Drivers1Service.delete(String(id))
      setDrivers((current) => current.filter((driver) => driver.ID !== id))
      setToast({ type: 'success', text: `Driver "${name}" deleted.` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not delete the driver.' })
    }
  }, [])

  const deleteVehicle = useCallback(async (id: number, vehicleNumber: string) => {
    try {
      await VehiclesService.delete(String(id))
      setVehicles((current) => current.filter((vehicle) => vehicle.ID !== id))
      setToast({ type: 'success', text: `Vehicle "${vehicleNumber}" deleted.` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not delete the vehicle.' })
    }
  }, [])

  const deleteWasteCategory = useCallback(async (id: number, category: string) => {
    try {
      await Waste_CategoriesService.delete(String(id))
      setCategories((current) => current.filter((item) => item.ID !== id))
      setToast({ type: 'success', text: `Waste category "${category}" deleted.` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not delete the waste category.' })
    }
  }, [])

  const saveAdhocCustomer = useCallback((values: Pick<Form, 'CustomerName' | 'CustomerLocation' | 'CustomerTenant'>) => {
    setForm((current) => ({
      ...current,
      Customer: '',
      CustomerName: values.CustomerName,
      CustomerLocation: values.CustomerLocation,
      CustomerTenant: values.CustomerTenant,
      CustomerLevel4: '',
      CustomerLevel5: '',
      CustomerLevel6: '',
      CustomerLevel7: '',
      CustomerLevel8: '',
      CustomerLevel9: '',
      CustomerLevel10: '',
      IsAdhocCustomer: true,
    }))
    setAdhocModalOpen(false)
  }, [])

  const getCustomerStepIssue = () => {
    const customerName = form.CustomerName.trim()
    if (!customerName) return 'Please select a customer before continuing.'
    if (form.IsAdhocCustomer) return ''

    const customerExists = customerAreas.some((option) => getCustomerName(option) === customerName)
    if (!customerExists) return 'Please select a customer from the list before continuing.'

    for (let levelIndex = 1; levelIndex < CUSTOMER_LEVELS.length; levelIndex += 1) {
      const options = getCustomerLevelOptions(customerAreas, form, levelIndex)
      if (options.length > 0 && !getFormCustomerLevelValue(form, levelIndex)) {
        return `Please select ${CUSTOMER_LEVELS[levelIndex].label.toLowerCase()} before continuing.`
      }
    }

    return ''
  }

  const canAdvance = (s: number) => {
    if (loadState !== 'loaded') return false
    if (s === 0 && !form.Title.trim()) return false
    if (s === 1 && getCustomerStepIssue()) return false
    if (s === 2 && hasIncompleteWasteLine(form.WasteItems)) return false
    if (s === 3 && !signatureDataUrl) return false
    return true
  }

  const getStepBlockedMessage = (s: number) => {
    if (loadState !== 'loaded') return 'Reference data must load before continuing.'
    if (s === 1) return getCustomerStepIssue() || 'Finish this page before moving to the next step.'
    if (s === 2 && hasIncompleteWasteLine(form.WasteItems)) return 'Complete each waste category and tonnage row before continuing.'
    if (s === 3 && !signatureDataUrl) return 'Please add the customer signature before continuing.'
    return 'Finish this page before moving to the next step.'
  }

  const showStepBlockedToast = (s: number) => {
    setToast({ type: 'error', text: getStepBlockedMessage(s) })
  }

  const next = () => {
    if (step >= LAST_STEP) return
    if (!canAdvance(step)) {
      showStepBlockedToast(step)
      return
    }
    setStep(step + 1)
  }
  const prev = () => { if (step > 0) setStep(step - 1) }

  const goToStep = (targetStep: number) => {
    if (submitting || submitted || targetStep === step) return
    if (targetStep < step) {
      setStep(targetStep)
      return
    }
    showStepBlockedToast(step)
  }

  /* ── Submit ── */
  const submit = useCallback(async () => {
    if (loadState !== 'loaded') {
      setToast({ type: 'error', text: 'Reference data is not ready. Retry loading before submitting.' })
      return
    }
    if (!form.Title.trim()) {
      setToast({ type: 'error', text: 'Title is required.' })
      return
    }
    if (!signatureDataUrl) {
      setToast({ type: 'error', text: 'Please add the customer signature before submitting.' })
      return
    }
    if (hasIncompleteWasteLine(form.WasteItems)) {
      setToast({ type: 'error', text: 'Complete each waste category and tonnage row before submitting.' })
      return
    }
    setSubmitting(true)
    const createdOrders: ServiceOrdersRead[] = []
    const createdWasteItems: ServiceOrderWasteItemsRead[] = []
    let createdProofQueue: ServiceOrderProofQueueRead | null = null

    const rollbackCreatedRecords = async () => {
      const failures: string[] = []
      const tryDelete = async (label: string, action: () => Promise<void>) => {
        try {
          await action()
        } catch (error) {
          failures.push(`${label}: ${error instanceof Error ? error.message : 'cleanup failed'}`)
        }
      }

      if (createdProofQueue?.ID) {
        await tryDelete(`proof queue ${createdProofQueue.ID}`, () => ServiceOrderProofQueueService.delete(String(createdProofQueue?.ID)))
      }

      for (const wasteItem of [...createdWasteItems].reverse()) {
        if (wasteItem.ID) {
          await tryDelete(`waste item ${wasteItem.ID}`, () => ServiceOrderWasteItemsService.delete(String(wasteItem.ID)))
        }
      }

      for (const order of [...createdOrders].reverse()) {
        if (order.ID) {
          await tryDelete(`service order ${order.ID}`, () => ServiceOrdersService.delete(String(order.ID)))
        }
      }

      return failures
    }

    try {
      const wasteLines = getCompleteWasteLines(form.WasteItems)
      const orderTitle = form.Title || generateOrderTitle()
      const basePayload: Partial<Omit<ServiceOrdersWrite, 'ID'>> = {
        Title: orderTitle,
        Customer: form.Customer || undefined,
        CustomerName: form.CustomerName || undefined,
        CustomerLocation: form.CustomerLocation || undefined,
        CustomerTenant: form.CustomerTenant || undefined,
        CustomerLevel4: form.CustomerLevel4 || undefined,
        CustomerLevel5: form.CustomerLevel5 || undefined,
        CustomerLevel6: form.CustomerLevel6 || undefined,
        CustomerLevel7: form.CustomerLevel7 || undefined,
        CustomerLevel8: form.CustomerLevel8 || undefined,
        CustomerLevel9: form.CustomerLevel9 || undefined,
        CustomerLevel10: form.CustomerLevel10 || undefined,
        IsAdhocCustomer: form.IsAdhocCustomer,
        DriverName: form.DriverName || undefined,
        VehicleNumber: form.VehicleNumber || undefined,
        DateOfCollection: form.DateOfCollection || undefined,
        Notes: form.Notes || undefined,
      }

      const serviceOrderRows = wasteLines.length > 0 ? wasteLines : [null]
      for (const line of serviceOrderRows) {
        const tonnage = line ? Number(line.Tonnage) : undefined
        const payload: Partial<Omit<ServiceOrdersWrite, 'ID'>> = {
          ...basePayload,
          WasteCategory: line?.WasteCategory || undefined,
          Tonnage: Number.isFinite(tonnage) ? tonnage : undefined,
        }
        const res = await ServiceOrdersService.create(payload as Omit<ServiceOrdersWrite, 'ID'>)
        if (res.data?.ID) {
          createdOrders.push(res.data)
        }
      }

      if (createdOrders.length < serviceOrderRows.length) {
        const foundOrders = await findServiceOrdersByTitle(orderTitle, serviceOrderRows.length)
        createdOrders.splice(0, createdOrders.length, ...foundOrders)
      }

      if (createdOrders.length < serviceOrderRows.length) {
        throw new Error(`Could not confirm the created service order rows for "${orderTitle}".`)
      }

      const proofOwnerOrder = createdOrders[0]
      if (!proofOwnerOrder?.ID) throw new Error('Could not create the service order.')

      if (wasteLines.length > 0) {
        for (const line of wasteLines) {
          const tonnage = Number(line.Tonnage)
          const wasteItem: Partial<Omit<ServiceOrderWasteItemsWrite, 'ID'>> = {
            Title: orderTitle,
            WasteCategory: line.WasteCategory,
            Tonnage: Number.isFinite(tonnage) ? tonnage : undefined,
          }
          const wasteResult = await ServiceOrderWasteItemsService.create(wasteItem as Omit<ServiceOrderWasteItemsWrite, 'ID'>)
          if (wasteResult.data?.ID) createdWasteItems.push(wasteResult.data)
        }
      }

      const proofUpload = await uploadServiceOrderProof({
        serviceOrderId: Number(proofOwnerOrder.ID),
        orderTitle,
        signatureFileName: `${orderTitle}-signature.png`,
        signatureBase64: dataUrlToBase64(signatureDataUrl),
        beforePhotoFileName: beforePhotoFile ? getPhotoFileName(orderTitle, 'before', beforePhotoFile) : undefined,
        beforePhotoBase64: beforePhotoFile ? await fileToBase64(beforePhotoFile) : undefined,
        afterPhotoFileName: afterPhotoFile ? getPhotoFileName(orderTitle, 'after', afterPhotoFile) : undefined,
        afterPhotoBase64: afterPhotoFile ? await fileToBase64(afterPhotoFile) : undefined,
      })
      createdProofQueue = proofUpload.queueItem ?? null

      const proofUrlFields: Partial<Omit<ServiceOrdersWrite, 'ID'>> = {
        SignatureUrl: proofUpload.signatureUrl || undefined,
        BeforePhotoUrl: proofUpload.beforePhotoUrl || undefined,
        AfterPhotoUrl: proofUpload.afterPhotoUrl || undefined,
      }
      const hasProofUrls = Object.values(proofUrlFields).some(Boolean)
      let savedOrders = [...createdOrders]
      if (hasProofUrls) {
        savedOrders = []
        for (const order of createdOrders) {
          if (!order.ID) continue
          const updated = await ServiceOrdersService.update(String(order.ID), proofUrlFields)
          savedOrders.push((updated.data ?? { ...order, ...proofUrlFields }) as ServiceOrdersRead)
        }
      }

      setOrders((p) => [...savedOrders, ...p])
      const rowText = savedOrders.length === 1 ? '1 service order row' : `${savedOrders.length} service order rows`
      setToast({ type: 'success', text: `Order "${orderTitle}" created with ${rowText}. Proof media queued for processing.` })
      setSubmitted(true)
    } catch (e) {
      const rollbackFailures = await rollbackCreatedRecords()
      const orderIdText = createdOrders.length > 0
        ? ` Service order IDs: ${createdOrders.map((order) => order.ID).filter(Boolean).join(', ')}.`
        : ''
      const rollbackText = rollbackFailures.length > 0
        ? ` Rollback needs manual cleanup.${orderIdText}`
        : ' Created records were rolled back; you can retry.'
      setToast({ type: 'error', text: `${e instanceof Error ? e.message : 'Unexpected error.'}${rollbackText}` })
    } finally {
      setSubmitting(false)
    }
  }, [afterPhotoFile, beforePhotoFile, form, loadState, signatureDataUrl])

  /* ── Step completion states ── */
  const stepDone = (idx: number) => {
    if (submitted) return true
    return idx < step
  }

  /* ── Render ── */
  return (
    <div className="cora-shell">
      {/* Mobile hamburger — only opens; close button is inside the sidebar */}
      {!sidebarOpen && (
        <button className="cora-mobile-toggle" onClick={() => setSidebarOpen(true)} aria-label="Open menu">
          {I.menu}
        </button>
      )}

      <Sidebar
        view={view}
        onNav={navigate}
        onAdminNav={openAdmin}
        open={sidebarOpen}
        collapsed={sidebarCollapsed}
        onClose={() => setSidebarOpen(false)}
        isAdmin={isAdmin}
      />

      <main className="cora-main">
        {/* Top bar */}
        <div className="cora-topbar">
          <button className="cora-desktop-toggle" onClick={() => setSidebarCollapsed(!sidebarCollapsed)} aria-label="Toggle Sidebar">
            {I.menu}
          </button>
          <div className="cora-breadcrumb">
            CORA <span className="cora-breadcrumb-sep">/</span>{' '}
            <span className="cora-breadcrumb-current">
              {view === 'form' ? 'New Service Order' : view === 'admin' ? 'Admin' : 'All Orders'}
            </span>
          </div>
        </div>

        <div className="cora-page">
          {/* Header */}
          <div className="cora-page-header">
            <h1>{view === 'form' ? 'Create Service Order' : view === 'admin' ? 'Admin Reference Data' : 'Service Orders'}</h1>
            <p>{view === 'form'
              ? 'Complete the wizard below to submit a new service order.'
              : view === 'admin'
                ? 'Add drivers, vehicles, and waste categories for the service order form.'
                : 'View and manage all service order records.'}</p>
          </div>

          {/* ════ LIST VIEW ════ */}
          {view === 'list' && isAdmin && adminUnlocked && <OrdersTable orders={orders} loadState={loadState} />}

          {view === 'admin' && isAdmin && adminUnlocked && (
            <AdminReferencePage
              activeTab={adminTab}
              onTabChange={setAdminTab}
              drivers={drivers}
              vehicles={vehicles}
              categories={categories}
              busy={loadState !== 'loaded'}
              currentUserEmail={currentUserEmail}
              onAddDriver={addDriver}
              onAddVehicle={addVehicle}
              onAddWasteCategory={addWasteCategory}
              onUpdateDriver={updateDriver}
              onUpdateVehicle={updateVehicle}
              onUpdateWasteCategory={updateWasteCategory}
              onDeleteDriver={deleteDriver}
              onDeleteVehicle={deleteVehicle}
              onDeleteWasteCategory={deleteWasteCategory}
            />
          )}

          {/* ════ FORM WIZARD VIEW ════ */}
          {view === 'form' && (
            <>
              {/* Stepper */}
              <div className="cora-stepper">
                {STEPS.map((s, i) => (
                  <button
                    key={s.label}
                    className={`cora-step ${i === step && !submitted ? 'active' : ''} ${stepDone(i) ? 'completed' : ''}`}
                    onClick={() => goToStep(i)}
                  >
                    <span className="cora-step-num">
                      {stepDone(i) ? <span style={{ display: 'flex', width: 14, height: 14 }}>{I.check}</span> : i + 1}
                    </span>
                    <span className="cora-step-info">
                      <span className="cora-step-label">{s.label}</span>
                      <span className="cora-step-desc">{s.desc}</span>
                    </span>
                  </button>
                ))}
              </div>

              {/* Card */}
              <div className="cora-card">
                <div className="cora-card-head">
                  <span className="cora-card-title">
                    {submitted ? 'Order Created' : STEPS[step].label}
                  </span>
                  <span className="cora-card-badge">
                    {submitted ? 'DONE' : `Step ${step + 1} of ${STEPS.length}`}
                  </span>
                </div>

                <div className="cora-card-body">
                  {loadState === 'loading' && (
                    <div className="cora-form-grid">
                      {Array.from({ length: 4 }).map((_, i) => (
                        <div className="cora-field" key={i}>
                          <div className="cora-skeleton" style={{ height: 14, width: '35%', marginBottom: 6 }} />
                          <div className="cora-skeleton" />
                        </div>
                      ))}
                    </div>
                  )}

                  {loadState === 'error' && (
                    <div className="cora-load-error">
                      <h2>Reference Data Unavailable</h2>
                      <p>{loadError || 'Could not load reference data.'}</p>
                      <button className="cora-btn cora-btn-primary" type="button" onClick={() => void loadReferenceData()}>
                        {I.refresh} Retry Loading
                      </button>
                    </div>
                  )}

                  {loadState === 'loaded' && !submitted && (
                    <>
                      {step === 0 && <StepBasic form={form} set={set} busy={submitting} />}
                      {step === 1 && (
                        <StepCustomer
                          form={form}
                          set={set}
                          busy={submitting}
                          customerAreas={customerAreas}
                          onOpenAdhocModal={() => setAdhocModalOpen(true)}
                        />
                      )}
                      {step === 2 && (
                        <StepAssignment
                          form={form}
                          set={set}
                          updateForm={(updater) => setForm(updater)}
                          busy={submitting}
                          driverOptions={driverOptions}
                          vehicles={vehicles}
                          categories={categories}
                        />
                      )}
                      {step === 3 && (
                        <StepProof
                          signatureDataUrl={signatureDataUrl}
                          setSignatureDataUrl={setSignatureDataUrl}
                          beforePhotoFile={beforePhotoFile}
                          setBeforePhotoFile={setBeforePhotoFile}
                          afterPhotoFile={afterPhotoFile}
                          setAfterPhotoFile={setAfterPhotoFile}
                          beforePhotoPreviewUrl={beforePhotoPreviewUrl}
                          afterPhotoPreviewUrl={afterPhotoPreviewUrl}
                          busy={submitting}
                        />
                      )}
                      {step === 4 && (
                        <StepReview
                          form={form}
                          signatureDataUrl={signatureDataUrl}
                          beforePhotoPreviewUrl={beforePhotoPreviewUrl}
                          afterPhotoPreviewUrl={afterPhotoPreviewUrl}
                        />
                      )}
                    </>
                  )}

                  {loadState === 'loaded' && submitted && (
                    <div className="cora-success-panel">
                      <div className="cora-success-icon">{I.check}</div>
                      <h2>Service Order Created</h2>
                      <p>Your order has been submitted successfully and is now being processed.</p>
                      <button className="cora-btn cora-btn-primary" onClick={reset} id="btn-new-order">
                        {I.plus} Create Another Order
                      </button>
                    </div>
                  )}
                </div>

                {loadState === 'loaded' && !submitted && (
                  <div className="cora-card-foot">
                    <div className="cora-card-foot-left">
                      <button
                        className="cora-btn cora-btn-outline cora-btn-danger"
                        onClick={() => setResetModalOpen(true)}
                        disabled={submitting}
                        id="btn-reset"
                      >
                        {I.refresh} Start Over
                      </button>
                      <button
                        className="cora-btn cora-btn-ghost"
                        onClick={clearCurrentStepFields}
                        disabled={submitting || step === LAST_STEP}
                        id="btn-clear-step"
                      >
                        {I.x} Clear Page
                      </button>
                    </div>
                    <div className="cora-card-foot-right">
                      {step > 0 && (
                        <button className="cora-btn cora-btn-outline" onClick={prev} disabled={submitting} id="btn-prev">
                          {I.chevL} Back
                        </button>
                      )}
                      {step < LAST_STEP ? (
                        <button
                          className="cora-btn cora-btn-teal"
                          onClick={next}
                          disabled={submitting}
                          id="btn-next"
                        >
                          Next {I.chevR}
                        </button>
                      ) : (
                        <button
                          className="cora-btn cora-btn-primary"
                          onClick={submit}
                          disabled={submitting || loadState !== 'loaded' || !form.Title.trim() || !signatureDataUrl}
                          id="btn-submit"
                        >
                          {submitting
                            ? <><span className="cora-spinner" /> Submitting…</>
                            : <>{I.send} Submit Order</>}
                        </button>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </>
          )}
        </div>
      </main>

      {toast && <ToastNotif toast={toast} onClose={() => setToast(null)} />}
      {adminPasscodeOpen && (
        <AdminPasscodeModal
          onCancel={() => setAdminPasscodeOpen(false)}
          onUnlock={unlockAdmin}
        />
      )}
      {resetModalOpen && (
        <ConfirmResetModal
          onCancel={() => setResetModalOpen(false)}
          onConfirm={() => {
            setResetModalOpen(false)
            reset()
          }}
        />
      )}
      {adhocModalOpen && (
        <AdhocCustomerModal
          initialValues={{
            CustomerName: form.CustomerName,
            CustomerLocation: form.CustomerLocation,
            CustomerTenant: form.CustomerTenant,
          }}
          onCancel={() => setAdhocModalOpen(false)}
          onSave={saveAdhocCustomer}
          busy={submitting}
        />
      )}
    </div>
  )
}

export default App
