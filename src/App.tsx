/* ═══════════════════════════════════════════════════════
   CORA Environment — Service Order Wizard
   Single-file architecture: ~18 inline icons, 4-step
   wizard, CRUD via PowerApps SharePoint connector.
   ═══════════════════════════════════════════════════════ */

import { useEffect, useState, useCallback, useMemo, useRef } from 'react'
import type { ChangeEvent, DragEvent, PointerEvent as ReactPointerEvent, ReactNode } from 'react'
import { getContext } from '@microsoft/power-apps/app'
import type { Drivers1Read } from './generated/models/Drivers1Model'
import type { Customer_area_data_clean_finalRead, Customer_area_data_clean_finalWrite } from './generated/models/Customer_area_data_clean_finalModel'
import type { VehiclesRead, VehiclesWrite } from './generated/models/VehiclesModel'
import type { Waste_CategoriesRead, Waste_CategoriesWrite } from './generated/models/Waste_CategoriesModel'
import type { ServiceOrdersRead } from './generated/models/ServiceOrdersModel'
import { Drivers1Service } from './generated/services/Drivers1Service'
import { Customer_area_data_clean_finalService } from './generated/services/Customer_area_data_clean_finalService'
import { VehiclesService } from './generated/services/VehiclesService'
import { Waste_CategoriesService } from './generated/services/Waste_CategoriesService'
import { ServiceOrdersService } from './generated/services/ServiceOrdersService'
import { preprocessImageForOcr, runWeightOcr } from './weightOcr'
import type { CropRect } from './weightOcr'
import type { CachedReferenceData, Form, PersistedDraft, PersistedForm, PersistedWasteLine, WasteLine } from './lib/offlineStore'
import { clearDraft, fromStoredMediaFile, getQueueSummary, loadDraft, saveDraft, toStoredMediaFile } from './lib/offlineStore'
import { createClientSubmissionId, enqueueSubmission, processPendingQueue } from './lib/offlineQueue'
import { readCachedReferenceData, writeCachedReferenceData } from './lib/referenceCache'
import logoImage from './assets/logo.png?inline'
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
type AdminTab = 'customers' | 'drivers' | 'vehicles' | 'waste'
type ProtectedDestination = { view: 'list' } | { view: 'admin'; tab: AdminTab }
type SubmissionMode = 'draft' | 'queued'

interface Toast {
  type: 'success' | 'error'
  text: string
  actionLabel?: string
  onAction?: () => void
  durationMs?: number
}

type DriverOptionSource = Drivers1Read
type CustomerAreaOption = Customer_area_data_clean_finalRead
type CustomerAdminForm = {
  Title: string
  field_1: string
  field_2: string
  field_3: string
  field_4: string
  field_5: string
  field_6: string
  field_7: string
  field_8: string
  field_9: string
}
type CustomerAdminFieldKey = keyof CustomerAdminForm
type CustomerAdminErrors = Partial<Record<CustomerAdminFieldKey, string>>
type DeleteTarget =
  | { tab: 'customers'; id: number; name: string }
  | { tab: 'drivers'; id: number; name: string }
  | { tab: 'vehicles'; id: number; name: string }
  | { tab: 'waste'; id: number; name: string }
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
const createEmptyCustomerAdminForm = (): CustomerAdminForm => ({
  Title: '',
  field_1: '',
  field_2: '',
  field_3: '',
  field_4: '',
  field_5: '',
  field_6: '',
  field_7: '',
  field_8: '',
  field_9: '',
})

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

const getCustomerPathValues = (row: CustomerAreaOption) =>
  CUSTOMER_LEVELS
    .map((_, levelIndex) => getCustomerLevelValue(row, levelIndex))
    .filter(Boolean)

const getCustomerPathLabel = (row: CustomerAreaOption) => getCustomerPathValues(row).join(' / ')

const getCustomerPathOptions = (rows: CustomerAreaOption[]) => {
  const options = new Map<string, { value: string, label: string, row: CustomerAreaOption }>()
  rows.forEach((row) => {
    const label = getCustomerPathLabel(row)
    if (!label) return
    const key = normalize(label)
    if (!options.has(key)) options.set(key, { value: label, label, row })
  })
  return Array.from(options.values()).sort((a, b) => a.label.localeCompare(b.label))
}

const getFormCustomerPathValue = (form: Form) =>
  CUSTOMER_LEVELS
    .map((level) => String(form[level.key] ?? '').trim())
    .filter(Boolean)
    .join(' / ')

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

const CUSTOMER_ADMIN_FIELDS: { key: keyof CustomerAdminForm; label: string; placeholder: string; required?: boolean }[] = [
  { key: 'Title', label: 'Customer Name', placeholder: 'Enter customer name', required: true },
  { key: 'field_1', label: 'Location', placeholder: 'Enter location', required: true },
  { key: 'field_2', label: 'Sub-location / Path', placeholder: 'Enter sub-location or slash-delimited path' },
  { key: 'field_3', label: 'Level 4', placeholder: 'Optional level 4 value' },
  { key: 'field_4', label: 'Level 5', placeholder: 'Optional level 5 value' },
  { key: 'field_5', label: 'Level 6', placeholder: 'Optional level 6 value' },
  { key: 'field_6', label: 'Level 7', placeholder: 'Optional level 7 value' },
  { key: 'field_7', label: 'Level 8', placeholder: 'Optional level 8 value' },
  { key: 'field_8', label: 'Level 9', placeholder: 'Optional level 9 value' },
  { key: 'field_9', label: 'Level 10', placeholder: 'Optional level 10 value' },
]

const getCustomerAdminFormFromRow = (row: CustomerAreaOption): CustomerAdminForm => ({
  Title: row.Title?.trim() || '',
  field_1: readString(row.field_1),
  field_2: readString(row.field_2),
  field_3: readString(row.field_3),
  field_4: readString(row.field_4),
  field_5: readString(row.field_5),
  field_6: readString(row.field_6),
  field_7: readString(row.field_7),
  field_8: readString(row.field_8),
  field_9: readString(row.field_9),
})

const normalizeCustomerAdminForm = (value: CustomerAdminForm): CustomerAdminForm => ({
  Title: value.Title.trim(),
  field_1: value.field_1.trim(),
  field_2: value.field_2.trim(),
  field_3: value.field_3.trim(),
  field_4: value.field_4.trim(),
  field_5: value.field_5.trim(),
  field_6: value.field_6.trim(),
  field_7: value.field_7.trim(),
  field_8: value.field_8.trim(),
  field_9: value.field_9.trim(),
})

const isCustomerAdminFormValid = (value: CustomerAdminForm) =>
  Boolean(value.Title.trim() && value.field_1.trim())

const getCustomerAdminErrors = (
  value: CustomerAdminForm,
  existingRows: CustomerAreaOption[],
  editingId?: number,
): CustomerAdminErrors => {
  const cleaned = normalizeCustomerAdminForm(value)
  const errors: CustomerAdminErrors = {}
  if (!cleaned.Title) errors.Title = 'Customer name is required.'
  if (!cleaned.field_1) errors.field_1 = 'Location is required.'
  if (cleaned.field_2 && cleaned.field_2.startsWith('/')) {
    errors.field_2 = 'Sub-location path should not start with a slash.'
  }
  const isDuplicate = existingRows.some((row) => row.ID !== editingId && (
    normalize(getCustomerName(row)) === normalize(cleaned.Title)
    && normalize(getCustomerArea(row)) === normalize(cleaned.field_1)
    && normalize(getCustomerSubLocationRaw(row)) === normalize(cleaned.field_2)
  ))
  if (isDuplicate) {
    errors.Title = 'A customer with this name and location already exists.'
    if (!cleaned.field_2) errors.field_1 = 'Change the location or sub-location to create a unique customer.'
  }
  return errors
}

const getCustomerAdminSummary = (row: CustomerAreaOption) => {
  const parts = [
    getCustomerArea(row),
    getCustomerSubLocationRaw(row),
    readString(row.field_3),
    readString(row.field_4),
    readString(row.field_5),
    readString(row.field_6),
    readString(row.field_7),
    readString(row.field_8),
    readString(row.field_9),
  ].filter(Boolean)
  return parts.join(' / ')
}

const pad = (value: number) => String(value).padStart(2, '0')

const formatDateInputValue = (date: Date) =>
  `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`

const fileToDataUrl = (file: Blob | File) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => {
      if (typeof reader.result === 'string') {
        resolve(reader.result)
        return
      }
      reject(new Error('Could not read image preview.'))
    }
    reader.onerror = () => reject(reader.error ?? new Error('Could not read image preview.'))
    reader.readAsDataURL(file)
  })

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

const buildPersistedForm = async (form: Form): Promise<PersistedForm> => ({
  ...form,
  WasteItems: await Promise.all(form.WasteItems.map(async (line): Promise<PersistedWasteLine> => ({
    id: line.id,
    WasteCategory: line.WasteCategory,
    Tonnage: line.Tonnage,
    scalePhoto: await toStoredMediaFile(line.scalePhotoFile, `${line.id}-scale-photo.jpg`),
    scaleOcrCropPreviewDataUrl: line.scaleOcrCropPreviewUrl,
    scaleOcrCrop: line.scaleOcrCrop,
    scaleOcrStatus: line.scaleOcrStatus,
    scaleOcrText: line.scaleOcrText,
    scaleOcrSuggestion: line.scaleOcrSuggestion,
    scaleOcrConfidence: line.scaleOcrConfidence,
    scaleOcrReasons: line.scaleOcrReasons,
    scaleOcrError: line.scaleOcrError,
    scaleOcrRequestId: line.scaleOcrRequestId,
  }))),
})

const restorePersistedForm = async (persisted: PersistedForm): Promise<Form> => ({
  ...persisted,
  WasteItems: await Promise.all(persisted.WasteItems.map(async (line): Promise<WasteLine> => {
    const scalePhotoFile = await fromStoredMediaFile(line.scalePhoto)
    return {
      id: line.id,
      WasteCategory: line.WasteCategory,
      Tonnage: line.Tonnage,
      scalePhotoFile: scalePhotoFile ?? undefined,
      scalePhotoPreviewUrl: scalePhotoFile ? await fileToDataUrl(scalePhotoFile) : undefined,
      scaleOcrCropPreviewUrl: line.scaleOcrCropPreviewDataUrl,
      scaleOcrCrop: line.scaleOcrCrop,
      scaleOcrStatus: line.scaleOcrStatus,
      scaleOcrText: line.scaleOcrText,
      scaleOcrSuggestion: line.scaleOcrSuggestion,
      scaleOcrConfidence: line.scaleOcrConfidence,
      scaleOcrReasons: line.scaleOcrReasons,
      scaleOcrError: line.scaleOcrError,
      scaleOcrRequestId: line.scaleOcrRequestId,
    }
  })),
})

const hasWasteLineValue = (line: WasteLine) =>
  Boolean(line.WasteCategory.trim() || line.Tonnage.trim())

const isWasteLineComplete = (line: WasteLine) => {
  const tonnage = Number(line.Tonnage)
  return Boolean(line.WasteCategory.trim()) && line.Tonnage.trim() !== '' && Number.isFinite(tonnage) && tonnage >= 0
}

const getCompleteWasteLines = (lines: WasteLine[]) => lines.filter(isWasteLineComplete)

const hasIncompleteWasteLine = (lines: WasteLine[]) =>
  lines.some((line) => hasWasteLineValue(line) && !isWasteLineComplete(line))

const useFilePreviewDataUrl = (file: File | null) => {
  const [previewUrl, setPreviewUrl] = useState('')

  useEffect(() => {
    let cancelled = false

    if (!file) {
      return () => {
        cancelled = true
      }
    }

    void fileToDataUrl(file)
      .then((url) => {
        if (!cancelled) setPreviewUrl(url)
      })
      .catch(() => {
        if (!cancelled) setPreviewUrl('')
      })

    return () => {
      cancelled = true
    }
  }, [file])

  return file ? previewUrl : ''
}

const STEPS = [
  { label: 'Basic Info', desc: 'Title & date', icon: I.fileTxt },
  { label: 'Customer', desc: 'Customer details', icon: I.building },
  { label: 'Assignment', desc: 'Driver & vehicle', icon: I.truck },
  { label: 'Proof', desc: 'Signature & photos', icon: I.clipList },
  { label: 'Review', desc: 'Confirm & submit', icon: I.check },
] as const

const LAST_STEP = STEPS.length - 1

const parseIsoDate = (value: string) => {
  const [year, month, day] = value.split('-').map(Number)
  if (!year || !month || !day) return null
  const parsed = new Date(year, month - 1, day)
  if (
    Number.isNaN(parsed.getTime())
    || parsed.getFullYear() !== year
    || parsed.getMonth() !== month - 1
    || parsed.getDate() !== day
  ) {
    return null
  }
  return parsed
}

const formatIsoDate = (date: Date) => {
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

const formatDateDisplay = (value: string) => {
  const parsed = parseIsoDate(value)
  if (!parsed) return ''
  return parsed.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  })
}

const isSameCalendarDay = (left: Date | null, right: Date | null) =>
  Boolean(
    left
    && right
    && left.getFullYear() === right.getFullYear()
    && left.getMonth() === right.getMonth()
    && left.getDate() === right.getDate(),
  )

/* ═══════════════════════════════════════════════════════
   Sub-components
   ═══════════════════════════════════════════════════════ */

function CustomSelect({
  id, value, onChange, disabled, options, placeholder, searchable = true, searchPlaceholder,
}: {
  id: string, value: string, onChange: (val: string) => void, disabled?: boolean,
  options: { value: string, label: string }[], placeholder?: string, searchable?: boolean, searchPlaceholder?: string
}) {
  const [isOpen, setIsOpen] = useState(false)
  const [searchTerm, setSearchTerm] = useState('')
  const [activeIndex, setActiveIndex] = useState(-1)
  const containerRef = useRef<HTMLDivElement>(null)
  const triggerRef = useRef<HTMLButtonElement>(null)
  const searchInputRef = useRef<HTMLInputElement>(null)
  const optionRefs = useRef<(HTMLButtonElement | null)[]>([])
  const listboxId = `${id}-listbox`

  useEffect(() => {
    const handleClickOutside = (event: globalThis.PointerEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false)
        setSearchTerm('')
        setActiveIndex(-1)
      }
    }
    document.addEventListener('pointerdown', handleClickOutside)
    return () => document.removeEventListener('pointerdown', handleClickOutside)
  }, [])

  const selectedOption = options.find(o => o.value === value)
  const normalizedSearch = searchTerm.trim().toLowerCase()
  const filteredOptions = normalizedSearch
    ? options.filter((option) => option.label.toLowerCase().includes(normalizedSearch))
    : options
  const getInitialActiveIndex = useCallback((nextOptions: { value: string, label: string }[]) => {
    if (nextOptions.length === 0) return -1
    const nextSelectedIndex = nextOptions.findIndex((option) => option.value === value)
    return nextSelectedIndex >= 0 ? nextSelectedIndex : 0
  }, [value])

  useEffect(() => {
    if (!isOpen) return
    window.setTimeout(() => {
      if (searchable) {
        searchInputRef.current?.focus()
        return
      }
      if (activeIndex >= 0) optionRefs.current[activeIndex]?.focus()
    }, 0)
  }, [activeIndex, isOpen, searchable])

  useEffect(() => {
    if (!isOpen || activeIndex < 0) return
    optionRefs.current[activeIndex]?.scrollIntoView({ block: 'nearest' })
  }, [activeIndex, isOpen])

  const closeMenu = useCallback((returnFocus = false) => {
    setIsOpen(false)
    setSearchTerm('')
    setActiveIndex(-1)
    if (returnFocus) window.setTimeout(() => triggerRef.current?.focus(), 0)
  }, [])

  const selectOption = useCallback((nextValue: string) => {
    onChange(nextValue)
    closeMenu(true)
  }, [closeMenu, onChange])

  const openMenu = useCallback(() => {
    if (disabled) return
    setActiveIndex(getInitialActiveIndex(filteredOptions))
    setIsOpen(true)
  }, [disabled, filteredOptions, getInitialActiveIndex])

  const handleSearchChange = (nextSearchTerm: string) => {
    const nextNormalizedSearch = nextSearchTerm.trim().toLowerCase()
    const nextFilteredOptions = nextNormalizedSearch
      ? options.filter((option) => option.label.toLowerCase().includes(nextNormalizedSearch))
      : options
    setSearchTerm(nextSearchTerm)
    setActiveIndex(getInitialActiveIndex(nextFilteredOptions))
  }

  const moveActiveIndex = useCallback((direction: 1 | -1) => {
    if (!filteredOptions.length) return
    setActiveIndex((current) => {
      if (current < 0) return direction === 1 ? 0 : filteredOptions.length - 1
      const nextIndex = current + direction
      if (nextIndex < 0) return filteredOptions.length - 1
      if (nextIndex >= filteredOptions.length) return 0
      return nextIndex
    })
  }, [filteredOptions.length])

  const handleTriggerKeyDown = (event: React.KeyboardEvent<HTMLButtonElement>) => {
    if (disabled) return
    if (event.key === 'ArrowDown') {
      event.preventDefault()
      if (!isOpen) {
        openMenu()
        return
      }
      moveActiveIndex(1)
      return
    }
    if (event.key === 'ArrowUp') {
      event.preventDefault()
      if (!isOpen) {
        openMenu()
        return
      }
      moveActiveIndex(-1)
      return
    }
    if (event.key === 'Home' && isOpen) {
      event.preventDefault()
      setActiveIndex(filteredOptions.length > 0 ? 0 : -1)
      return
    }
    if (event.key === 'End' && isOpen) {
      event.preventDefault()
      setActiveIndex(filteredOptions.length > 0 ? filteredOptions.length - 1 : -1)
      return
    }
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault()
      if (!isOpen) {
        openMenu()
        return
      }
      if (activeIndex >= 0 && filteredOptions[activeIndex]) {
        selectOption(filteredOptions[activeIndex].value)
      }
      return
    }
    if (event.key === 'Escape' && isOpen) {
      event.preventDefault()
      closeMenu(true)
    }
  }

  const handleSearchKeyDown = (event: React.KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'ArrowDown') {
      event.preventDefault()
      moveActiveIndex(1)
      return
    }
    if (event.key === 'ArrowUp') {
      event.preventDefault()
      moveActiveIndex(-1)
      return
    }
    if (event.key === 'Home') {
      event.preventDefault()
      setActiveIndex(filteredOptions.length > 0 ? 0 : -1)
      return
    }
    if (event.key === 'End') {
      event.preventDefault()
      setActiveIndex(filteredOptions.length > 0 ? filteredOptions.length - 1 : -1)
      return
    }
    if (event.key === 'Enter' && activeIndex >= 0 && filteredOptions[activeIndex]) {
      event.preventDefault()
      selectOption(filteredOptions[activeIndex].value)
      return
    }
    if (event.key === 'Escape') {
      event.preventDefault()
      closeMenu(true)
      return
    }
    if (event.key === 'Tab') {
      closeMenu(false)
    }
  }

  return (
    <div className={`cora-custom-select ${disabled ? 'disabled' : ''} ${isOpen ? 'open' : ''}`} ref={containerRef}>
      <button
        ref={triggerRef}
        type="button"
        id={id}
        className={`cora-custom-select-trigger ${!selectedOption && placeholder ? 'placeholder' : ''}`}
        onClick={() => {
          if (isOpen) {
            closeMenu(true)
            return
          }
          openMenu()
        }}
        onKeyDown={handleTriggerKeyDown}
        role="combobox"
        aria-haspopup="listbox"
        aria-expanded={isOpen}
        aria-controls={listboxId}
        aria-activedescendant={isOpen && activeIndex >= 0 ? `${id}-option-${activeIndex}` : undefined}
      >
        <span className="cora-custom-select-value">
          {selectedOption ? selectedOption.label : (placeholder || 'Select...')}
        </span>
        <svg className="cora-custom-select-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <polyline points="6 9 12 15 18 9"/>
        </svg>
      </button>
      {isOpen && (
        <ul className="cora-custom-select-menu" role="listbox" id={listboxId} aria-labelledby={id}>
          {searchable && (
            <li className="cora-custom-select-search-wrap">
              <input
                ref={searchInputRef}
                className="cora-custom-select-search"
                value={searchTerm}
                onChange={(event) => handleSearchChange(event.target.value)}
                onKeyDown={handleSearchKeyDown}
                placeholder={searchPlaceholder ?? 'Search options...'}
                aria-label={searchPlaceholder ?? 'Search options'}
                aria-controls={listboxId}
                aria-activedescendant={activeIndex >= 0 ? `${id}-option-${activeIndex}` : undefined}
              />
            </li>
          )}
          {filteredOptions.length === 0 && <li className="cora-custom-select-empty">No matching options</li>}
          {filteredOptions.map((opt, index) => (
            <li key={opt.value} role="presentation">
              <button
                ref={(element) => {
                  optionRefs.current[index] = element
                }}
                id={`${id}-option-${index}`}
                type="button"
                role="option"
                aria-selected={value === opt.value}
                tabIndex={searchable ? -1 : index === activeIndex ? 0 : -1}
                className={`cora-custom-select-option ${value === opt.value ? 'selected' : ''} ${index === activeIndex ? 'active' : ''}`}
                onMouseEnter={() => setActiveIndex(index)}
                onClick={() => selectOption(opt.value)}
              >
                {opt.label}
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}

function CustomDatePicker({ id, value, onChange, disabled }: { id: string, value: string, onChange: (val: string) => void, disabled?: boolean }) {
  const [isOpen, setIsOpen] = useState(false)
  const [viewDate, setViewDate] = useState(() => parseIsoDate(value) ?? new Date())
  const [focusedDate, setFocusedDate] = useState(() => parseIsoDate(value) ?? new Date())
  const containerRef = useRef<HTMLDivElement>(null)
  const triggerRef = useRef<HTMLButtonElement>(null)
  const dayRefs = useRef<Record<string, HTMLButtonElement | null>>({})

  useEffect(() => {
    const handleClickOutside = (event: globalThis.PointerEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false)
      }
    }
    document.addEventListener('pointerdown', handleClickOutside)
    return () => document.removeEventListener('pointerdown', handleClickOutside)
  }, [])

  const syncCalendarToValue = useCallback(() => {
    const baseDate = parseIsoDate(value) ?? new Date()
    setViewDate(new Date(baseDate.getFullYear(), baseDate.getMonth(), 1))
    setFocusedDate(baseDate)
  }, [value])

  const openCalendar = useCallback(() => {
    if (disabled) return
    syncCalendarToValue()
    setIsOpen(true)
  }, [disabled, syncCalendarToValue])

  const currentYear = viewDate.getFullYear()
  const currentMonth = viewDate.getMonth()
  const selectedDate = parseIsoDate(value)

  const startOfMonth = new Date(currentYear, currentMonth, 1)
  const endOfMonth = new Date(currentYear, currentMonth + 1, 0)

  const startDayOfWeek = startOfMonth.getDay()
  const daysInMonth = endOfMonth.getDate()

  const days: (Date | null)[] = []
  for (let i = 0; i < startDayOfWeek; i++) {
    days.push(null)
  }
  for (let i = 1; i <= daysInMonth; i++) {
    days.push(new Date(currentYear, currentMonth, i))
  }

  const handlePrevMonth = () => setViewDate(new Date(currentYear, currentMonth - 1, 1))
  const handleNextMonth = () => setViewDate(new Date(currentYear, currentMonth + 1, 1))

  const handleSelect = (d: Date) => {
    onChange(formatIsoDate(d))
    setIsOpen(false)
    window.setTimeout(() => triggerRef.current?.focus(), 0)
  }

  useEffect(() => {
    if (!isOpen) return
    const focusKey = formatIsoDate(focusedDate)
    window.setTimeout(() => dayRefs.current[focusKey]?.focus(), 0)
  }, [focusedDate, isOpen])

  const moveFocusedDate = (delta: number) => {
    const nextDate = new Date(focusedDate.getFullYear(), focusedDate.getMonth(), focusedDate.getDate() + delta)
    setFocusedDate(nextDate)
    setViewDate(new Date(nextDate.getFullYear(), nextDate.getMonth(), 1))
  }

  const jumpWithinWeek = (direction: 'start' | 'end') => {
    const currentDay = focusedDate.getDay()
    moveFocusedDate(direction === 'start' ? -currentDay : 6 - currentDay)
  }

  const changeViewedMonth = (delta: number) => {
    const nextDate = new Date(focusedDate.getFullYear(), focusedDate.getMonth() + delta, focusedDate.getDate())
    setFocusedDate(nextDate)
    setViewDate(new Date(nextDate.getFullYear(), nextDate.getMonth(), 1))
  }

  const closeCalendar = (returnFocus = false) => {
    setIsOpen(false)
    if (returnFocus) window.setTimeout(() => triggerRef.current?.focus(), 0)
  }

  const handleTriggerKeyDown = (event: React.KeyboardEvent<HTMLButtonElement>) => {
    if (disabled) return
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp' || event.key === 'Enter' || event.key === ' ') {
      event.preventDefault()
      openCalendar()
      return
    }
    if (event.key === 'Escape' && isOpen) {
      event.preventDefault()
      closeCalendar(true)
    }
  }

  const handleDayKeyDown = (event: React.KeyboardEvent<HTMLButtonElement>, date: Date) => {
    if (event.key === 'ArrowRight') {
      event.preventDefault()
      moveFocusedDate(1)
      return
    }
    if (event.key === 'ArrowLeft') {
      event.preventDefault()
      moveFocusedDate(-1)
      return
    }
    if (event.key === 'ArrowDown') {
      event.preventDefault()
      moveFocusedDate(7)
      return
    }
    if (event.key === 'ArrowUp') {
      event.preventDefault()
      moveFocusedDate(-7)
      return
    }
    if (event.key === 'Home') {
      event.preventDefault()
      jumpWithinWeek('start')
      return
    }
    if (event.key === 'End') {
      event.preventDefault()
      jumpWithinWeek('end')
      return
    }
    if (event.key === 'PageUp') {
      event.preventDefault()
      changeViewedMonth(-1)
      return
    }
    if (event.key === 'PageDown') {
      event.preventDefault()
      changeViewedMonth(1)
      return
    }
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault()
      handleSelect(date)
      return
    }
    if (event.key === 'Escape') {
      event.preventDefault()
      closeCalendar(true)
      return
    }
    if (event.key === 'Tab') {
      closeCalendar(false)
    }
  }

  return (
    <div className={`cora-custom-select ${disabled ? 'disabled' : ''} ${isOpen ? 'open' : ''}`} ref={containerRef}>
      <button
        ref={triggerRef}
        type="button"
        id={id}
        className={`cora-custom-select-trigger ${!value ? 'placeholder' : ''}`}
        onClick={() => {
          if (disabled) return
          if (isOpen) {
            closeCalendar(false)
            return
          }
          openCalendar()
        }}
        onKeyDown={handleTriggerKeyDown}
        aria-haspopup="dialog"
        aria-expanded={isOpen}
        aria-controls={`${id}-calendar`}
      >
        <span className="cora-custom-select-value">
          {formatDateDisplay(value) || 'Select date...'}
        </span>
        <svg className="cora-custom-select-icon" style={{ transform: 'none' }} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
          <line x1="16" y1="2" x2="16" y2="6" />
          <line x1="8" y1="2" x2="8" y2="6" />
          <line x1="3" y1="10" x2="21" y2="10" />
        </svg>
      </button>
      {isOpen && (
        <div className="cora-date-picker-menu" id={`${id}-calendar`} role="dialog" aria-modal="false" aria-label="Choose collection date">
          <div className="cora-date-picker-head">
            <button type="button" onClick={handlePrevMonth} className="cora-date-picker-nav" aria-label="Previous month">{I.chevL}</button>
            <div className="cora-date-picker-title">
              {viewDate.toLocaleString('default', { month: 'short', year: 'numeric' })}
            </div>
            <button type="button" onClick={handleNextMonth} className="cora-date-picker-nav" aria-label="Next month">{I.chevR}</button>
          </div>
          <div className="cora-date-picker-grid" role="grid" aria-labelledby={id}>
            {['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'].map(w => (
              <div key={w} className="cora-date-picker-dayname" role="columnheader">{w}</div>
            ))}
            {days.map((d, i) => (
              d ? (
                <button
                  key={i}
                  ref={(element) => {
                    dayRefs.current[formatIsoDate(d)] = element
                  }}
                  type="button"
                  className={`cora-date-picker-day ${isSameCalendarDay(selectedDate, d) ? 'selected' : ''} ${isSameCalendarDay(focusedDate, d) ? 'focused' : ''}`}
                  onClick={() => handleSelect(d)}
                  onFocus={() => setFocusedDate(d)}
                  onKeyDown={(event) => handleDayKeyDown(event, d)}
                  tabIndex={isSameCalendarDay(focusedDate, d) ? 0 : -1}
                  aria-selected={isSameCalendarDay(selectedDate, d)}
                  aria-label={d.toLocaleDateString('en-US', {
                    weekday: 'long',
                    year: 'numeric',
                    month: 'long',
                    day: 'numeric',
                  })}
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
  view, adminTab, onNav, onAdminNav, open, collapsed, onClose, isAdmin,
}: {
  view: View
  adminTab: AdminTab
  onNav: (v: View) => void
  onAdminNav: (tab: AdminTab) => void
  open: boolean
  collapsed: boolean
  onClose: () => void
  isAdmin: boolean
}) {
  return (
    <>
      <div className={`cora-sidebar-backdrop ${open ? 'visible' : ''}`} onClick={onClose} aria-hidden={!open} />
      <aside
        className={`cora-sidebar ${open ? 'open' : ''} ${collapsed ? 'collapsed' : ''}`}
        aria-label="Primary navigation"
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
          <button className={`cora-nav-item ${view === 'form' ? 'active' : ''}`} onClick={() => { onNav('form'); onClose() }} id="nav-new-order" aria-current={view === 'form' ? 'page' : undefined}>
            {I.plus} New Order
          </button>
          <button
            className={`cora-nav-item ${view === 'list' ? 'active' : ''}`}
            onClick={() => { onNav('list'); onClose() }}
            disabled={!isAdmin}
            id="nav-orders"
            aria-current={view === 'list' ? 'page' : undefined}
          >
            {I.clipList} All Orders
          </button>
        </div>

        <div className="cora-sidebar-section">
          <div className="cora-sidebar-section-title">Reference</div>
          <button
            className={`cora-nav-item ${view === 'admin' && adminTab === 'customers' ? 'active' : ''}`}
            onClick={() => { onAdminNav('customers'); onClose() }}
            disabled={!isAdmin}
            aria-current={view === 'admin' && adminTab === 'customers' ? 'page' : undefined}
          >
            {I.building} Customers
          </button>
          <button
            className={`cora-nav-item ${view === 'admin' && adminTab === 'drivers' ? 'active' : ''}`}
            onClick={() => { onAdminNav('drivers'); onClose() }}
            disabled={!isAdmin}
            aria-current={view === 'admin' && adminTab === 'drivers' ? 'page' : undefined}
          >
            {I.user} Drivers
          </button>
          <button
            className={`cora-nav-item ${view === 'admin' && adminTab === 'vehicles' ? 'active' : ''}`}
            onClick={() => { onAdminNav('vehicles'); onClose() }}
            disabled={!isAdmin}
            aria-current={view === 'admin' && adminTab === 'vehicles' ? 'page' : undefined}
          >
            {I.truck} Vehicles
          </button>
          <button
            className={`cora-nav-item ${view === 'admin' && adminTab === 'waste' ? 'active' : ''}`}
            onClick={() => { onAdminNav('waste'); onClose() }}
            disabled={!isAdmin}
            aria-current={view === 'admin' && adminTab === 'waste' ? 'page' : undefined}
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
  useEffect(() => {
    const t = setTimeout(onClose, toast.durationMs ?? (toast.actionLabel ? 7000 : 4200))
    return () => clearTimeout(t)
  }, [onClose, toast.actionLabel, toast.durationMs])
  return (
    <div className="cora-toast-wrap">
      <div
        className={`cora-toast ${toast.type}`}
        role={toast.type === 'error' ? 'alert' : 'status'}
        aria-live={toast.type === 'error' ? 'assertive' : 'polite'}
        aria-atomic="true"
      >
        <span className="cora-toast-icon">{toast.type === 'success' ? '✓' : '✕'}</span>
        <span className="cora-toast-text">{toast.text}</span>
        {toast.actionLabel && toast.onAction && (
          <button
            className="cora-toast-action"
            type="button"
            onClick={() => {
              toast.onAction?.()
              onClose()
            }}
          >
            {toast.actionLabel}
          </button>
        )}
        <button className="cora-toast-close" onClick={onClose} aria-label="Dismiss">×</button>
      </div>
    </div>
  )
}

/* ── Orders Table ── */
function OrdersTable({ orders, loadState }: { orders: ServiceOrdersRead[]; loadState: Load }) {
  const [search, setSearch] = useState('')
  const [selectedDate, setSelectedDate] = useState('')

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
  const hasSearch = Boolean(normalizedSearch)
  const hasDateFilter = Boolean(selectedDate)
  const filtered = orders.filter((o) => {
    const matchesSearch = !hasSearch || cols.some((c) => String(o[c] ?? '').toLowerCase().includes(normalizedSearch))
    const matchesDate = !hasDateFilter || String(o.DateOfCollection ?? '').slice(0, 10) === selectedDate
    return matchesSearch && matchesDate
  })

  const formatDateDisplay = (val: string) => {
    if (!val) return ''
    const [y, m, d] = val.split('-')
    if (!y || !m || !d) return ''
    return `${d}/${m}/${y}`
  }

  const activeFilterLabel = [
    hasSearch ? 'search' : '',
    hasDateFilter ? `date ${formatDateDisplay(selectedDate)}` : '',
  ].filter(Boolean).join(' and ')

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
          <div className="cora-table-date-filter">
            <CustomDatePicker
              id="orders-date-filter"
              value={selectedDate}
              onChange={setSelectedDate}
            />
            {hasDateFilter && (
              <button className="cora-table-date-clear" onClick={() => setSelectedDate('')} aria-label="Clear date filter">
                Clear date
              </button>
            )}
          </div>
          {(hasSearch || hasDateFilter) && (
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
          <span>No orders match {activeFilterLabel || 'the current filters'}.</span>
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
  activeTab, onTabChange, customerAreas, drivers, vehicles, categories, busy, currentUserEmail,
  onAddCustomer, onAddDriver, onAddVehicle, onAddWasteCategory,
  onUpdateCustomer, onUpdateDriver, onUpdateVehicle, onUpdateWasteCategory,
  onDeleteCustomer, onDeleteDriver, onDeleteVehicle, onDeleteWasteCategory,
}: {
  activeTab: AdminTab
  onTabChange: (tab: AdminTab) => void
  customerAreas: CustomerAreaOption[]
  drivers: Drivers1Read[]
  vehicles: VehiclesRead[]
  categories: Waste_CategoriesRead[]
  busy: boolean
  currentUserEmail: string
  onAddCustomer: (value: CustomerAdminForm) => Promise<void>
  onAddDriver: (name: string) => Promise<void>
  onAddVehicle: (vehicleNumber: string) => Promise<void>
  onAddWasteCategory: (category: string) => Promise<void>
  onUpdateCustomer: (id: number, value: CustomerAdminForm) => Promise<void>
  onUpdateDriver: (id: number, name: string) => Promise<void>
  onUpdateVehicle: (id: number, vehicleNumber: string) => Promise<void>
  onUpdateWasteCategory: (id: number, category: string) => Promise<void>
  onDeleteCustomer: (id: number, name: string) => Promise<void>
  onDeleteDriver: (id: number, name: string) => Promise<void>
  onDeleteVehicle: (id: number, vehicleNumber: string) => Promise<void>
  onDeleteWasteCategory: (id: number, category: string) => Promise<void>
}) {
  const [customerSearch, setCustomerSearch] = useState('')
  const [driverName, setDriverName] = useState('')
  const [vehicleNumber, setVehicleNumber] = useState('')
  const [wasteCategory, setWasteCategory] = useState('')
  const [saving, setSaving] = useState<AdminTab | null>(null)
  const [editing, setEditing] = useState<{ tab: AdminTab; id: number; value: string } | null>(null)
  const [editingError, setEditingError] = useState('')
  const [simpleErrors, setSimpleErrors] = useState<Partial<Record<'drivers' | 'vehicles' | 'waste', string>>>({})
  const [customerModal, setCustomerModal] = useState<{
    mode: 'create' | 'edit'
    id?: number
    value: CustomerAdminForm
    errors: CustomerAdminErrors
  } | null>(null)
  const [deleteTarget, setDeleteTarget] = useState<DeleteTarget | null>(null)
  const [busyItem, setBusyItem] = useState<string | null>(null)

  const validateSimpleValue = useCallback((tab: 'drivers' | 'vehicles' | 'waste', rawValue: string, editingId?: number) => {
    const cleaned = rawValue.trim()
    if (!cleaned) {
      return tab === 'drivers'
        ? 'Driver name is required.'
        : tab === 'vehicles'
          ? 'Vehicle number is required.'
          : 'Waste category is required.'
    }

    if (tab === 'drivers' && drivers.some((driver) => driver.ID !== editingId && normalize(getDriverName(driver)) === normalize(cleaned))) {
      return `"${cleaned}" already exists in drivers.`
    }
    if (tab === 'vehicles' && vehicles.some((vehicle) => vehicle.ID !== editingId && normalize(vehicle.Title || '') === normalize(cleaned))) {
      return `"${cleaned}" already exists in vehicles.`
    }
    if (tab === 'waste' && categories.some((category) => category.ID !== editingId && normalize(category.Title || '') === normalize(cleaned))) {
      return `"${cleaned}" already exists in waste categories.`
    }
    return ''
  }, [categories, drivers, vehicles])

  const submit = async (tab: Exclude<AdminTab, 'customers'>) => {
    const values = {
      drivers: driverName,
      vehicles: vehicleNumber,
      waste: wasteCategory,
    }
    const value = values[tab]
    const error = validateSimpleValue(tab, value)
    if (error) {
      setSimpleErrors((current) => ({ ...current, [tab]: error }))
      return
    }

    setSaving(tab)
    try {
      if (tab === 'drivers') {
        await onAddDriver(value.trim())
        setDriverName('')
      } else if (tab === 'vehicles') {
        await onAddVehicle(value.trim())
        setVehicleNumber('')
      } else {
        await onAddWasteCategory(value.trim())
        setWasteCategory('')
      }
      setSimpleErrors((current) => ({ ...current, [tab]: '' }))
    } finally {
      setSaving(null)
    }
  }

  const tabs: { key: AdminTab; label: string; count: number; icon: ReactNode }[] = [
    { key: 'customers', label: 'Customers', count: customerAreas.length, icon: I.building },
    { key: 'drivers', label: 'Drivers', count: drivers.length, icon: I.user },
    { key: 'vehicles', label: 'Vehicles', count: vehicles.length, icon: I.truck },
    { key: 'waste', label: 'Waste Categories', count: categories.length, icon: I.recycle },
  ]
  const activeItems = activeTab === 'drivers'
    ? drivers.map((driver) => ({ id: driver.ID, value: getDriverName(driver) }))
    : activeTab === 'vehicles'
      ? vehicles.map((vehicle) => ({ id: vehicle.ID, value: vehicle.Title?.trim() || '' }))
      : activeTab === 'waste'
        ? categories.map((category) => ({ id: category.ID, value: category.Title?.trim() || '' }))
        : []

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
  } as const
  const activeSimpleConfig = activeTab === 'customers' ? null : activeConfig[activeTab]
  const normalizedCustomerSearch = customerSearch.trim().toLowerCase()
  const customerItems = customerAreas
    .filter((item): item is CustomerAreaOption & { ID: number } => Boolean(item.ID && getCustomerName(item)))
    .sort((a, b) => getCustomerName(a).localeCompare(getCustomerName(b)))
  const filteredCustomers = customerItems.filter((item) => {
    if (!normalizedCustomerSearch) return true
    const haystack = `${getCustomerName(item)} ${getCustomerAdminSummary(item)}`.toLowerCase()
    return haystack.includes(normalizedCustomerSearch)
  })

  const openCustomerModal = (mode: 'create' | 'edit', item?: CustomerAreaOption & { ID: number }) => {
    setCustomerModal({
      mode,
      id: item?.ID,
      value: item ? getCustomerAdminFormFromRow(item) : createEmptyCustomerAdminForm(),
      errors: {},
    })
  }

  const setCustomerModalField = (key: CustomerAdminFieldKey, nextValue: string) => {
    setCustomerModal((current) => (
      current
        ? {
          ...current,
          value: { ...current.value, [key]: nextValue },
          errors: { ...current.errors, [key]: '' },
        }
        : current
    ))
  }

  const submitCustomerModal = async () => {
    if (!customerModal) return
    const errors = getCustomerAdminErrors(customerModal.value, customerAreas, customerModal.id)
    if (Object.keys(errors).length > 0) {
      setCustomerModal((current) => current ? { ...current, errors } : current)
      return
    }

    setSaving('customers')
    try {
      if (customerModal.mode === 'create') {
        await onAddCustomer(normalizeCustomerAdminForm(customerModal.value))
      } else if (customerModal.id) {
        await onUpdateCustomer(customerModal.id, normalizeCustomerAdminForm(customerModal.value))
      }
      setCustomerModal(null)
    } finally {
      setSaving(null)
    }
  }

  const updateItem = async (id: number, value: string) => {
    const cleaned = value.trim()
    const error = validateSimpleValue(activeTab as 'drivers' | 'vehicles' | 'waste', cleaned, id)
    if (error) {
      setEditingError(error)
      return
    }
    const key = `update-${activeTab}-${id}`
    setBusyItem(key)
    try {
      if (activeTab === 'drivers') await onUpdateDriver(id, cleaned)
      else if (activeTab === 'vehicles') await onUpdateVehicle(id, cleaned)
      else await onUpdateWasteCategory(id, cleaned)
      setEditing(null)
      setEditingError('')
    } finally {
      setBusyItem(null)
    }
  }

  const confirmDelete = async () => {
    if (!deleteTarget) return
    const key = `delete-${deleteTarget.tab}-${deleteTarget.id}`
    setBusyItem(key)
    try {
      if (deleteTarget.tab === 'customers') await onDeleteCustomer(deleteTarget.id, deleteTarget.name)
      else if (deleteTarget.tab === 'drivers') await onDeleteDriver(deleteTarget.id, deleteTarget.name)
      else if (deleteTarget.tab === 'vehicles') await onDeleteVehicle(deleteTarget.id, deleteTarget.name)
      else await onDeleteWasteCategory(deleteTarget.id, deleteTarget.name)
      if (editing?.id === deleteTarget.id && editing.tab === deleteTarget.tab) {
        setEditing(null)
        setEditingError('')
      }
      setDeleteTarget(null)
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
                setEditingError('')
                setSimpleErrors({})
                onTabChange(tab.key)
              }}
              disabled={busy || Boolean(saving) || Boolean(busyItem)}
            >
              {tab.icon} {tab.label} <span>{tab.count}</span>
            </button>
          ))}
        </div>

        {activeTab === 'customers' ? (
          <>
            <div className="cora-form-grid cora-admin-form cora-admin-customer-manage">
              <div className="cora-section-divider"><span>Manage Customers</span></div>
              <div className="cora-field span">
                <p className="cora-field-hint">Use search to find customers quickly. Open the customer editor to add or update the hierarchy fields that drive the order form.</p>
              </div>
              <div className="cora-admin-toolbar">
                <div className="cora-table-search-wrap">
                  <svg className="cora-table-search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" />
                  </svg>
                  <input
                    className="cora-table-search"
                    value={customerSearch}
                    onChange={(event) => setCustomerSearch(event.target.value)}
                    placeholder="Search customers or locations..."
                    aria-label="Search customers"
                    disabled={busy || Boolean(busyItem)}
                  />
                  {customerSearch && (
                    <button className="cora-table-search-clear" type="button" onClick={() => setCustomerSearch('')} aria-label="Clear customer search">
                      {I.x}
                    </button>
                  )}
                </div>
                <div className="cora-admin-toolbar-actions">
                  <span className="cora-table-filter-badge">{filteredCustomers.length} of {customerItems.length} shown</span>
                  <button
                    type="button"
                    className="cora-btn cora-btn-primary"
                    onClick={() => openCustomerModal('create')}
                    disabled={busy || Boolean(saving) || Boolean(busyItem)}
                  >
                    {I.plus} Add Customer
                  </button>
                </div>
              </div>
            </div>
          </>
        ) : (
          <div className="cora-form-grid cora-admin-form">
            <div className="cora-section-divider"><span>{activeSimpleConfig?.title}</span></div>
            <div className="cora-field">
              <label className="cora-label" htmlFor={`admin-${activeTab}`}>{activeSimpleConfig?.label}</label>
              <input
                id={`admin-${activeTab}`}
                className={`cora-input ${simpleErrors[activeTab] ? 'invalid' : ''}`}
                value={activeSimpleConfig?.value ?? ''}
                onChange={(event) => {
                  activeSimpleConfig?.onChange(event.target.value)
                  setSimpleErrors((current) => ({ ...current, [activeTab]: '' }))
                }}
                onBlur={() => {
                  const nextError = validateSimpleValue(activeTab, activeSimpleConfig?.value ?? '')
                  setSimpleErrors((current) => ({ ...current, [activeTab]: nextError }))
                }}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault()
                    void submit(activeTab)
                  }
                }}
                placeholder={activeSimpleConfig?.placeholder}
                disabled={busy || Boolean(saving)}
              />
              {simpleErrors[activeTab] && <p className="cora-field-error">{simpleErrors[activeTab]}</p>}
            </div>
            <div className="cora-field cora-admin-submit-field">
              <label className="cora-label" aria-hidden="true">&nbsp;</label>
              <button
                type="button"
                className="cora-btn cora-btn-primary"
                onClick={() => void submit(activeTab)}
                disabled={busy || Boolean(saving)}
              >
                {saving === activeTab ? <><span className="cora-spinner" /> Saving</> : <>{I.plus} Add</>}
              </button>
            </div>
          </div>
        )}

        <div className="cora-admin-list">
          <div className="cora-section-divider"><span>Current {tabs.find((tab) => tab.key === activeTab)?.label}</span></div>
          {activeTab === 'customers' ? (
            customerItems.length === 0 ? (
              <div className="cora-table-empty">No customers have been added yet. Use the form above to add one.</div>
            ) : filteredCustomers.length === 0 ? (
              <div className="cora-table-empty">No customers match "{customerSearch.trim()}". Try a different name or location.</div>
            ) : (
              <div className="cora-admin-crud-list">
                {filteredCustomers.map((item) => {
                  const itemName = getCustomerName(item)
                  return (
                    <div className="cora-admin-crud-item cora-admin-customer-item" key={`customers-${item.ID}`}>
                      <div className="cora-admin-customer-summary">
                        <span className="cora-admin-crud-name">{itemName}</span>
                        <span className="cora-admin-customer-meta">{getCustomerAdminSummary(item) || 'No location details'}</span>
                      </div>
                      <div className="cora-admin-crud-actions">
                        <button
                          type="button"
                          className="cora-btn cora-btn-outline"
                          onClick={() => openCustomerModal('edit', item)}
                          disabled={Boolean(busyItem)}
                        >
                          Edit
                        </button>
                        <button
                          type="button"
                          className="cora-btn cora-btn-danger"
                          onClick={() => setDeleteTarget({ tab: 'customers', id: item.ID, name: itemName })}
                          disabled={Boolean(busyItem)}
                        >
                          {busyItem === `delete-customers-${item.ID}` ? <><span className="cora-spinner" /> Deleting</> : 'Delete'}
                        </button>
                      </div>
                    </div>
                  )
                })}
              </div>
            )
          ) : activeItems.filter((item) => item.id && item.value).length === 0 ? (
            <div className="cora-table-empty">No {tabs.find((tab) => tab.key === activeTab)?.label.toLowerCase()} have been added yet. Use the form above to add one.</div>
          ) : (
            <div className="cora-admin-crud-list">
              {activeItems.filter((item): item is { id: number; value: string } => Boolean(item.id && item.value)).map((item) => (
                <div className="cora-admin-crud-item" key={`${activeTab}-${item.id}`}>
                  {editing?.tab === activeTab && editing.id === item.id ? (
                    <div className="cora-admin-inline-edit">
                      <input
                        className={`cora-input cora-admin-edit-input ${editingError ? 'invalid' : ''}`}
                        value={editing.value}
                        onChange={(event) => {
                          setEditing({ ...editing, value: event.target.value })
                          setEditingError('')
                        }}
                        onBlur={() => setEditingError(validateSimpleValue(activeTab, editing.value, item.id))}
                        onKeyDown={(event) => {
                          if (event.key === 'Enter') {
                            event.preventDefault()
                            void updateItem(item.id, editing.value)
                          }
                          if (event.key === 'Escape') {
                            setEditing(null)
                            setEditingError('')
                          }
                        }}
                        disabled={Boolean(busyItem)}
                        autoFocus
                      />
                      {editingError && <p className="cora-field-error">{editingError}</p>}
                    </div>
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
                          disabled={Boolean(busyItem)}
                        >
                          {busyItem === `update-${activeTab}-${item.id}` ? <><span className="cora-spinner" /> Saving</> : 'Save'}
                        </button>
                        <button
                          type="button"
                          className="cora-btn cora-btn-outline"
                          onClick={() => {
                            setEditing(null)
                            setEditingError('')
                          }}
                          disabled={Boolean(busyItem)}
                        >
                          Cancel
                        </button>
                      </>
                    ) : (
                      <>
                        <button
                          type="button"
                          className="cora-btn cora-btn-outline"
                          onClick={() => {
                            setEditing({ tab: activeTab, id: item.id, value: item.value })
                            setEditingError('')
                          }}
                          disabled={Boolean(busyItem)}
                        >
                          Edit
                        </button>
                        <button
                          type="button"
                          className="cora-btn cora-btn-danger"
                          onClick={() => setDeleteTarget({ tab: activeTab, id: item.id, name: item.value })}
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
      {customerModal && (
        <CustomerEditorModal
          mode={customerModal.mode}
          value={customerModal.value}
          errors={customerModal.errors}
          busy={saving === 'customers'}
          onChange={setCustomerModalField}
          onCancel={() => { if (saving !== 'customers') setCustomerModal(null) }}
          onSubmit={() => void submitCustomerModal()}
        />
      )}
      {deleteTarget && (
        <ConfirmDeleteModal
          target={deleteTarget}
          busy={busyItem === `delete-${deleteTarget.tab}-${deleteTarget.id}`}
          onCancel={() => { if (!busyItem) setDeleteTarget(null) }}
          onConfirm={() => void confirmDelete()}
        />
      )}
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
  form, set, busy, customerAreas,
}: {
  form: Form
  set: <K extends keyof Form>(k: K, v: Form[K]) => void
  busy: boolean
  customerAreas: CustomerAreaOption[]
}) {
  const levelOptions = CUSTOMER_LEVELS.map((_, levelIndex) => getCustomerLevelOptions(customerAreas, form, levelIndex))
  const customerPathOptions = getCustomerPathOptions(customerAreas)
  const visibleLevels = CUSTOMER_LEVELS.filter((_, levelIndex) => levelIndex === 0 || levelOptions[levelIndex].length > 0)
  const selectedCustomerPath = customerPathOptions.some((option) => option.value === getFormCustomerPathValue(form))
    ? getFormCustomerPathValue(form)
    : ''

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

  const chooseCustomerPath = (value: string) => {
    const option = customerPathOptions.find((item) => item.value === value)
    if (!option) return
    CUSTOMER_LEVELS.forEach((level, levelIndex) => {
      set(level.key, getCustomerLevelValue(option.row, levelIndex))
    })
    set('Customer', getCustomerName(option.row))
    set('IsAdhocCustomer', false)
  }

  return (
    <div className="cora-form-grid" key="step-customer">
      <div className="cora-section-divider"><span>Customer Details</span></div>
      <div className="cora-field span">
        <p className="cora-field-hint">
          Follow the customer hierarchy shown here so the order is linked to the correct site and sub-location.
        </p>
      </div>
      {visibleLevels.map((level) => {
        const levelIndex = CUSTOMER_LEVELS.indexOf(level)
        if (levelIndex === 0) {
          return (
            <div className="cora-field" key={level.key}>
              <label className="cora-label" htmlFor="f-customer-level-1">{level.label}</label>
              <CustomSelect
                id="f-customer-level-1"
                value={selectedCustomerPath}
                onChange={chooseCustomerPath}
                disabled={busy}
                placeholder={level.placeholder}
                searchable
                searchPlaceholder="Search any customer level..."
                options={customerPathOptions.map((option) => ({ value: option.value, label: option.label }))}
              />
            </div>
          )
        }
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

function CustomerEditorModal({
  mode, value, errors, busy, onChange, onCancel, onSubmit,
}: {
  mode: 'create' | 'edit'
  value: CustomerAdminForm
  errors: CustomerAdminErrors
  busy: boolean
  onChange: (key: CustomerAdminFieldKey, nextValue: string) => void
  onCancel: () => void
  onSubmit: () => void
}) {
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !busy) onCancel()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [busy, onCancel])

  return (
    <div className="cora-modal-backdrop" role="presentation" onMouseDown={() => { if (!busy) onCancel() }}>
      <div className="cora-modal cora-admin-customer-modal" role="dialog" aria-modal="true" aria-labelledby="customer-editor-title" onMouseDown={(event) => event.stopPropagation()}>
        <div className="cora-modal-head">
          <div>
            <h2 id="customer-editor-title">{mode === 'create' ? 'Add Customer' : 'Edit Customer'}</h2>
            <p>Manage the customer hierarchy used by the service order wizard. Customer name and location are required.</p>
          </div>
          <button className="cora-modal-close" type="button" onClick={onCancel} aria-label="Close customer editor" disabled={busy}>
            {I.x}
          </button>
        </div>
        <div className="cora-modal-body">
          <div className="cora-form-grid cora-admin-customer-modal-grid">
            <div className="cora-section-divider"><span>Identity</span></div>
            {CUSTOMER_ADMIN_FIELDS.slice(0, 3).map((field) => (
              <div className={`cora-field ${field.key === 'field_2' ? 'span' : ''}`} key={`customer-modal-${field.key}`}>
                <label className="cora-label" htmlFor={`customer-modal-${field.key}`}>
                  {field.label} {field.required ? <span className="req">*</span> : null}
                </label>
                <input
                  id={`customer-modal-${field.key}`}
                  className={`cora-input ${errors[field.key] ? 'invalid' : ''}`}
                  value={value[field.key]}
                  onChange={(event) => onChange(field.key, event.target.value)}
                  placeholder={field.placeholder}
                  disabled={busy}
                  autoFocus={field.key === 'Title'}
                />
                {field.key === 'field_2' && <p className="cora-field-hint">Use `/` to separate deeper location steps when needed, for example `Tower A / Level 2`.</p>}
                {errors[field.key] && <p className="cora-field-error">{errors[field.key]}</p>}
              </div>
            ))}
            <div className="cora-section-divider"><span>Optional Hierarchy</span></div>
            {CUSTOMER_ADMIN_FIELDS.slice(3).map((field) => (
              <div className="cora-field" key={`customer-optional-${field.key}`}>
                <label className="cora-label" htmlFor={`customer-modal-${field.key}`}>{field.label}</label>
                <input
                  id={`customer-modal-${field.key}`}
                  className={`cora-input ${errors[field.key] ? 'invalid' : ''}`}
                  value={value[field.key]}
                  onChange={(event) => onChange(field.key, event.target.value)}
                  placeholder={field.placeholder}
                  disabled={busy}
                />
                {errors[field.key] && <p className="cora-field-error">{errors[field.key]}</p>}
              </div>
            ))}
          </div>
        </div>
        <div className="cora-modal-foot">
          <button className="cora-btn cora-btn-outline" type="button" onClick={onCancel} disabled={busy}>Cancel</button>
          <button className="cora-btn cora-btn-primary" type="button" onClick={onSubmit} disabled={busy}>
            {busy ? <><span className="cora-spinner" /> Saving</> : mode === 'create' ? 'Save Customer' : 'Update Customer'}
          </button>
        </div>
      </div>
    </div>
  )
}

function ConfirmDeleteModal({
  target, busy, onCancel, onConfirm,
}: {
  target: DeleteTarget
  busy: boolean
  onCancel: () => void
  onConfirm: () => void
}) {
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !busy) onCancel()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [busy, onCancel])

  const targetLabel = target.tab === 'waste'
    ? 'waste category'
    : target.tab === 'vehicles'
      ? 'vehicle'
      : target.tab === 'drivers'
        ? 'driver'
        : 'customer'

  return (
    <div className="cora-modal-backdrop" role="presentation" onMouseDown={() => { if (!busy) onCancel() }}>
      <div className="cora-modal" role="dialog" aria-modal="true" aria-labelledby="confirm-delete-title" onMouseDown={(event) => event.stopPropagation()} style={{ width: 'min(420px, 100%)' }}>
        <div className="cora-modal-head">
          <div>
            <h2 id="confirm-delete-title" style={{ color: 'var(--error)' }}>Delete {targetLabel}?</h2>
            <p>Delete "{target.name}" from the admin reference data. You can undo this shortly after deletion.</p>
          </div>
          <button className="cora-modal-close" type="button" onClick={onCancel} aria-label="Cancel delete" disabled={busy}>
            {I.x}
          </button>
        </div>
        <div className="cora-modal-foot">
          <button className="cora-btn cora-btn-outline" type="button" onClick={onCancel} disabled={busy}>Cancel</button>
          <button className="cora-btn cora-btn-danger" type="button" onClick={onConfirm} disabled={busy}>
            {busy ? <><span className="cora-spinner" /> Deleting</> : `Delete ${targetLabel}`}
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

  const pointerToFrame = (event: ReactPointerEvent<HTMLElement>) => {
    const canvas = canvasRef.current
    if (!canvas) return null
    const rect = canvas.getBoundingClientRect()
    return {
      x: (event.clientX - rect.left) / rect.width,
      y: (event.clientY - rect.top) / rect.height,
    }
  }

  const startInteraction = (event: ReactPointerEvent<HTMLElement>, mode: 'move' | 'resize') => {
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

  const startDraw = (event: ReactPointerEvent<HTMLDivElement>) => {
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

  const moveInteraction = (event: ReactPointerEvent<HTMLDivElement>) => {
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

  const endInteraction = (event: ReactPointerEvent<HTMLDivElement>) => {
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
    const nextLines = form.WasteItems.filter((line) => line.id !== id)
    set('WasteItems', nextLines.length > 0 ? nextLines : [createWasteLine()])
  }

  const clearScalePhoto = (id: string) => {
    updateForm((current) => ({
      ...current,
      WasteItems: current.WasteItems.map((line) => {
        if (line.id !== id) return line
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

  const runOcrForLine = async (id: string, imageSource: string, requestId: string, crop?: CropRect) => {
    try {
      const preprocessed = await preprocessImageForOcr(imageSource, crop)
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
    const previewUrl = await fileToDataUrl(file)
    updateForm((current) => ({
      ...current,
      WasteItems: current.WasteItems.map((line) => {
        if (line.id !== id) return line
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

  const applyScaleCrop = (id: string, crop: CropRect, imageSource: string) => {
    const line = form.WasteItems.find((item) => item.id === id)
    if (!line) return
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
    void runOcrForLine(id, imageSource, requestId, crop)
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
          onApply={(crop) => applyScaleCrop(cropEditingLine.id, crop, cropEditingLine.scalePhotoPreviewUrl!)}
        />
      )}
      <div className="cora-section-divider"><span>Resource Assignment</span></div>
      <div className="cora-field span">
        <p className="cora-field-hint">
          Driver and vehicle are optional, but each waste line should include a category and tonnage before you continue.
        </p>
      </div>
      <div className="cora-field">
        <label className="cora-label" htmlFor="f-driver">Driver</label>
        <CustomSelect
          id="f-driver"
          value={form.DriverName}
          onChange={(val) => set('DriverName', val)}
          disabled={busy}
          placeholder="Select driver…"
          searchPlaceholder="Search drivers..."
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
          searchPlaceholder="Search vehicles..."
          options={vehicles.map((v) => ({ value: v.Title ?? '', label: v.Title ?? `Vehicle ${v.ID}` }))}
        />
      </div>

      <div className="cora-section-divider"><span>Waste Details</span></div>
      <div className="cora-field span">
        <div className="cora-waste-lines">
          {form.WasteItems.map((line, index) => (
            <div className="cora-waste-line" key={line.id}>
              <div className="cora-waste-line-head">
                <strong>Waste item {index + 1}</strong>
                <span>Capture the material, tonnage, and optional scale evidence for this line.</span>
              </div>
              <div className="cora-field">
                <label className="cora-label" htmlFor={`f-waste-category-${line.id}`}>Waste Category</label>
                <CustomSelect
                  id={`f-waste-category-${line.id}`}
                  value={line.WasteCategory}
                  onChange={(val) => updateWasteLine(line.id, { WasteCategory: val })}
                  disabled={busy}
                  placeholder="Select waste category..."
                  searchPlaceholder="Search waste categories..."
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
                  inputMode="decimal"
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
                    {line.scalePhotoPreviewUrl ? 'Replace scale photo' : 'Take or choose scale photo'}
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

  const pointFromEvent = (event: ReactPointerEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current
    if (!canvas) return null
    const rect = canvas.getBoundingClientRect()
    return { x: event.clientX - rect.left, y: event.clientY - rect.top }
  }

  const beginDrawing = (event: ReactPointerEvent<HTMLCanvasElement>) => {
    if (disabled) return
    const point = pointFromEvent(event)
    if (!point) return
    event.currentTarget.setPointerCapture(event.pointerId)
    drawingRef.current = true
    hasDrawnRef.current = false
    strokeDistanceRef.current = 0
    lastPointRef.current = point
  }

  const draw = (event: ReactPointerEvent<HTMLCanvasElement>) => {
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
          <span className="cora-photo-dropzone-label">Tap to take or choose a photo</span>
          <span className="cora-photo-dropzone-hint">Use the camera on mobile, or drag and drop an image on desktop.</span>
        </label>
      )}
      {file && <span className="cora-photo-name">{file.name}</span>}
      <div className="cora-photo-actions">
        <label className={`cora-btn cora-btn-secondary ${disabled ? 'disabled' : ''}`} htmlFor={id}>
          {file ? 'Replace photo' : 'Take or choose photo'}
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
        <p className="cora-field-hint">The service order cannot be submitted until the signature is captured.</p>
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
        <p className="cora-field-hint">Optional. Helpful when you need a record of the area before collection.</p>
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
        <p className="cora-field-hint">Optional. Helpful when you need a record of the area after collection.</p>
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
          <table className="cora-review-table-summary">
            <tbody>
              {rows.map(([k, v]) => (
                <tr key={k}>
                  <td className="cora-review-cell-label">{k}</td>
                  <td className="cora-review-cell-value">{v}</td>
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
              <table className="cora-review-table-waste">
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
          {!beforePhotoPreviewUrl && !afterPhotoPreviewUrl && !signatureDataUrl && (
            <div className="cora-proof-preview-empty">No proof captured yet.</div>
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
  const [usingCachedReferenceData, setUsingCachedReferenceData] = useState(false)
  const [cacheFetchedAt, setCacheFetchedAt] = useState('')

  const [form, setForm] = useState<Form>(() => createBlankForm())
  const [draftClientSubmissionId, setDraftClientSubmissionId] = useState(() => createClientSubmissionId())
  const [step, setStep] = useState(0)
  const [submitting, setSubmitting] = useState(false)
  const [submitted, setSubmitted] = useState(false)
  const [submissionMode, setSubmissionMode] = useState<SubmissionMode>('draft')
  const [toast, setToast] = useState<Toast | null>(null)
  const [resetModalOpen, setResetModalOpen] = useState(false)
  const [signatureDataUrl, setSignatureDataUrl] = useState('')
  const [beforePhotoFile, setBeforePhotoFile] = useState<File | null>(null)
  const [afterPhotoFile, setAfterPhotoFile] = useState<File | null>(null)
  const [draftReady, setDraftReady] = useState(false)
  const [isOnline, setIsOnline] = useState(() => window.navigator.onLine)
  const [queuePendingCount, setQueuePendingCount] = useState(0)
  const [queueLastError, setQueueLastError] = useState('')
  const [queueSyncing, setQueueSyncing] = useState(false)

  const refreshQueueState = useCallback(async () => {
    const summary = await getQueueSummary()
    setQueuePendingCount(summary.pendingCount)
    setQueueLastError(summary.lastError)
  }, [])

  const syncQueuedSubmissions = useCallback(async () => {
    if (!window.navigator.onLine) {
      await refreshQueueState()
      return
    }
    setQueueSyncing(true)
    try {
      await processPendingQueue({
        onSubmissionSynced: (item) => {
          if (item.clientSubmissionId === draftClientSubmissionId && submissionMode === 'queued') {
            setForm(createBlankForm())
            setSignatureDataUrl('')
            setBeforePhotoFile(null)
            setAfterPhotoFile(null)
            setStep(0)
            setSubmitted(false)
            setSubmissionMode('draft')
            setDraftClientSubmissionId(createClientSubmissionId())
            void clearDraft()
            setToast({ type: 'success', text: `Queued order "${item.orderTitle}" synced successfully.` })
          }
        },
        onSubmissionFailed: (_, error) => {
          setToast({ type: 'error', text: error })
        },
        onQueueUpdated: () => {
          void refreshQueueState()
        },
      })
    } finally {
      setQueueSyncing(false)
      await refreshQueueState()
    }
  }, [draftClientSubmissionId, refreshQueueState, submissionMode])

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
      setUsingCachedReferenceData(false)
      const cacheRecord: CachedReferenceData = {
        id: 'main',
        fetchedAt: new Date().toISOString(),
        customerAreas: c,
        drivers: d.data && Array.isArray(d.data) ? d.data : [],
        vehicles: v.data && Array.isArray(v.data) ? v.data : [],
        categories: w.data && Array.isArray(w.data) ? w.data : [],
      }
      setCacheFetchedAt(cacheRecord.fetchedAt)
      await writeCachedReferenceData(cacheRecord)
      setLoadState('loaded')
    } catch (error) {
      const cached = await readCachedReferenceData()
      if (cached) {
        setCustomerAreas(cached.customerAreas)
        setDrivers(cached.drivers)
        setVehicles(cached.vehicles)
        setCategories(cached.categories)
        setOrders([])
        setUsingCachedReferenceData(true)
        setCacheFetchedAt(cached.fetchedAt)
        setLoadError(error instanceof Error ? error.message : 'Could not load live reference data.')
        setLoadState('loaded')
        return
      }
      setLoadError(error instanceof Error ? error.message : 'Could not load reference data.')
      setLoadState('error')
    }
  }, [])

  useEffect(() => {
    void loadReferenceData()
  }, [loadReferenceData])

  useEffect(() => {
    let cancelled = false
    const hydrateOfflineState = async () => {
      try {
        const [draft, cachedReferences] = await Promise.all([loadDraft(), readCachedReferenceData(), refreshQueueState()])
        if (cancelled) return
        if (cachedReferences) {
          setCustomerAreas(cachedReferences.customerAreas)
          setDrivers(cachedReferences.drivers)
          setVehicles(cachedReferences.vehicles)
          setCategories(cachedReferences.categories)
          setUsingCachedReferenceData(true)
          setCacheFetchedAt(cachedReferences.fetchedAt)
          setLoadState('loaded')
        }
        if (draft) {
          setForm(await restorePersistedForm(draft.payload.form))
          setSignatureDataUrl(draft.payload.signatureDataUrl)
          setBeforePhotoFile(await fromStoredMediaFile(draft.payload.beforePhoto))
          setAfterPhotoFile(await fromStoredMediaFile(draft.payload.afterPhoto))
          setStep(draft.step)
          setSubmitted(draft.submissionStatus === 'queued')
          setSubmissionMode(draft.submissionStatus)
          setDraftClientSubmissionId(draft.clientSubmissionId)
        }
      } finally {
        if (!cancelled) setDraftReady(true)
      }
    }
    void hydrateOfflineState()
    return () => {
      cancelled = true
    }
  }, [refreshQueueState])

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
    const handleOnline = () => {
      setIsOnline(true)
      void syncQueuedSubmissions()
    }
    const handleOffline = () => setIsOnline(false)
    window.addEventListener('online', handleOnline)
    window.addEventListener('offline', handleOffline)
    return () => {
      window.removeEventListener('online', handleOnline)
      window.removeEventListener('offline', handleOffline)
    }
  }, [syncQueuedSubmissions])

  useEffect(() => {
    if (!draftReady) return
    const handle = window.setTimeout(() => {
      void (async () => {
        const draft: PersistedDraft = {
          id: 'active',
          clientSubmissionId: draftClientSubmissionId,
          step,
          updatedAt: new Date().toISOString(),
          locked: submissionMode === 'queued',
          submissionStatus: submissionMode,
          payload: {
            form: await buildPersistedForm(form),
            signatureDataUrl,
            beforePhoto: await toStoredMediaFile(beforePhotoFile, 'before-photo.jpg'),
            afterPhoto: await toStoredMediaFile(afterPhotoFile, 'after-photo.jpg'),
          },
        }
        await saveDraft(draft)
      })()
    }, 300)
    return () => window.clearTimeout(handle)
  }, [
    afterPhotoFile,
    beforePhotoFile,
    draftClientSubmissionId,
    draftReady,
    form,
    signatureDataUrl,
    step,
    submissionMode,
  ])

  useEffect(() => {
    if (!draftReady) return
    void syncQueuedSubmissions()
  }, [draftReady, syncQueuedSubmissions])

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

  const beforePhotoPreviewUrl = useFilePreviewDataUrl(beforePhotoFile)
  const afterPhotoPreviewUrl = useFilePreviewDataUrl(afterPhotoFile)

  const reset = useCallback(() => {
    setForm(createBlankForm())
    setSignatureDataUrl('')
    setBeforePhotoFile(null)
    setAfterPhotoFile(null)
    setStep(0)
    setSubmitted(false)
    setSubmissionMode('draft')
    setDraftClientSubmissionId(createClientSubmissionId())
    setSubmitting(false)
    void clearDraft()
  }, [])

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
  }, [step, submitted, submitting])

  const persistReferenceCache = useCallback(async (next: Partial<CachedReferenceData>) => {
    const record: CachedReferenceData = {
      id: 'main',
      fetchedAt: new Date().toISOString(),
      customerAreas,
      drivers,
      vehicles,
      categories,
      ...next,
    }
    setCacheFetchedAt(record.fetchedAt)
    setUsingCachedReferenceData(false)
    await writeCachedReferenceData(record)
  }, [categories, customerAreas, drivers, vehicles])

  const showUndoToast = useCallback((text: string, onAction: () => void) => {
    setToast({
      type: 'success',
      text,
      actionLabel: 'Undo',
      onAction,
      durationMs: 7000,
    })
  }, [])

  const addCustomer = useCallback(async (value: CustomerAdminForm) => {
    const cleaned = normalizeCustomerAdminForm(value)
    if (!isCustomerAdminFormValid(cleaned)) return
    if (customerAreas.some((item) => (
      normalize(getCustomerName(item)) === normalize(cleaned.Title)
      && normalize(getCustomerArea(item)) === normalize(cleaned.field_1)
      && normalize(getCustomerSubLocationRaw(item)) === normalize(cleaned.field_2)
    ))) {
      setToast({ type: 'error', text: `Customer "${cleaned.Title}" already exists for that location.` })
      return
    }
    try {
      const result = await Customer_area_data_clean_finalService.create(cleaned as Omit<Customer_area_data_clean_finalWrite, 'ID'>)
      if (!result.data?.ID) throw new Error('Could not add the customer.')
      const nextCustomers = [...customerAreas, result.data as CustomerAreaOption]
      setCustomerAreas(nextCustomers)
      await persistReferenceCache({ customerAreas: nextCustomers })
      setToast({ type: 'success', text: `Customer "${cleaned.Title}" added.` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not add the customer.' })
    }
  }, [customerAreas, persistReferenceCache])

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
      const nextDrivers = [...drivers, result.data as Drivers1Read]
      setDrivers(nextDrivers)
      await persistReferenceCache({ drivers: nextDrivers })
      setToast({ type: 'success', text: `Driver "${cleaned}" added.` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not add the driver.' })
    }
  }, [driverOptions, drivers, persistReferenceCache])

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
      const nextVehicles = [...vehicles, result.data as VehiclesRead]
      setVehicles(nextVehicles)
      await persistReferenceCache({ vehicles: nextVehicles })
      setToast({ type: 'success', text: `Vehicle "${cleaned}" added.` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not add the vehicle.' })
    }
  }, [persistReferenceCache, vehicles])

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
      const nextCategories = [...categories, result.data as Waste_CategoriesRead]
      setCategories(nextCategories)
      await persistReferenceCache({ categories: nextCategories })
      setToast({ type: 'success', text: `Waste category "${cleaned}" added.` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not add the waste category.' })
    }
  }, [categories, persistReferenceCache])

  const updateDriver = useCallback(async (id: number, name: string) => {
    const cleaned = name.trim()
    if (!cleaned) return
    if (drivers.some((driver) => driver.ID !== id && normalize(getDriverName(driver)) === normalize(cleaned))) {
      setToast({ type: 'error', text: `"${cleaned}" already exists in drivers.` })
      return
    }
    try {
      const result = await Drivers1Service.update(String(id), { Title: cleaned })
      const nextDrivers = drivers.map((driver) => (
        driver.ID === id ? (result.data ?? { ...driver, Title: cleaned }) as Drivers1Read : driver
      ))
      setDrivers(nextDrivers)
      await persistReferenceCache({ drivers: nextDrivers })
      setToast({ type: 'success', text: `Driver updated to "${cleaned}".` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not update the driver.' })
    }
  }, [drivers, persistReferenceCache])

  const updateVehicle = useCallback(async (id: number, vehicleNumber: string) => {
    const cleaned = vehicleNumber.trim()
    if (!cleaned) return
    if (vehicles.some((vehicle) => vehicle.ID !== id && normalize(vehicle.Title || '') === normalize(cleaned))) {
      setToast({ type: 'error', text: `"${cleaned}" already exists in vehicles.` })
      return
    }
    try {
      const result = await VehiclesService.update(String(id), { Title: cleaned })
      const nextVehicles = vehicles.map((vehicle) => (
        vehicle.ID === id ? (result.data ?? { ...vehicle, Title: cleaned }) as VehiclesRead : vehicle
      ))
      setVehicles(nextVehicles)
      await persistReferenceCache({ vehicles: nextVehicles })
      setToast({ type: 'success', text: `Vehicle updated to "${cleaned}".` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not update the vehicle.' })
    }
  }, [persistReferenceCache, vehicles])

  const updateWasteCategory = useCallback(async (id: number, category: string) => {
    const cleaned = category.trim()
    if (!cleaned) return
    if (categories.some((item) => item.ID !== id && normalize(item.Title || '') === normalize(cleaned))) {
      setToast({ type: 'error', text: `"${cleaned}" already exists in waste categories.` })
      return
    }
    try {
      const result = await Waste_CategoriesService.update(String(id), { Title: cleaned })
      const nextCategories = categories.map((item) => (
        item.ID === id ? (result.data ?? { ...item, Title: cleaned }) as Waste_CategoriesRead : item
      ))
      setCategories(nextCategories)
      await persistReferenceCache({ categories: nextCategories })
      setToast({ type: 'success', text: `Waste category updated to "${cleaned}".` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not update the waste category.' })
    }
  }, [categories, persistReferenceCache])

  const updateCustomer = useCallback(async (id: number, value: CustomerAdminForm) => {
    const cleaned = normalizeCustomerAdminForm(value)
    if (!isCustomerAdminFormValid(cleaned)) return
    if (customerAreas.some((item) => item.ID !== id && (
      normalize(getCustomerName(item)) === normalize(cleaned.Title)
      && normalize(getCustomerArea(item)) === normalize(cleaned.field_1)
      && normalize(getCustomerSubLocationRaw(item)) === normalize(cleaned.field_2)
    ))) {
      setToast({ type: 'error', text: `Customer "${cleaned.Title}" already exists for that location.` })
      return
    }
    try {
      const result = await Customer_area_data_clean_finalService.update(
        String(id),
        cleaned as Partial<Omit<Customer_area_data_clean_finalWrite, 'ID'>>,
      )
      const nextCustomers = customerAreas.map((item) => (
        item.ID === id ? (result.data ?? { ...item, ...cleaned }) as CustomerAreaOption : item
      ))
      setCustomerAreas(nextCustomers)
      await persistReferenceCache({ customerAreas: nextCustomers })
      setToast({ type: 'success', text: `Customer "${cleaned.Title}" updated.` })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not update the customer.' })
    }
  }, [customerAreas, persistReferenceCache])

  const deleteDriver = useCallback(async (id: number, name: string) => {
    const deletedDriver = drivers.find((driver) => driver.ID === id)
    try {
      await Drivers1Service.delete(String(id))
      const nextDrivers = drivers.filter((driver) => driver.ID !== id)
      setDrivers(nextDrivers)
      await persistReferenceCache({ drivers: nextDrivers })
      showUndoToast(`Driver "${name}" deleted.`, () => {
        if (!deletedDriver) return
        void (async () => {
          try {
            const restored = await Drivers1Service.create({ Title: getDriverName(deletedDriver) || name })
            const restoredDrivers = [...nextDrivers, restored.data as Drivers1Read]
            setDrivers(restoredDrivers)
            await persistReferenceCache({ drivers: restoredDrivers })
            setToast({ type: 'success', text: `Driver "${name}" restored.` })
          } catch (error) {
            setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not restore the driver.' })
          }
        })()
      })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not delete the driver.' })
    }
  }, [drivers, persistReferenceCache, showUndoToast])

  const deleteVehicle = useCallback(async (id: number, vehicleNumber: string) => {
    const deletedVehicle = vehicles.find((vehicle) => vehicle.ID === id)
    try {
      await VehiclesService.delete(String(id))
      const nextVehicles = vehicles.filter((vehicle) => vehicle.ID !== id)
      setVehicles(nextVehicles)
      await persistReferenceCache({ vehicles: nextVehicles })
      showUndoToast(`Vehicle "${vehicleNumber}" deleted.`, () => {
        if (!deletedVehicle) return
        void (async () => {
          try {
            const restored = await VehiclesService.create({ Title: deletedVehicle.Title?.trim() || vehicleNumber } as Omit<VehiclesWrite, 'ID'>)
            const restoredVehicles = [...nextVehicles, restored.data as VehiclesRead]
            setVehicles(restoredVehicles)
            await persistReferenceCache({ vehicles: restoredVehicles })
            setToast({ type: 'success', text: `Vehicle "${vehicleNumber}" restored.` })
          } catch (error) {
            setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not restore the vehicle.' })
          }
        })()
      })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not delete the vehicle.' })
    }
  }, [persistReferenceCache, showUndoToast, vehicles])

  const deleteWasteCategory = useCallback(async (id: number, category: string) => {
    const deletedCategory = categories.find((item) => item.ID === id)
    try {
      await Waste_CategoriesService.delete(String(id))
      const nextCategories = categories.filter((item) => item.ID !== id)
      setCategories(nextCategories)
      await persistReferenceCache({ categories: nextCategories })
      showUndoToast(`Waste category "${category}" deleted.`, () => {
        if (!deletedCategory) return
        void (async () => {
          try {
            const restored = await Waste_CategoriesService.create({ Title: deletedCategory.Title?.trim() || category } as Omit<Waste_CategoriesWrite, 'ID'>)
            const restoredCategories = [...nextCategories, restored.data as Waste_CategoriesRead]
            setCategories(restoredCategories)
            await persistReferenceCache({ categories: restoredCategories })
            setToast({ type: 'success', text: `Waste category "${category}" restored.` })
          } catch (error) {
            setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not restore the waste category.' })
          }
        })()
      })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not delete the waste category.' })
    }
  }, [categories, persistReferenceCache, showUndoToast])

  const deleteCustomer = useCallback(async (id: number, name: string) => {
    const deletedCustomer = customerAreas.find((item) => item.ID === id)
    try {
      await Customer_area_data_clean_finalService.delete(String(id))
      const nextCustomers = customerAreas.filter((item) => item.ID !== id)
      setCustomerAreas(nextCustomers)
      await persistReferenceCache({ customerAreas: nextCustomers })
      showUndoToast(`Customer "${name}" deleted.`, () => {
        if (!deletedCustomer) return
        void (async () => {
          try {
            const restored = await Customer_area_data_clean_finalService.create(
              normalizeCustomerAdminForm(getCustomerAdminFormFromRow(deletedCustomer)) as Omit<Customer_area_data_clean_finalWrite, 'ID'>,
            )
            const restoredCustomers = [...nextCustomers, restored.data as CustomerAreaOption]
            setCustomerAreas(restoredCustomers)
            await persistReferenceCache({ customerAreas: restoredCustomers })
            setToast({ type: 'success', text: `Customer "${name}" restored.` })
          } catch (error) {
            setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not restore the customer.' })
          }
        })()
      })
    } catch (error) {
      setToast({ type: 'error', text: error instanceof Error ? error.message : 'Could not delete the customer.' })
    }
  }, [customerAreas, persistReferenceCache, showUndoToast])

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

    try {
      const payload = {
        form: await buildPersistedForm(form),
        signatureDataUrl,
        beforePhoto: await toStoredMediaFile(beforePhotoFile, 'before-photo.jpg'),
        afterPhoto: await toStoredMediaFile(afterPhotoFile, 'after-photo.jpg'),
      }
      const queuedSubmission = await enqueueSubmission(draftClientSubmissionId, payload)
      await saveDraft({
        id: 'active',
        clientSubmissionId: draftClientSubmissionId,
        step,
        updatedAt: new Date().toISOString(),
        locked: true,
        submissionStatus: 'queued',
        payload,
      })
      setSubmissionMode('queued')
      setSubmitted(true)
      await refreshQueueState()

      if (window.navigator.onLine) {
        await syncQueuedSubmissions()
      } else {
        setToast({ type: 'success', text: `Order "${queuedSubmission.orderTitle}" saved offline and queued for sync.` })
      }
    } catch (e) {
      setToast({ type: 'error', text: e instanceof Error ? e.message : 'Unexpected error.' })
    } finally {
      setSubmitting(false)
    }
  }, [
    afterPhotoFile,
    beforePhotoFile,
    draftClientSubmissionId,
    form,
    loadState,
    refreshQueueState,
    signatureDataUrl,
    step,
    syncQueuedSubmissions,
  ])

  /* ── Step completion states ── */
  const stepDone = (idx: number) => {
    if (submitted) return true
    return idx < step
  }

  const activeReferenceLabel: Record<AdminTab, string> = {
    customers: 'Customers',
    drivers: 'Drivers',
    vehicles: 'Vehicles',
    waste: 'Waste Categories',
  }
  const breadcrumbCurrentLabel = view === 'form'
    ? 'New Service Order'
    : view === 'admin'
      ? `Admin / ${activeReferenceLabel[adminTab]}`
      : 'All Orders'
  const pageTitle = view === 'form'
    ? 'Create Service Order'
    : view === 'admin'
      ? activeReferenceLabel[adminTab]
      : 'Service Orders'
  const pageDescription = view === 'form'
    ? 'Complete the guided flow below to create a new service order from any device.'
    : view === 'admin'
      ? `Manage ${activeReferenceLabel[adminTab].toLowerCase()} used by the service order workflow.`
      : 'Review service order records and filter the history by customer, resource, or collection date.'

  const syncStatus = !isOnline
    ? 'offline'
    : queueSyncing
      ? 'syncing'
      : queueLastError
        ? 'error'
        : queuePendingCount > 0
          ? 'pending'
          : 'ready'
  const syncHeadline = syncStatus === 'offline'
    ? 'Saved locally while offline'
    : syncStatus === 'syncing'
      ? 'Sync in progress'
      : syncStatus === 'error'
        ? 'Sync needs attention'
        : syncStatus === 'pending'
          ? 'Queued submissions waiting to sync'
          : 'Everything is up to date'
  const syncBody = syncStatus === 'offline'
    ? 'New submissions stay on this device and will sync automatically after the connection returns.'
    : syncStatus === 'syncing'
      ? 'Queued service orders are being sent to Power Apps now.'
      : syncStatus === 'error'
        ? 'The last background sync failed. Review the error below and retry when you are ready.'
        : syncStatus === 'pending'
          ? `${queuePendingCount} queued submission${queuePendingCount === 1 ? '' : 's'} still need to sync.`
          : 'Live reference data is connected and there are no queued submissions.'
  const referenceDataSummary = usingCachedReferenceData
    ? `Using cached reference data${cacheFetchedAt ? ` from ${new Date(cacheFetchedAt).toLocaleString()}` : ''}.`
    : 'Reference data is live and ready.'
  const stepProgress = submitted ? 100 : Math.round(((step + 1) / STEPS.length) * 100)

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
        adminTab={adminTab}
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
              {breadcrumbCurrentLabel}
            </span>
          </div>
        </div>

        <div className="cora-page">
          {/* Header */}
          <div className="cora-page-header">
            <h1>{pageTitle}</h1>
            <p>{pageDescription}</p>
          </div>

          <div className={`cora-sync-banner is-${syncStatus}`} role="status" aria-live="polite">
            <div className="cora-sync-banner-main">
              <span className="cora-sync-banner-badge">{syncHeadline}</span>
              <strong>{syncBody}</strong>
              <p>{referenceDataSummary}</p>
              {queueLastError && <span className="cora-sync-banner-error">Last sync error: {queueLastError}</span>}
            </div>
            <div className="cora-sync-banner-actions">
              <span className="cora-sync-banner-stat">
                {queuePendingCount > 0
                  ? `${queuePendingCount} queued submission${queuePendingCount === 1 ? '' : 's'}`
                  : '0 queued submissions'}
              </span>
              <button
                className="cora-btn cora-btn-outline"
                type="button"
                onClick={() => void syncQueuedSubmissions()}
                disabled={!isOnline || queueSyncing || queuePendingCount === 0}
              >
                {queueSyncing ? 'Syncing…' : 'Retry Sync'}
              </button>
            </div>
          </div>

          {/* ════ LIST VIEW ════ */}
          {view === 'list' && isAdmin && adminUnlocked && <OrdersTable orders={orders} loadState={loadState} />}

          {view === 'admin' && isAdmin && adminUnlocked && (
            <AdminReferencePage
              activeTab={adminTab}
              onTabChange={setAdminTab}
              customerAreas={customerAreas}
              drivers={drivers}
              vehicles={vehicles}
              categories={categories}
              busy={loadState !== 'loaded'}
              currentUserEmail={currentUserEmail}
              onAddCustomer={addCustomer}
              onAddDriver={addDriver}
              onAddVehicle={addVehicle}
              onAddWasteCategory={addWasteCategory}
              onUpdateCustomer={updateCustomer}
              onUpdateDriver={updateDriver}
              onUpdateVehicle={updateVehicle}
              onUpdateWasteCategory={updateWasteCategory}
              onDeleteCustomer={deleteCustomer}
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

              <div className="cora-mobile-progress" aria-label={`Step ${submitted ? STEPS.length : step + 1} of ${STEPS.length}`}>
                <div className="cora-mobile-progress-top">
                  <strong>{submitted ? 'Completed' : STEPS[step].label}</strong>
                  <span>{submitted ? 'All steps done' : `Step ${step + 1} of ${STEPS.length}`}</span>
                </div>
                <div className="cora-mobile-progress-track" aria-hidden="true">
                  <span style={{ width: `${stepProgress}%` }} />
                </div>
                <p>{submitted ? 'Your latest order has been saved.' : STEPS[step].desc}</p>
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
                      <h2>{submissionMode === 'queued' ? 'Service Order Queued' : 'Service Order Created'}</h2>
                      <p>{submissionMode === 'queued'
                        ? (isOnline
                          ? 'Your order snapshot is saved locally and is syncing in the background.'
                          : 'Your order snapshot is saved locally and will sync automatically when internet access returns.')
                        : 'Your order has been submitted successfully and is now being processed.'}</p>
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
                        {I.x} Clear This Step
                      </button>
                    </div>
                    <div className="cora-card-foot-right">
                      {step > 0 && (
                        <button className="cora-btn cora-btn-outline" onClick={prev} disabled={submitting} id="btn-prev">
                          {I.chevL} Back to {STEPS[step - 1].label}
                        </button>
                      )}
                      {step < LAST_STEP ? (
                        <button
                          className="cora-btn cora-btn-teal"
                          onClick={next}
                          disabled={submitting}
                          id="btn-next"
                        >
                          Continue to {STEPS[step + 1].label} {I.chevR}
                        </button>
                      ) : (
                        <button
                          className="cora-btn cora-btn-primary"
                          onClick={submit}
                          disabled={submitting || loadState !== 'loaded' || !form.Title.trim() || !signatureDataUrl}
                          id="btn-submit"
                        >
                          {submitting
                            ? <><span className="cora-spinner" /> Saving order…</>
                            : <>{I.send} {isOnline ? 'Submit Order' : 'Save Offline'}</>}
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
    </div>
  )
}

export default App
