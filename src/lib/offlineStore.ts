import type { Customer_area_data_clean_finalRead } from '../generated/models/Customer_area_data_clean_finalModel'
import type { Drivers1Read } from '../generated/models/Drivers1Model'
import type { VehiclesRead } from '../generated/models/VehiclesModel'
import type { Waste_CategoriesRead } from '../generated/models/Waste_CategoriesModel'
import type { CropRect, ScaleOcrStatus } from '../weightOcr'

export interface WasteLine {
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

export interface Form {
  Title: string
  Customer: string
  CustomerName: string
  CustomerLocation: string
  CustomerTenant: string
  CustomerLevel4: string
  CustomerLevel5: string
  CustomerLevel6: string
  CustomerLevel7: string
  CustomerLevel8: string
  CustomerLevel9: string
  CustomerLevel10: string
  IsAdhocCustomer: boolean
  DriverName: string
  VehicleNumber: string
  DateOfCollection: string
  WasteItems: WasteLine[]
  Notes: string
}

export interface StoredMediaFile {
  blob: Blob
  name: string
  type: string
  lastModified: number
}

export interface PersistedWasteLine extends Omit<WasteLine, 'scalePhotoFile' | 'scalePhotoPreviewUrl' | 'scaleOcrCropPreviewUrl'> {
  scalePhoto?: StoredMediaFile | null
  scaleOcrCropPreviewDataUrl?: string
}

export interface PersistedForm extends Omit<Form, 'WasteItems'> {
  WasteItems: PersistedWasteLine[]
}

export interface SubmissionPayload {
  form: PersistedForm
  signatureDataUrl: string
  beforePhoto?: StoredMediaFile | null
  afterPhoto?: StoredMediaFile | null
}

export type QueuedSubmissionStatus = 'queued' | 'syncing' | 'failed' | 'synced'

export interface QueuedSubmissionRemoteState {
  serviceOrderIds?: number[]
  proofQueueId?: number
}

export interface QueuedSubmission {
  id: string
  clientSubmissionId: string
  orderTitle: string
  createdAt: string
  updatedAt: string
  status: QueuedSubmissionStatus
  attempts: number
  lastError?: string
  payload: SubmissionPayload
  remote?: QueuedSubmissionRemoteState
}

export interface PersistedDraft {
  id: 'active'
  clientSubmissionId: string
  step: number
  updatedAt: string
  locked: boolean
  submissionStatus: 'draft' | 'queued'
  payload: SubmissionPayload
}

export interface CachedReferenceData {
  id: 'main'
  fetchedAt: string
  customerAreas: Customer_area_data_clean_finalRead[]
  drivers: Drivers1Read[]
  vehicles: VehiclesRead[]
  categories: Waste_CategoriesRead[]
}

const DB_NAME = 'cora-offline-store'
const DB_VERSION = 1
const DRAFT_STORE = 'drafts'
const QUEUE_STORE = 'submissionQueue'
const CACHE_STORE = 'referenceCache'

let openPromise: Promise<IDBDatabase> | null = null

const openDatabase = () => {
  openPromise ??= new Promise<IDBDatabase>((resolve, reject) => {
    const request = window.indexedDB.open(DB_NAME, DB_VERSION)

    request.onupgradeneeded = () => {
      const db = request.result
      if (!db.objectStoreNames.contains(DRAFT_STORE)) {
        db.createObjectStore(DRAFT_STORE, { keyPath: 'id' })
      }
      if (!db.objectStoreNames.contains(QUEUE_STORE)) {
        const queueStore = db.createObjectStore(QUEUE_STORE, { keyPath: 'id' })
        queueStore.createIndex('status', 'status', { unique: false })
        queueStore.createIndex('createdAt', 'createdAt', { unique: false })
      }
      if (!db.objectStoreNames.contains(CACHE_STORE)) {
        db.createObjectStore(CACHE_STORE, { keyPath: 'id' })
      }
    }

    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error ?? new Error('Could not open offline storage.'))
  })

  return openPromise
}

const withStore = async <T>(
  storeName: string,
  mode: IDBTransactionMode,
  action: (store: IDBObjectStore) => Promise<T>,
) => {
  const db = await openDatabase()
  const transaction = db.transaction(storeName, mode)
  const store = transaction.objectStore(storeName)
  const result = await action(store)

  await new Promise<void>((resolve, reject) => {
    transaction.oncomplete = () => resolve()
    transaction.onabort = () => reject(transaction.error ?? new Error('IndexedDB transaction aborted.'))
    transaction.onerror = () => reject(transaction.error ?? new Error('IndexedDB transaction failed.'))
  })

  return result
}

const wrapRequest = <T>(request: IDBRequest<T>) =>
  new Promise<T>((resolve, reject) => {
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error ?? new Error('IndexedDB request failed.'))
  })

export const toStoredMediaFile = (file: Blob | File | null | undefined, fallbackName = 'image.jpg'): StoredMediaFile | null => {
  if (!file) return null
  const typedFile = file as File
  return {
    blob: file,
    name: typedFile.name || fallbackName,
    type: file.type || 'application/octet-stream',
    lastModified: typeof typedFile.lastModified === 'number' ? typedFile.lastModified : Date.now(),
  }
}

export const fromStoredMediaFile = (stored: StoredMediaFile | null | undefined) => {
  if (!stored) return null
  return new File([stored.blob], stored.name, {
    type: stored.type || stored.blob.type || 'application/octet-stream',
    lastModified: stored.lastModified || Date.now(),
  })
}

export const saveDraft = async (draft: PersistedDraft) => withStore(DRAFT_STORE, 'readwrite', async (store) => {
  await wrapRequest(store.put(draft))
})

export const loadDraft = async () => withStore(PERSISTED_STORE_NAMES.drafts, 'readonly', async (store) => {
  const result = await wrapRequest(store.get('active'))
  return (result as PersistedDraft | undefined) ?? null
})

export const clearDraft = async () => withStore(DRAFT_STORE, 'readwrite', async (store) => {
  await wrapRequest(store.delete('active'))
})

export const saveReferenceCache = async (cache: CachedReferenceData) => withStore(CACHE_STORE, 'readwrite', async (store) => {
  await wrapRequest(store.put(cache))
})

export const loadReferenceCache = async () => withStore(CACHE_STORE, 'readonly', async (store) => {
  const result = await wrapRequest(store.get('main'))
  return (result as CachedReferenceData | undefined) ?? null
})

export const putQueuedSubmission = async (item: QueuedSubmission) => withStore(QUEUE_STORE, 'readwrite', async (store) => {
  await wrapRequest(store.put(item))
})

export const getQueuedSubmission = async (id: string) => withStore(QUEUE_STORE, 'readonly', async (store) => {
  const result = await wrapRequest(store.get(id))
  return (result as QueuedSubmission | undefined) ?? null
})

export const listQueuedSubmissions = async () => withStore(QUEUE_STORE, 'readonly', async (store) => {
  const result = await wrapRequest(store.getAll())
  return ((result as QueuedSubmission[]) ?? []).sort((a, b) => (
    a.createdAt.localeCompare(b.createdAt) || a.id.localeCompare(b.id)
  ))
})

export const listPendingQueuedSubmissions = async () => {
  const items = await listQueuedSubmissions()
  return items.filter((item) => item.status === 'queued' || item.status === 'failed')
}

export const deleteQueuedSubmission = async (id: string) => withStore(QUEUE_STORE, 'readwrite', async (store) => {
  await wrapRequest(store.delete(id))
})

export const getQueueSummary = async () => {
  const items = await listQueuedSubmissions()
  const pendingCount = items.filter((item) => item.status === 'queued' || item.status === 'failed' || item.status === 'syncing').length
  const lastFailure = [...items]
    .filter((item) => item.status === 'failed' && item.lastError)
    .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt))[0]
  return {
    pendingCount,
    lastError: lastFailure?.lastError ?? '',
  }
}

const PERSISTED_STORE_NAMES = {
  drafts: DRAFT_STORE,
  queue: QUEUE_STORE,
  cache: CACHE_STORE,
} as const

