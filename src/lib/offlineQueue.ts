import {
  getQueueSummary,
  listPendingQueuedSubmissions,
  putQueuedSubmission,
  type QueuedSubmission,
  type SubmissionPayload,
} from './offlineStore'
import { processQueuedSubmission } from './submitProcessor'

interface ProcessQueueCallbacks {
  onSubmissionSynced?: (item: QueuedSubmission) => void
  onSubmissionFailed?: (item: QueuedSubmission, error: string) => void
  onQueueUpdated?: () => void
}

let queueRunPromise: Promise<void> | null = null

const nowIso = () => new Date().toISOString()

export const createClientSubmissionId = () =>
  `submission-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`

export const enqueueSubmission = async (
  clientSubmissionId: string,
  payload: SubmissionPayload,
) => {
  const timestamp = nowIso()
  const queuedSubmission: QueuedSubmission = {
    id: clientSubmissionId,
    clientSubmissionId,
    orderTitle: payload.form.Title,
    createdAt: timestamp,
    updatedAt: timestamp,
    status: 'queued',
    attempts: 0,
    payload,
  }
  await putQueuedSubmission(queuedSubmission)
  return queuedSubmission
}

export const processPendingQueue = async (callbacks: ProcessQueueCallbacks = {}) => {
  if (queueRunPromise) {
    await queueRunPromise
    return
  }

  queueRunPromise = (async () => {
    const pendingItems = await listPendingQueuedSubmissions()
    for (const item of pendingItems) {
      const syncingItem: QueuedSubmission = {
        ...item,
        status: 'syncing',
        updatedAt: nowIso(),
        lastError: undefined,
      }
      await putQueuedSubmission(syncingItem)
      callbacks.onQueueUpdated?.()

      try {
        const result = await processQueuedSubmission(syncingItem)
        const syncedItem: QueuedSubmission = {
          ...syncingItem,
          status: 'synced',
          updatedAt: nowIso(),
          attempts: syncingItem.attempts + 1,
          lastError: undefined,
          remote: {
            serviceOrderIds: result.createdOrders.map((order) => Number(order.ID)).filter(Number.isFinite),
            proofQueueId: result.proofQueueItem?.ID,
          },
        }
        await putQueuedSubmission(syncedItem)
        callbacks.onSubmissionSynced?.(syncedItem)
      } catch (error) {
        const failedItem: QueuedSubmission = {
          ...syncingItem,
          status: 'failed',
          updatedAt: nowIso(),
          attempts: syncingItem.attempts + 1,
          lastError: error instanceof Error ? error.message : 'Could not sync the queued submission.',
        }
        await putQueuedSubmission(failedItem)
        callbacks.onSubmissionFailed?.(failedItem, failedItem.lastError || 'Could not sync the queued submission.')
      }

      callbacks.onQueueUpdated?.()
    }
  })()

  try {
    await queueRunPromise
  } finally {
    queueRunPromise = null
  }
}

export const readQueueSummary = () => getQueueSummary()

