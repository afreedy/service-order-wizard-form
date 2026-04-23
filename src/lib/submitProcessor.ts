import type { ServiceOrderProofQueueRead, ServiceOrderProofQueueWrite } from '../generated/models/ServiceOrderProofQueueModel'
import type { ServiceOrdersRead, ServiceOrdersWrite } from '../generated/models/ServiceOrdersModel'
import type { ServiceOrderWasteItemsRead, ServiceOrderWasteItemsWrite } from '../generated/models/ServiceOrderWasteItemsModel'
import { ServiceOrderProofQueueService } from '../generated/services/ServiceOrderProofQueueService'
import { ServiceOrdersService } from '../generated/services/ServiceOrdersService'
import { ServiceOrderWasteItemsService } from '../generated/services/ServiceOrderWasteItemsService'
import type { PersistedWasteLine, QueuedSubmission, StoredMediaFile, SubmissionPayload } from './offlineStore'

interface ProofUploadResponse {
  queueItem?: ServiceOrderProofQueueRead
  queued: boolean
  signatureUrl?: string
  beforePhotoUrl?: string
  afterPhotoUrl?: string
}

export interface ProcessQueuedSubmissionResult {
  createdOrders: ServiceOrdersRead[]
  proofQueueItem?: ServiceOrderProofQueueRead
}

const delay = (ms: number) => new Promise((resolve) => window.setTimeout(resolve, ms))

const escapeODataString = (value: string) => value.replace(/'/g, "''")

const blobToBase64 = async (file: StoredMediaFile | null | undefined) => {
  if (!file) return undefined
  if (file.dataUrl) return file.dataUrl.split(',')[1] ?? ''
  if (!file.blob) return undefined
  const buffer = await file.blob.arrayBuffer()
  let binary = ''
  const bytes = new Uint8Array(buffer)
  const chunkSize = 0x8000
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize)
    binary += String.fromCharCode(...chunk)
  }
  return window.btoa(binary)
}

const getPhotoFileName = (orderTitle: string, label: 'before' | 'after', file: StoredMediaFile) =>
  `${orderTitle}-${label}-${file.name.replace(/[^a-z0-9]/gi, '-').slice(0, 24)}.jpg`

const findServiceOrdersByTitle = async (title: string) => {
  const filter = `Title eq '${escapeODataString(title)}'`
  const result = await ServiceOrdersService.getAll({ filter, top: 200 })
  return ((result.data ?? []) as ServiceOrdersRead[])
    .filter((order) => order.Title === title)
    .sort((a, b) => Number(a.ID ?? 0) - Number(b.ID ?? 0))
}

const findWasteItemsByTitle = async (title: string) => {
  const filter = `Title eq '${escapeODataString(title)}'`
  const result = await ServiceOrderWasteItemsService.getAll({ filter, top: 200 })
  return ((result.data ?? []) as ServiceOrderWasteItemsRead[])
    .filter((item) => item.Title === title)
    .sort((a, b) => Number(a.ID ?? 0) - Number(b.ID ?? 0))
}

const findProofQueueByOrderTitle = async (title: string) => {
  const filter = `OrderTitle eq '${escapeODataString(title)}'`
  const result = await ServiceOrderProofQueueService.getAll({ filter, top: 20 })
  return ((result.data ?? []) as ServiceOrderProofQueueRead[]).find((item) => item.OrderTitle === title)
}

const lineKey = (line: { WasteCategory?: string; Tonnage?: number | string | null }) =>
  `${String(line.WasteCategory ?? '').trim()}::${String(line.Tonnage ?? '').trim()}`

const countByKey = <T extends { WasteCategory?: string; Tonnage?: number | string | null }>(items: T[]) => {
  const counts = new Map<string, number>()
  for (const item of items) {
    const key = lineKey(item)
    counts.set(key, (counts.get(key) ?? 0) + 1)
  }
  return counts
}

const getCompleteWasteLines = (lines: PersistedWasteLine[]) => lines.filter((line) => {
  const tonnage = Number(line.Tonnage)
  return Boolean(line.WasteCategory.trim()) && line.Tonnage.trim() !== '' && Number.isFinite(tonnage) && tonnage >= 0
})

const getBaseOrderPayload = (payload: SubmissionPayload): Partial<Omit<ServiceOrdersWrite, 'ID'>> => ({
  Title: payload.form.Title || undefined,
  Customer: payload.form.Customer || undefined,
  CustomerName: payload.form.CustomerName || undefined,
  CustomerLocation: payload.form.CustomerLocation || undefined,
  CustomerTenant: payload.form.CustomerTenant || undefined,
  CustomerLevel4: payload.form.CustomerLevel4 || undefined,
  CustomerLevel5: payload.form.CustomerLevel5 || undefined,
  CustomerLevel6: payload.form.CustomerLevel6 || undefined,
  CustomerLevel7: payload.form.CustomerLevel7 || undefined,
  CustomerLevel8: payload.form.CustomerLevel8 || undefined,
  CustomerLevel9: payload.form.CustomerLevel9 || undefined,
  CustomerLevel10: payload.form.CustomerLevel10 || undefined,
  IsAdhocCustomer: payload.form.IsAdhocCustomer,
  DriverName: payload.form.DriverName || undefined,
  VehicleNumber: payload.form.VehicleNumber || undefined,
  DateOfCollection: payload.form.DateOfCollection || undefined,
  Notes: payload.form.Notes || undefined,
})

const rollbackNewRecords = async (
  createdProofQueue: ServiceOrderProofQueueRead | null,
  createdWasteItems: ServiceOrderWasteItemsRead[],
  createdOrders: ServiceOrdersRead[],
) => {
  const failures: string[] = []
  const tryDelete = async (label: string, action: () => Promise<void>) => {
    try {
      await action()
    } catch (error) {
      failures.push(`${label}: ${error instanceof Error ? error.message : 'cleanup failed'}`)
    }
  }

  if (createdProofQueue?.ID) {
    await tryDelete(`proof queue ${createdProofQueue.ID}`, () => ServiceOrderProofQueueService.delete(String(createdProofQueue.ID)))
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

const uploadServiceOrderProof = async (
  payload: SubmissionPayload,
  serviceOrderId: number,
): Promise<ProofUploadResponse> => {
  const existing = await findProofQueueByOrderTitle(payload.form.Title)
  if (existing?.ID) {
    return {
      queueItem: existing,
      queued: true,
      signatureUrl: existing.SignatureUrl || undefined,
      beforePhotoUrl: existing.BeforePhotoUrl || undefined,
      afterPhotoUrl: existing.AfterPhotoUrl || undefined,
    }
  }

  const queueItem: Partial<Omit<ServiceOrderProofQueueWrite, 'ID'>> = {
    Title: `${payload.form.Title}-proof`,
    ServiceOrderId: String(serviceOrderId),
    OrderTitle: payload.form.Title,
    SignatureFileName: `${payload.form.Title}-signature.png`,
    SignatureBase64: payload.signatureDataUrl.split(',')[1] ?? '',
    Processed: false,
  }

  if (payload.beforePhoto) {
    queueItem.BeforePhotoFileName = getPhotoFileName(payload.form.Title, 'before', payload.beforePhoto)
    queueItem.BeforePhotoBase64 = await blobToBase64(payload.beforePhoto)
  }

  if (payload.afterPhoto) {
    queueItem.AfterPhotoFileName = getPhotoFileName(payload.form.Title, 'after', payload.afterPhoto)
    queueItem.AfterPhotoBase64 = await blobToBase64(payload.afterPhoto)
  }

  const result = await ServiceOrderProofQueueService.create(queueItem as Omit<ServiceOrderProofQueueWrite, 'ID'>)
  if (result.data?.ID) {
    return { queueItem: result.data, queued: true }
  }

  for (let attempt = 0; attempt < 5; attempt += 1) {
    const confirmed = await findProofQueueByOrderTitle(payload.form.Title)
    if (confirmed?.ID) return { queueItem: confirmed, queued: true }
    await delay(400)
  }

  throw new Error('Could not confirm the proof media queue item.')
}

export const processQueuedSubmission = async (
  queuedSubmission: QueuedSubmission,
): Promise<ProcessQueuedSubmissionResult> => {
  const payload = queuedSubmission.payload
  const createdOrdersThisAttempt: ServiceOrdersRead[] = []
  const createdWasteItemsThisAttempt: ServiceOrderWasteItemsRead[] = []
  let createdProofQueueThisAttempt: ServiceOrderProofQueueRead | null = null

  try {
    const wasteLines = getCompleteWasteLines(payload.form.WasteItems)
    const desiredOrderRows: Array<{ WasteCategory: string; Tonnage: number | undefined }> = wasteLines.length > 0
      ? wasteLines.map((line) => ({ WasteCategory: line.WasteCategory, Tonnage: Number(line.Tonnage) }))
      : [{ WasteCategory: '', Tonnage: undefined }]
    const basePayload = getBaseOrderPayload(payload)

    let existingOrders = await findServiceOrdersByTitle(payload.form.Title)
    const existingOrderCounts = countByKey(existingOrders.map((order) => ({
      WasteCategory: order.WasteCategory ?? '',
      Tonnage: order.Tonnage ?? '',
    })))
    const desiredOrderCounts = countByKey(desiredOrderRows)

    for (const desiredRow of desiredOrderRows) {
      const key = lineKey(desiredRow)
      const desiredCount = desiredOrderCounts.get(key) ?? 0
      const existingCount = existingOrderCounts.get(key) ?? 0
      if (existingCount >= desiredCount) continue
      const createPayload: Partial<Omit<ServiceOrdersWrite, 'ID'>> = {
        ...basePayload,
        WasteCategory: desiredRow.WasteCategory || undefined,
        Tonnage: typeof desiredRow.Tonnage === 'number' && Number.isFinite(desiredRow.Tonnage)
          ? desiredRow.Tonnage
          : undefined,
      }
      const result = await ServiceOrdersService.create(createPayload as Omit<ServiceOrdersWrite, 'ID'>)
      if (result.data?.ID) {
        createdOrdersThisAttempt.push(result.data)
        existingOrderCounts.set(key, (existingOrderCounts.get(key) ?? 0) + 1)
      }
    }

    existingOrders = await findServiceOrdersByTitle(payload.form.Title)
    if (existingOrders.length < desiredOrderRows.length) {
      throw new Error(`Could not confirm the created service order rows for "${payload.form.Title}".`)
    }

    const proofOwnerOrder = existingOrders[0]
    if (!proofOwnerOrder?.ID) throw new Error('Could not create the service order.')

    if (wasteLines.length > 0) {
      let existingWasteItems = await findWasteItemsByTitle(payload.form.Title)
      const existingWasteCounts = countByKey(existingWasteItems.map((item) => ({
        WasteCategory: item.WasteCategory ?? '',
        Tonnage: item.Tonnage ?? '',
      })))
      const desiredWasteCounts = countByKey(wasteLines)

      for (const line of wasteLines) {
        const key = lineKey(line)
        const desiredCount = desiredWasteCounts.get(key) ?? 0
        const existingCount = existingWasteCounts.get(key) ?? 0
        if (existingCount >= desiredCount) continue
        const tonnage = Number(line.Tonnage)
        const wasteItem: Partial<Omit<ServiceOrderWasteItemsWrite, 'ID'>> = {
          Title: payload.form.Title,
          WasteCategory: line.WasteCategory,
          Tonnage: Number.isFinite(tonnage) ? tonnage : undefined,
        }
        const wasteResult = await ServiceOrderWasteItemsService.create(wasteItem as Omit<ServiceOrderWasteItemsWrite, 'ID'>)
        if (wasteResult.data?.ID) {
          createdWasteItemsThisAttempt.push(wasteResult.data)
          existingWasteCounts.set(key, (existingWasteCounts.get(key) ?? 0) + 1)
        }
      }

      existingWasteItems = await findWasteItemsByTitle(payload.form.Title)
      if (existingWasteItems.length < wasteLines.length) {
        throw new Error(`Could not confirm the waste line rows for "${payload.form.Title}".`)
      }
    }

    const proofUpload = await uploadServiceOrderProof(payload, Number(proofOwnerOrder.ID))
    createdProofQueueThisAttempt = proofUpload.queueItem ?? null

    const proofUrlFields: Partial<Omit<ServiceOrdersWrite, 'ID'>> = {
      SignatureUrl: proofUpload.signatureUrl || undefined,
      BeforePhotoUrl: proofUpload.beforePhotoUrl || undefined,
      AfterPhotoUrl: proofUpload.afterPhotoUrl || undefined,
    }
    const hasProofUrls = Object.values(proofUrlFields).some(Boolean)
    let savedOrders = [...existingOrders]
    if (hasProofUrls) {
      savedOrders = []
      for (const order of existingOrders) {
        if (!order.ID) continue
        const updated = await ServiceOrdersService.update(String(order.ID), proofUrlFields)
        savedOrders.push((updated.data ?? { ...order, ...proofUrlFields }) as ServiceOrdersRead)
      }
    }

    return {
      createdOrders: savedOrders,
      proofQueueItem: proofUpload.queueItem,
    }
  } catch (error) {
    const rollbackFailures = await rollbackNewRecords(
      createdProofQueueThisAttempt,
      createdWasteItemsThisAttempt,
      createdOrdersThisAttempt,
    )
    const rollbackText = rollbackFailures.length > 0
      ? ` Rollback needs manual cleanup. ${rollbackFailures.join(' ')}`
      : ''
    throw new Error(`${error instanceof Error ? error.message : 'Unexpected error.'}${rollbackText}`)
  }
}
