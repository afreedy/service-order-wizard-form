import { beforeEach, describe, expect, it, vi } from 'vitest'
import { makeQueuedSubmission } from '../test/fixtures'

const getQueueSummary = vi.fn()
const listPendingQueuedSubmissions = vi.fn()
const putQueuedSubmission = vi.fn()
const processQueuedSubmission = vi.fn()

vi.mock('./offlineStore', () => ({
  getQueueSummary,
  listPendingQueuedSubmissions,
  putQueuedSubmission,
}))

vi.mock('./submitProcessor', () => ({
  processQueuedSubmission,
}))

describe('offlineQueue', () => {
  beforeEach(() => {
    vi.resetModules()
    vi.clearAllMocks()
  })

  it('enqueues submissions with queue metadata', async () => {
    const { enqueueSubmission } = await import('./offlineQueue')
    const payload = makeQueuedSubmission().payload

    const queued = await enqueueSubmission('submission-123', payload)

    expect(putQueuedSubmission).toHaveBeenCalledWith(expect.objectContaining({
      id: 'submission-123',
      clientSubmissionId: 'submission-123',
      orderTitle: payload.form.Title,
      status: 'queued',
      attempts: 0,
      payload,
    }))
    expect(queued).toEqual(expect.objectContaining({
      id: 'submission-123',
      clientSubmissionId: 'submission-123',
      orderTitle: payload.form.Title,
      status: 'queued',
    }))
  })

  it('creates stable-looking client submission ids', async () => {
    const { createClientSubmissionId } = await import('./offlineQueue')

    expect(createClientSubmissionId()).toMatch(/^submission-\d+-[a-z0-9]{8}$/)
  })

  it('marks items as synced and emits callbacks after a successful queue run', async () => {
    const item = makeQueuedSubmission()
    listPendingQueuedSubmissions.mockResolvedValue([item])
    processQueuedSubmission.mockResolvedValue({
      createdOrders: [{ ID: 11 }, { ID: 12 }],
      proofQueueItem: { ID: 91 },
    })

    const onSubmissionSynced = vi.fn()
    const onQueueUpdated = vi.fn()
    const { processPendingQueue } = await import('./offlineQueue')

    await processPendingQueue({ onSubmissionSynced, onQueueUpdated })

    expect(putQueuedSubmission).toHaveBeenNthCalledWith(1, expect.objectContaining({
      id: item.id,
      status: 'syncing',
      attempts: 0,
      lastError: undefined,
    }))
    expect(putQueuedSubmission).toHaveBeenNthCalledWith(2, expect.objectContaining({
      id: item.id,
      status: 'synced',
      attempts: 1,
      remote: {
        serviceOrderIds: [11, 12],
        proofQueueId: 91,
      },
    }))
    expect(onSubmissionSynced).toHaveBeenCalledWith(expect.objectContaining({
      id: item.id,
      status: 'synced',
    }))
    expect(onQueueUpdated).toHaveBeenCalledTimes(2)
  })

  it('marks items as failed when submit processing throws', async () => {
    const item = makeQueuedSubmission()
    listPendingQueuedSubmissions.mockResolvedValue([item])
    processQueuedSubmission.mockRejectedValue(new Error('Power Apps unavailable'))

    const onSubmissionFailed = vi.fn()
    const { processPendingQueue } = await import('./offlineQueue')

    await processPendingQueue({ onSubmissionFailed })

    expect(putQueuedSubmission).toHaveBeenNthCalledWith(2, expect.objectContaining({
      id: item.id,
      status: 'failed',
      attempts: 1,
      lastError: 'Power Apps unavailable',
    }))
    expect(onSubmissionFailed).toHaveBeenCalledWith(
      expect.objectContaining({ id: item.id, status: 'failed' }),
      'Power Apps unavailable',
    )
  })

  it('serializes concurrent queue runs', async () => {
    const item = makeQueuedSubmission()
    listPendingQueuedSubmissions.mockResolvedValue([item])
    processQueuedSubmission.mockResolvedValue({ createdOrders: [{ ID: 44 }] })

    const { processPendingQueue } = await import('./offlineQueue')

    await Promise.all([processPendingQueue(), processPendingQueue()])

    expect(listPendingQueuedSubmissions).toHaveBeenCalledTimes(1)
    expect(processQueuedSubmission).toHaveBeenCalledTimes(1)
  })

  it('reads the queue summary from offline storage', async () => {
    getQueueSummary.mockResolvedValue({ pendingCount: 2, lastError: 'Boom' })

    const { readQueueSummary } = await import('./offlineQueue')

    await expect(readQueueSummary()).resolves.toEqual({ pendingCount: 2, lastError: 'Boom' })
  })

  it('handles non-Error failures with the default queue message', async () => {
    const item = makeQueuedSubmission()
    listPendingQueuedSubmissions.mockResolvedValue([item])
    processQueuedSubmission.mockRejectedValue('boom')

    const onSubmissionFailed = vi.fn()
    const { processPendingQueue } = await import('./offlineQueue')

    await processPendingQueue({ onSubmissionFailed })

    expect(putQueuedSubmission).toHaveBeenLastCalledWith(expect.objectContaining({
      status: 'failed',
      lastError: 'Could not sync the queued submission.',
    }))
    expect(onSubmissionFailed).toHaveBeenCalledWith(
      expect.objectContaining({ status: 'failed' }),
      'Could not sync the queued submission.',
    )
  })

  it('falls back to the default callback message when an Error has an empty message', async () => {
    const item = makeQueuedSubmission()
    listPendingQueuedSubmissions.mockResolvedValue([item])
    processQueuedSubmission.mockRejectedValue(new Error(''))

    const onSubmissionFailed = vi.fn()
    const { processPendingQueue } = await import('./offlineQueue')

    await processPendingQueue({ onSubmissionFailed })

    expect(onSubmissionFailed).toHaveBeenCalledWith(
      expect.objectContaining({ status: 'failed', lastError: '' }),
      'Could not sync the queued submission.',
    )
  })
})
