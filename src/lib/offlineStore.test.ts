import { beforeEach, describe, expect, it, vi } from 'vitest'
import { makeDraft, makeQueuedSubmission } from '../test/fixtures'
import {
  clearDraft,
  deleteQueuedSubmission,
  fromStoredMediaFile,
  getQueuedSubmission,
  getQueueSummary,
  listPendingQueuedSubmissions,
  listQueuedSubmissions,
  loadDraft,
  loadReferenceCache,
  OFFLINE_STORAGE_INTERRUPTED_MESSAGE,
  putQueuedSubmission,
  resetDatabaseConnection,
  saveDraft,
  saveReferenceCache,
  toStoredMediaFile,
} from './offlineStore'

describe('offlineStore', () => {
  beforeEach(async () => {
    resetDatabaseConnection()
    await clearDraft()
    await saveReferenceCache({
      id: 'main',
      fetchedAt: '',
      customerAreas: [],
      drivers: [],
      vehicles: [],
      categories: [],
    })
    const queued = await listQueuedSubmissions()
    await Promise.all(queued.map((item) => deleteQueuedSubmission(item.id)))
  })

  it('round-trips stored media files', async () => {
    const file = new File(['proof-image'], 'proof.png', {
      type: 'image/png',
      lastModified: 12345,
    })

    const stored = await toStoredMediaFile(file)
    const restored = await fromStoredMediaFile(stored)

    expect(stored).toMatchObject({
      name: 'proof.png',
      type: 'image/png',
      lastModified: 12345,
    })
    expect(stored?.dataUrl).toContain('data:image/png;base64,')
    expect(restored).not.toBeNull()
    await expect(restored?.text()).resolves.toBe('proof-image')
  })

  it('handles nullable and blob-backed stored media values', async () => {
    await expect(toStoredMediaFile(null)).resolves.toBeNull()
    await expect(fromStoredMediaFile(null)).resolves.toBeNull()

    const stored = {
      blob: new Blob(['blob-only'], { type: 'text/plain' }),
      name: '',
      type: '',
      lastModified: 0,
    }

    const restored = await fromStoredMediaFile(stored)

    expect(restored?.name).toBe('')
    expect(restored?.type).toBe('text/plain')
    expect(typeof restored?.lastModified).toBe('number')
    await expect(restored?.text()).resolves.toBe('blob-only')
  })

  it('rejects when FileReader does not return a string data url', async () => {
    const OriginalFileReader = globalThis.FileReader

    class BadFileReader {
      result: ArrayBuffer | null = new ArrayBuffer(8)
      error: DOMException | null = null
      onload: null | (() => void) = null
      onerror: null | (() => void) = null

      readAsDataURL() {
        this.onload?.()
      }
    }

    vi.stubGlobal('FileReader', BadFileReader)

    try {
      await expect(
        toStoredMediaFile(new File(['x'], 'bad.bin', { type: 'application/octet-stream' })),
      ).rejects.toThrow('Could not prepare media for offline storage.')
    } finally {
      vi.stubGlobal('FileReader', OriginalFileReader)
    }
  })

  it('saves and clears the active draft', async () => {
    const draft = makeDraft()

    await saveDraft(draft)
    await expect(loadDraft()).resolves.toEqual(draft)

    await clearDraft()
    await expect(loadDraft()).resolves.toBeNull()
  })

  it('stores cached reference data', async () => {
    const cache = {
      id: 'main' as const,
      fetchedAt: '2026-04-24T10:00:00.000Z',
      customerAreas: [{ ID: 1, Title: 'Acme', field_1: 'HQ', field_2: 'Tower A' }],
      drivers: [{ ID: 2, Title: 'Jordan' }],
      vehicles: [{ ID: 3, Title: 'Truck 12', VehicleNumber: 'TRK-012' }],
      categories: [{ ID: 4, Title: 'General Waste' }],
    }

    await saveReferenceCache(cache)

    await expect(loadReferenceCache()).resolves.toEqual(cache)
  })

  it('lists pending queue items and summarizes failures', async () => {
    await putQueuedSubmission(makeQueuedSubmission({ id: 'queued-1', clientSubmissionId: 'queued-1', status: 'queued' }))
    await putQueuedSubmission(makeQueuedSubmission({
      id: 'syncing-1',
      clientSubmissionId: 'syncing-1',
      status: 'syncing',
      updatedAt: '2026-04-24T10:01:00.000Z',
    }))
    await putQueuedSubmission(makeQueuedSubmission({
      id: 'failed-1',
      clientSubmissionId: 'failed-1',
      status: 'failed',
      updatedAt: '2026-04-24T10:02:00.000Z',
      lastError: 'Network timeout',
    }))
    await putQueuedSubmission(makeQueuedSubmission({
      id: 'synced-1',
      clientSubmissionId: 'synced-1',
      status: 'synced',
    }))
    await putQueuedSubmission(makeQueuedSubmission({
      id: 'failed-older',
      clientSubmissionId: 'failed-older',
      status: 'failed',
      updatedAt: '2026-04-24T09:59:00.000Z',
      lastError: 'Older failure',
    }))

    await expect(listPendingQueuedSubmissions()).resolves.toEqual(expect.arrayContaining([
      expect.objectContaining({ id: 'queued-1', status: 'queued' }),
      expect.objectContaining({ id: 'syncing-1', status: 'syncing' }),
      expect.objectContaining({ id: 'failed-1', status: 'failed' }),
    ]))
    await expect(getQueueSummary()).resolves.toEqual({
      pendingCount: 4,
      lastError: 'Network timeout',
    })
  })

  it('retrieves and deletes specific queued submissions', async () => {
    const queued = makeQueuedSubmission({ id: 'lookup-1', clientSubmissionId: 'lookup-1' })

    await putQueuedSubmission(queued)
    await expect(getQueuedSubmission('lookup-1')).resolves.toEqual(queued)

    await deleteQueuedSubmission('lookup-1')
    await expect(getQueuedSubmission('lookup-1')).resolves.toBeNull()
  })

  it('recovers when a stale database transaction fails once', async () => {
    await putQueuedSubmission(makeQueuedSubmission({ id: 'before-stale', clientSubmissionId: 'before-stale' }))
    resetDatabaseConnection()
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open('cora-offline-store', 1)
      request.onsuccess = () => resolve(request.result)
      request.onerror = () => reject(request.error)
    })
    const dbPrototype = Object.getPrototypeOf(db) as IDBDatabase
    const originalTransaction = dbPrototype.transaction
    const transactionSpy = vi.spyOn(dbPrototype, 'transaction')
    db.close()
    let calls = 0
    transactionSpy.mockImplementation(function transaction(this: IDBDatabase, ...args: Parameters<IDBDatabase['transaction']>) {
      calls += 1
      if (calls === 1) {
        throw new DOMException('The database connection is closing', 'InvalidStateError')
      }
      return originalTransaction.apply(this, args)
    })

    try {
      await putQueuedSubmission(makeQueuedSubmission({ id: 'after-stale', clientSubmissionId: 'after-stale' }))
      const queued = await listQueuedSubmissions()

      expect(calls).toBeGreaterThanOrEqual(2)
      expect(queued.filter((item) => item.id === 'after-stale')).toHaveLength(1)
    } finally {
      transactionSpy.mockRestore()
      resetDatabaseConnection()
    }
  })

  it('shows a friendly error when a stale database transaction retry also fails', async () => {
    resetDatabaseConnection()
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open('cora-offline-store', 1)
      request.onsuccess = () => resolve(request.result)
      request.onerror = () => reject(request.error)
    })
    const transactionSpy = vi.spyOn(Object.getPrototypeOf(db), 'transaction')
    db.close()
    transactionSpy.mockImplementation(() => {
      throw new DOMException('The database connection is closed', 'InvalidStateError')
    })

    try {
      await expect(
        putQueuedSubmission(makeQueuedSubmission({ id: 'retry-fails', clientSubmissionId: 'retry-fails' })),
      ).rejects.toThrow(OFFLINE_STORAGE_INTERRUPTED_MESSAGE)
    } finally {
      transactionSpy.mockRestore()
      resetDatabaseConnection()
    }
  })

  it('returns empty state for an empty cache and queue', async () => {
    await expect(loadReferenceCache()).resolves.toEqual({
      id: 'main',
      fetchedAt: '',
      customerAreas: [],
      drivers: [],
      vehicles: [],
      categories: [],
    })
    await expect(getQueueSummary()).resolves.toEqual({
      pendingCount: 0,
      lastError: '',
    })
  })
})
