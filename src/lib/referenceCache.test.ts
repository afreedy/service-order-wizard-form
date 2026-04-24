import { describe, expect, it, vi } from 'vitest'

const loadReferenceCache = vi.fn()
const saveReferenceCache = vi.fn()

vi.mock('./offlineStore', () => ({
  loadReferenceCache,
  saveReferenceCache,
}))

describe('referenceCache', () => {
  it('reads cached reference data through offline storage', async () => {
    const cached = { id: 'main', fetchedAt: '2026-04-24T10:00:00.000Z' }
    loadReferenceCache.mockResolvedValue(cached)

    const { readCachedReferenceData } = await import('./referenceCache')

    await expect(readCachedReferenceData()).resolves.toBe(cached)
  })

  it('writes cached reference data through offline storage', async () => {
    const cached = { id: 'main', fetchedAt: '2026-04-24T10:00:00.000Z' }
    saveReferenceCache.mockResolvedValue(undefined)

    const { writeCachedReferenceData } = await import('./referenceCache')

    await writeCachedReferenceData(cached as never)

    expect(saveReferenceCache).toHaveBeenCalledWith(cached)
  })
})
