import type { CachedReferenceData } from './offlineStore'
import { loadReferenceCache, saveReferenceCache } from './offlineStore'

export const readCachedReferenceData = () => loadReferenceCache()

export const writeCachedReferenceData = (cache: CachedReferenceData) => saveReferenceCache(cache)

