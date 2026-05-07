import { describe, expect, it, vi } from 'vitest'
import { getDevicePerformanceProfile } from './devicePerformance'

const makeNavigator = (overrides: Partial<Navigator> & Record<string, unknown>) => ({
  userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0',
  hardwareConcurrency: 8,
  ...overrides,
})

describe('devicePerformance', () => {
  it('marks older Android Chrome tablets as low resource', () => {
    const profile = getDevicePerformanceProfile({
      navigator: makeNavigator({
        userAgent: 'Mozilla/5.0 (Linux; Android 10; SAMSUNG SM-T510) AppleWebKit/537.36 Chrome/110.0 Mobile Safari/537.36',
        hardwareConcurrency: 4,
        deviceMemory: 2,
      }) as Navigator,
      matchMedia: vi.fn().mockReturnValue({ matches: false }),
      viewport: { width: 800, height: 1280 },
    })

    expect(profile.isLowResourceMode).toBe(true)
    expect(profile.maxImageSide).toBeLessThan(720)
    expect(profile.maxOcrVariants).toBeLessThan(8)
    expect(profile.preferReducedEffects).toBe(true)
  })

  it('keeps capable desktop Chrome on the standard profile', () => {
    const profile = getDevicePerformanceProfile({
      navigator: makeNavigator({
        hardwareConcurrency: 12,
        deviceMemory: 16,
      }) as Navigator,
      matchMedia: vi.fn().mockReturnValue({ matches: false }),
      viewport: { width: 1440, height: 900 },
    })

    expect(profile.isLowResourceMode).toBe(false)
    expect(profile.maxImageSide).toBe(720)
    expect(profile.maxOcrVariants).toBe(8)
    expect(profile.preferReducedEffects).toBe(false)
  })

  it('honors reduced motion as a reduced-effects signal', () => {
    const profile = getDevicePerformanceProfile({
      navigator: makeNavigator({}) as Navigator,
      matchMedia: vi.fn().mockReturnValue({ matches: true }),
      viewport: { width: 1440, height: 900 },
    })

    expect(profile.preferReducedEffects).toBe(true)
  })
})
