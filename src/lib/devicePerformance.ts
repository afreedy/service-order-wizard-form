export interface DevicePerformanceProfile {
  isAndroid: boolean
  isChrome: boolean
  isLowResourceMode: boolean
  preferReducedEffects: boolean
  maxImageSide: number
  maxOcrImageSide: number
  maxOcrVariants: number
}

interface DevicePerformanceInputs {
  navigator?: Navigator
  matchMedia?: (query: string) => Pick<MediaQueryList, 'matches'>
  viewport?: { width: number; height: number }
}

const readNavigatorNumber = (navigatorValue: Navigator | undefined, key: 'deviceMemory' | 'hardwareConcurrency') => {
  const value = navigatorValue ? (navigatorValue as Navigator & Partial<Record<typeof key, number>>)[key] : undefined
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined
}

export const getDevicePerformanceProfile = (inputs: DevicePerformanceInputs = {}): DevicePerformanceProfile => {
  const nav = inputs.navigator ?? globalThis.navigator
  const userAgent = nav?.userAgent ?? ''
  const isAndroid = /Android/i.test(userAgent)
  const isChrome = /Chrome|CriOS/i.test(userAgent) && !/Edg|OPR|SamsungBrowser/i.test(userAgent)
  const memory = readNavigatorNumber(nav, 'deviceMemory')
  const cores = readNavigatorNumber(nav, 'hardwareConcurrency')
  const viewport = inputs.viewport ?? {
    width: globalThis.innerWidth || 0,
    height: globalThis.innerHeight || 0,
  }
  const matchMedia = inputs.matchMedia ?? globalThis.matchMedia?.bind(globalThis)
  const prefersReducedMotion = Boolean(matchMedia?.('(prefers-reduced-motion: reduce)').matches)
  const tabletSized = Math.min(viewport.width, viewport.height) >= 600 && Math.max(viewport.width, viewport.height) <= 1400
  const lowMemory = typeof memory === 'number' && memory <= 3
  const lowCoreCount = typeof cores === 'number' && cores <= 4
  const olderAndroidChromeTablet = isAndroid && isChrome && tabletSized && (lowMemory || lowCoreCount)
  const isLowResourceMode = prefersReducedMotion || olderAndroidChromeTablet

  return {
    isAndroid,
    isChrome,
    isLowResourceMode,
    preferReducedEffects: isLowResourceMode,
    maxImageSide: isLowResourceMode ? 560 : 720,
    maxOcrImageSide: isLowResourceMode ? 1280 : 2400,
    maxOcrVariants: isLowResourceMode ? 4 : 8,
  }
}
