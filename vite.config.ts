import { defineConfig } from 'vitest/config'
import react from '@vitejs/plugin-react'
import { powerApps } from '@microsoft/power-apps-vite/plugin'

export default defineConfig({
  base: './',
  plugins: [react(), powerApps()],
  test: {
    environment: 'jsdom',
    setupFiles: './src/test/setup.ts',
    css: true,
    coverage: {
      provider: 'v8',
      reporter: ['text', 'html'],
      include: ['src/lib/**/*.ts'],
      exclude: ['src/generated/**', 'src/main.tsx'],
      thresholds: {
        branches: 55,
        functions: 75,
        lines: 85,
        statements: 80,
      },
    },
  },
})
