import { render, screen, waitFor } from '@testing-library/react'
import { beforeEach, describe, expect, it, vi } from 'vitest'

const getContext = vi.fn()
const driversGetAll = vi.fn()
const customersGetAll = vi.fn()
const vehiclesGetAll = vi.fn()
const wasteGetAll = vi.fn()
const ordersGetAll = vi.fn()
const clearDraft = vi.fn()
const fromStoredMediaFile = vi.fn()
const getQueueSummary = vi.fn()
const loadDraft = vi.fn()
const saveDraft = vi.fn()
const toStoredMediaFile = vi.fn()
const createClientSubmissionId = vi.fn(() => 'submission-app-test')
const enqueueSubmission = vi.fn()
const processPendingQueue = vi.fn()
const readCachedReferenceData = vi.fn()
const writeCachedReferenceData = vi.fn()

vi.mock('@microsoft/power-apps/app', () => ({
  getContext,
}))

vi.mock('./generated/services/Drivers1Service', () => ({
  Drivers1Service: { getAll: driversGetAll },
}))

vi.mock('./generated/services/Customer_area_data_clean_finalService', () => ({
  Customer_area_data_clean_finalService: { getAll: customersGetAll },
}))

vi.mock('./generated/services/VehiclesService', () => ({
  VehiclesService: { getAll: vehiclesGetAll },
}))

vi.mock('./generated/services/Waste_CategoriesService', () => ({
  Waste_CategoriesService: { getAll: wasteGetAll },
}))

vi.mock('./generated/services/ServiceOrdersService', () => ({
  ServiceOrdersService: { getAll: ordersGetAll },
}))

vi.mock('./lib/offlineStore', () => ({
  clearDraft,
  fromStoredMediaFile,
  getQueueSummary,
  loadDraft,
  saveDraft,
  toStoredMediaFile,
}))

vi.mock('./lib/offlineQueue', () => ({
  createClientSubmissionId,
  enqueueSubmission,
  processPendingQueue,
}))

vi.mock('./lib/referenceCache', () => ({
  readCachedReferenceData,
  writeCachedReferenceData,
}))

vi.mock('./weightOcr', () => ({
  preprocessImageForOcr: vi.fn(),
  runWeightOcr: vi.fn(),
}))

describe('App', () => {
  beforeEach(() => {
    vi.resetModules()
    vi.clearAllMocks()

    getContext.mockResolvedValue({ user: { userPrincipalName: 'operator@cora-environment.com' } })
    customersGetAll.mockResolvedValue({ data: [{ ID: 1, Title: 'Acme', field_1: 'HQ', field_2: 'Tower A' }] })
    driversGetAll.mockResolvedValue({ data: [{ ID: 2, Title: 'Jordan' }] })
    vehiclesGetAll.mockResolvedValue({ data: [{ ID: 3, Title: 'Truck 12', VehicleNumber: 'TRK-012' }] })
    wasteGetAll.mockResolvedValue({ data: [{ ID: 4, Title: 'General Waste' }] })
    ordersGetAll.mockResolvedValue({ data: [] })
    loadDraft.mockResolvedValue(null)
    readCachedReferenceData.mockResolvedValue(null)
    getQueueSummary.mockResolvedValue({ pendingCount: 0, lastError: '' })
    fromStoredMediaFile.mockResolvedValue(null)
    clearDraft.mockResolvedValue(undefined)
    saveDraft.mockResolvedValue(undefined)
    toStoredMediaFile.mockResolvedValue(null)
    enqueueSubmission.mockResolvedValue(undefined)
    processPendingQueue.mockResolvedValue(undefined)
    writeCachedReferenceData.mockResolvedValue(undefined)
  })

  it('renders the wizard when live reference data loads successfully', async () => {
    const { default: App } = await import('./App')

    render(<App />)

    await screen.findByRole('heading', { name: 'Create Service Order' })
    await screen.findByText('Everything is up to date')
    expect(screen.getByText('Reference data is live and ready.')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Continue to Customer/i })).toBeInTheDocument()
    expect(writeCachedReferenceData).toHaveBeenCalledWith(expect.objectContaining({
      id: 'main',
      customerAreas: [{ ID: 1, Title: 'Acme', field_1: 'HQ', field_2: 'Tower A' }],
    }))
  })

  it('falls back to cached reference data when the live load fails', async () => {
    customersGetAll.mockRejectedValue(new Error('SharePoint unavailable'))
    readCachedReferenceData.mockResolvedValue({
      id: 'main',
      fetchedAt: '2026-04-24T10:00:00.000Z',
      customerAreas: [{ ID: 9, Title: 'Cached Customer', field_1: 'Cached HQ', field_2: 'Cached Tower' }],
      drivers: [],
      vehicles: [],
      categories: [],
    })

    const { default: App } = await import('./App')

    render(<App />)

    await screen.findByRole('heading', { name: 'Create Service Order' })
    await waitFor(() => {
      expect(screen.getByText(/Using cached reference data/)).toBeInTheDocument()
    })
    expect(screen.queryByText('Reference Data Unavailable')).not.toBeInTheDocument()
  })
})
