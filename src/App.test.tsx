import { render, screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
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
    sessionStorage.clear()
    window.history.pushState({}, '', '/')
    HTMLElement.prototype.scrollIntoView = vi.fn()

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

  it('searches customer hierarchy paths and fills all customer fields from the selected row', async () => {
    const user = userEvent.setup()
    customersGetAll.mockResolvedValue({
      data: [
        { ID: 1, Title: 'Acme', field_1: 'HQ', field_2: 'Tower A' },
        {
          ID: 2,
          Title: 'Beta',
          field_1: 'Plant',
          field_2: 'Zone B',
          field_3: 'Lab',
          field_4: 'Cold Room',
          field_5: 'Dock 9',
        },
        {
          ID: 3,
          Title: 'Beta',
          field_1: 'Plant',
          field_2: 'Zone B',
          field_3: 'Lab',
          field_4: 'Cold Room',
          field_5: 'Dock 9',
        },
        {
          ID: 4,
          Title: 'Beta',
          field_1: 'Warehouse',
          field_2: 'Zone C',
        },
      ],
    })
    const { default: App } = await import('./App')

    render(<App />)

    await user.click(await screen.findByRole('button', { name: /Continue to Customer/i }))
    await user.click(screen.getByRole('combobox', { name: 'Customer' }))
    await user.type(screen.getByRole('combobox', { name: 'Search any customer level...' }), 'Dock 9')

    const fullPath = 'Beta / Plant / Zone B / Lab / Cold Room / Dock 9'
    expect(screen.getAllByRole('option', { name: fullPath })).toHaveLength(1)
    await user.click(screen.getByRole('option', { name: fullPath }))

    expect(screen.getByRole('combobox', { name: 'Customer' })).toHaveTextContent(fullPath)
    expect(screen.getByRole('combobox', { name: 'Location' })).toHaveTextContent('Plant')
    expect(screen.getByRole('combobox', { name: 'Sub-location' })).toHaveTextContent('Zone B')
    expect(screen.getByRole('combobox', { name: 'Level 4' })).toHaveTextContent('Lab')
    expect(screen.getByRole('combobox', { name: 'Level 5' })).toHaveTextContent('Cold Room')
    expect(screen.getByRole('combobox', { name: 'Level 6' })).toHaveTextContent('Dock 9')

    await user.click(screen.getByRole('combobox', { name: 'Location' }))
    expect(screen.getByRole('option', { name: 'Warehouse' })).toBeInTheDocument()
  })

  it('supports keyboard selection in customer hierarchy autocomplete', async () => {
    const user = userEvent.setup()
    customersGetAll.mockResolvedValue({
      data: [
        {
          ID: 2,
          Title: 'Beta',
          field_1: 'Plant',
          field_2: 'Zone B',
          field_3: 'Lab',
          field_4: 'Cold Room',
          field_5: 'Dock 9',
        },
      ],
    })
    const { default: App } = await import('./App')

    render(<App />)

    await user.click(await screen.findByRole('button', { name: /Continue to Customer/i }))
    await user.click(screen.getByRole('combobox', { name: 'Customer' }))
    await user.type(screen.getByRole('combobox', { name: 'Search any customer level...' }), 'cold')
    await user.keyboard('{ArrowDown}{Enter}')

    expect(screen.getByRole('combobox', { name: 'Customer' })).toHaveTextContent('Beta / Plant / Zone B / Lab / Cold Room / Dock 9')
  })

  it('shows admin customer autocomplete suggestions and filters from the selected suggestion', async () => {
    const user = userEvent.setup()
    sessionStorage.setItem('cora-admin-unlocked', 'true')
    getContext.mockResolvedValue({ user: { userPrincipalName: 'bizdev@cora-environment.com' } })
    customersGetAll.mockResolvedValue({
      data: [
        { ID: 1, Title: 'Acme', field_1: 'HQ', field_2: 'Tower A' },
        { ID: 2, Title: 'Beta', field_1: 'Plant', field_2: 'Zone B', field_3: 'Lab' },
        { ID: 3, Title: 'Cedar', field_1: 'Depot', field_2: 'Bay 4' },
      ],
    })
    const { default: App } = await import('./App')

    render(<App />)

    await screen.findByRole('heading', { name: 'Create Service Order' })
    await user.click(screen.getAllByRole('button', { name: /Customers/i })[0])
    await screen.findByRole('heading', { name: 'Customers' })
    await user.type(screen.getByRole('combobox', { name: 'Search customers' }), 'lab')
    await user.click(await screen.findByRole('option', { name: 'Beta / Plant / Zone B / Lab' }))

    expect(screen.getByRole('combobox', { name: 'Search customers' })).toHaveValue('Beta / Plant / Zone B / Lab')
    expect(screen.getByText('1 of 3 shown')).toBeInTheDocument()
    expect(screen.getByText('Beta')).toBeInTheDocument()
    expect(screen.queryByText('Acme')).not.toBeInTheDocument()
  })

  it('shows order autocomplete suggestions and filters from the selected suggestion', async () => {
    const user = userEvent.setup()
    sessionStorage.setItem('cora-admin-unlocked', 'true')
    getContext.mockResolvedValue({ user: { userPrincipalName: 'bizdev@cora-environment.com' } })
    ordersGetAll.mockResolvedValue({
      data: [
        {
          ID: 11,
          Title: 'SO-001',
          CustomerName: 'Acme',
          DriverName: 'Jordan',
          VehicleNumber: 'Truck 12',
          DateOfCollection: '2026-04-28',
        },
        {
          ID: 12,
          Title: 'SO-002',
          CustomerName: 'Beta',
          DriverName: 'Riley',
          VehicleNumber: 'Van 3',
          DateOfCollection: '2026-04-27',
        },
      ],
    })
    const { default: App } = await import('./App')

    render(<App />)

    await screen.findByRole('heading', { name: 'Create Service Order' })
    await user.click(screen.getByRole('button', { name: /All Orders/i }))
    await screen.findByRole('heading', { name: 'Service Orders' })
    await user.type(screen.getByRole('combobox', { name: 'Search orders' }), 'truck')
    await user.click(await screen.findByRole('option', { name: 'Truck 12' }))

    expect(screen.getByRole('combobox', { name: 'Search orders' })).toHaveValue('Truck 12')
    expect(screen.getByText('1 of 2 shown')).toBeInTheDocument()
    expect(screen.getByText('SO-001')).toBeInTheDocument()
    expect(screen.queryByText('SO-002')).not.toBeInTheDocument()
  })

  it('opens the Mindef QR mock from the sidebar', async () => {
    const user = userEvent.setup()
    const { default: App } = await import('./App')

    render(<App />)

    await screen.findByRole('heading', { name: 'Create Service Order' })
    await user.click(screen.getByRole('button', { name: /Mindef QR Mock/i }))

    expect(await screen.findByRole('heading', { name: 'Mindef QR Mock' })).toBeInTheDocument()
    expect(screen.getByLabelText('Customer')).toHaveValue('')
    expect(screen.getByLabelText('Location')).toHaveValue('')
    expect(screen.getByLabelText('Sub-location')).toHaveValue('')
  })

  it('prefills the Mindef mock from QR URL parameters as read-only values', async () => {
    window.history.pushState(
      {},
      '',
      '/?view=mindef&customer=Mindef&location=Kranji%20Camp&sublocation=Cookhouse%201',
    )
    const { default: App } = await import('./App')

    render(<App />)

    expect(await screen.findByRole('heading', { name: 'Mindef QR Mock' })).toBeInTheDocument()
    expect(screen.getByLabelText('Customer')).toHaveValue('Mindef')
    expect(screen.getByLabelText('Location')).toHaveValue('Kranji Camp')
    expect(screen.getByLabelText('Sub-location')).toHaveValue('Cookhouse 1')
    expect(screen.getByLabelText('Customer')).toHaveAttribute('readonly')
    expect(screen.getByLabelText('Location')).toHaveAttribute('readonly')
    expect(screen.getByLabelText('Sub-location')).toHaveAttribute('readonly')
  })

  it('prefills the Mindef mock from Power Apps context query parameters', async () => {
    getContext.mockResolvedValue({
      user: { userPrincipalName: 'operator@cora-environment.com' },
      app: {
        queryParams: {
          view: 'mindef',
          customer: 'Mindef',
          location: 'Kranji Camp',
          sublocation: 'Cookhouse 1',
        },
      },
    })
    const { default: App } = await import('./App')

    render(<App />)

    expect(await screen.findByRole('heading', { name: 'Mindef QR Mock' })).toBeInTheDocument()
    expect(screen.getByLabelText('Customer')).toHaveValue('Mindef')
    expect(screen.getByLabelText('Location')).toHaveValue('Kranji Camp')
    expect(screen.getByLabelText('Sub-location')).toHaveValue('Cookhouse 1')
  })

  it('submits the Mindef mock locally without queueing or SharePoint writes', async () => {
    const user = userEvent.setup()
    window.history.pushState(
      {},
      '',
      '/?view=mindef&customer=Mindef&location=Kranji%20Camp&sublocation=Cookhouse%201',
    )
    const { default: App } = await import('./App')

    render(<App />)

    await screen.findByRole('heading', { name: 'Mindef QR Mock' })
    await user.type(screen.getByLabelText('Tonnage'), '12.5')

    enqueueSubmission.mockClear()
    processPendingQueue.mockClear()
    await user.click(screen.getByRole('button', { name: /Submit Mock/i }))

    expect(await screen.findByRole('heading', { name: 'Mindef Mock Submitted' })).toBeInTheDocument()
    expect(screen.getByText('Mindef / Kranji Camp / Cookhouse 1')).toBeInTheDocument()
    expect(screen.getByText('12.5 kg')).toBeInTheDocument()
    expect(enqueueSubmission).not.toHaveBeenCalled()
    expect(processPendingQueue).not.toHaveBeenCalled()
  })

  it('renders Excel-backed SNF fields and dropdown values', async () => {
    const user = userEvent.setup()
    const { default: App } = await import('./App')

    render(<App />)

    await screen.findByRole('heading', { name: 'Create Service Order' })
    await user.click(screen.getByRole('button', { name: /SNF Mock/i }))
    await screen.findByRole('heading', { name: 'Create SNF' })

    expect(screen.getByLabelText(/Notice type/i)).toBeInTheDocument()
    expect(screen.getByLabelText(/Company/i)).toBeInTheDocument()
    expect(screen.getByLabelText(/Contact person/i)).toBeInTheDocument()
    expect(screen.getByLabelText(/Service start date/i)).toBeInTheDocument()

    await user.click(screen.getByRole('combobox', { name: /Notice type/i }))
    expect(await screen.findByRole('option', { name: 'NEW CONTRACT' })).toBeInTheDocument()
    expect(screen.getByRole('option', { name: 'RENEWAL' })).toBeInTheDocument()
    expect(screen.getByRole('option', { name: 'TERMINATION OF SERVICE' })).toBeInTheDocument()
    expect(screen.getByRole('option', { name: 'OTHERS' })).toBeInTheDocument()
    await user.keyboard('{Escape}')

    await user.click(screen.getByRole('combobox', { name: /Type of service/i }))
    expect(await screen.findByRole('option', { name: 'REL' })).toBeInTheDocument()
    expect(screen.getByRole('option', { name: 'RECYCLING' })).toBeInTheDocument()
    await user.keyboard('{Escape}')

    await user.click(screen.getByRole('combobox', { name: /Service frequency/i }))
    expect(await screen.findByRole('option', { name: 'Twice Weeky' })).toBeInTheDocument()
    expect(screen.getByRole('option', { name: 'Once-A-Month' })).toBeInTheDocument()
    await user.keyboard('{Escape}')

    await user.click(screen.getByRole('combobox', { name: /Bin type/i }))
    expect(await screen.findByRole('option', { name: 'Smart Weighing Scale' })).toBeInTheDocument()
    expect(screen.getByRole('option', { name: '660L Recycling' })).toBeInTheDocument()
    await user.keyboard('{Escape}')

    await user.click(screen.getByRole('button', { name: /Add service line/i }))
    expect(screen.getByText('Service line 2')).toBeInTheDocument()
    await user.click(screen.getByRole('button', { name: /Remove service line 2/i }))
    expect(screen.queryByText('Service line 2')).not.toBeInTheDocument()
  }, 10000)

  it('requires core Excel SNF fields before submitting a mock request', async () => {
    const user = userEvent.setup()
    const { default: App } = await import('./App')

    render(<App />)

    await screen.findByRole('heading', { name: 'Create Service Order' })
    await user.click(screen.getByRole('button', { name: /SNF Mock/i }))
    await screen.findByRole('heading', { name: 'Create SNF' })

    await user.clear(screen.getByLabelText(/Company/i))
    expect(screen.getByRole('button', { name: /Submit SNF/i })).toBeDisabled()

    await user.type(screen.getByLabelText(/Company/i), "L'occitane Singapore Pte Ltd")
    await user.clear(screen.getByLabelText(/Contact no/i))
    await user.type(screen.getByLabelText(/Contact no/i), '9123 4567')
    await user.clear(screen.getByLabelText(/No\. of bins/i))
    await user.type(screen.getByLabelText(/No\. of bins/i), '1')

    const submitButton = screen.getByRole('button', { name: /Submit SNF/i })
    expect(submitButton).toBeEnabled()
    await user.click(submitButton)

    expect(screen.getAllByText("L'occitane Singapore Pte Ltd").length).toBeGreaterThan(0)
    expect(screen.getByText(/RECYCLING - One Time - 1 x 660L Recycling/i)).toBeInTheDocument()
  }, 10000)
})
