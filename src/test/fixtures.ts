import type { PersistedDraft, PersistedForm, QueuedSubmission, SubmissionPayload } from '../lib/offlineStore'

export const makePersistedForm = (overrides: Partial<PersistedForm> = {}): PersistedForm => ({
  Title: 'SO-TEST-001',
  Customer: 'Acme HQ',
  CustomerName: 'Acme',
  CustomerLocation: 'HQ',
  CustomerTenant: 'Tower A',
  CustomerLevel4: '',
  CustomerLevel5: '',
  CustomerLevel6: '',
  CustomerLevel7: '',
  CustomerLevel8: '',
  CustomerLevel9: '',
  CustomerLevel10: '',
  IsAdhocCustomer: false,
  DriverName: 'Jordan',
  VehicleNumber: 'TRK-001',
  DateOfCollection: '2026-04-24',
  WasteItems: [
    {
      id: 'waste-1',
      WasteCategory: 'General Waste',
      Tonnage: '12.5',
      scaleOcrStatus: 'done',
      scaleOcrSuggestion: '12.5',
    },
  ],
  Notes: 'Fixture payload',
  ...overrides,
})

export const makeSubmissionPayload = (overrides: Partial<SubmissionPayload> = {}): SubmissionPayload => ({
  form: makePersistedForm(),
  signatureDataUrl: 'data:image/png;base64,c2lnbmF0dXJl',
  beforePhoto: null,
  afterPhoto: null,
  ...overrides,
})

export const makeDraft = (overrides: Partial<PersistedDraft> = {}): PersistedDraft => ({
  id: 'active',
  clientSubmissionId: 'submission-test-001',
  step: 2,
  updatedAt: '2026-04-24T10:00:00.000Z',
  locked: false,
  submissionStatus: 'draft',
  payload: makeSubmissionPayload(),
  ...overrides,
})

export const makeQueuedSubmission = (overrides: Partial<QueuedSubmission> = {}): QueuedSubmission => {
  const payload = overrides.payload ?? makeSubmissionPayload()
  return {
    id: 'submission-test-001',
    clientSubmissionId: 'submission-test-001',
    orderTitle: payload.form.Title,
    createdAt: '2026-04-24T10:00:00.000Z',
    updatedAt: '2026-04-24T10:00:00.000Z',
    status: 'queued',
    attempts: 0,
    payload,
    ...overrides,
  }
}
