import { describe, expect, it, vi } from 'vitest'
import type { Drivers1Read, Drivers1Write } from '../generated/models/Drivers1Model'

vi.mock('../generated/services/Customer_area_data_clean_finalService', () => ({
  Customer_area_data_clean_finalService: {},
}))

vi.mock('../generated/services/Drivers1Service', () => ({
  Drivers1Service: {},
}))

vi.mock('../generated/services/VehiclesService', () => ({
  VehiclesService: {},
}))

import {
  type ReferenceService,
  toCustomerRows,
  upsertDriversReferenceData,
  upsertReferenceRows,
} from './referenceDataUpsert'

const createService = <T extends { ID?: number } & Record<string, unknown>>(existing: T[] = []) => ({
  getAll: vi.fn(async () => ({ data: existing })),
  create: vi.fn(async (record: Omit<T, 'ID'>) => ({ data: { ...record, ID: 900 } as T })),
  update: vi.fn(async (id: string, changedFields: Partial<Omit<T, 'ID'>>) => ({
    data: { ...existing.find((row) => String(row.ID) === id), ...changedFields, ID: Number(id) } as T,
  })),
})

const asDriverService = (service: unknown) =>
  service as unknown as ReferenceService<Drivers1Read, Drivers1Write>

describe('referenceDataUpsert', () => {
  it('creates missing rows, updates changed casing/spacing, and skips exact matches', async () => {
    const service = createService([
      { ID: 1, Title: 'Hisham' },
      { ID: 2, Title: 'SAMUDI' },
    ])

    const report = await upsertDriversReferenceData([
      ' HISHAM ',
      'SAMUDI',
      'Mohd Rafiq',
    ], asDriverService(service))

    expect(service.getAll).toHaveBeenCalledWith({ top: 5000 })
    expect(service.update).toHaveBeenCalledWith('1', { Title: 'HISHAM' })
    expect(service.create).toHaveBeenCalledWith({ Title: 'Mohd Rafiq' })
    expect(report.updated).toHaveLength(1)
    expect(report.created).toHaveLength(1)
    expect(report.skipped).toHaveLength(1)
    expect(report.failed).toHaveLength(0)
  })

  it('rejects duplicate rows in the import payload without creating duplicates', async () => {
    const service = createService()

    const report = await upsertDriversReferenceData(['HISHAM', ' hisham '], asDriverService(service))

    expect(service.create).toHaveBeenCalledTimes(1)
    expect(report.created).toHaveLength(1)
    expect(report.failed).toEqual([
      expect.objectContaining({
        key: 'hisham',
        message: 'Row 2 duplicates import row 1.',
      }),
    ])
  })

  it('fails safely when existing SharePoint rows are already duplicated', async () => {
    const service = createService([
      { ID: 1, Title: 'HISHAM' },
      { ID: 2, Title: ' hisham ' },
    ])

    const report = await upsertDriversReferenceData(['Hisham'], asDriverService(service))

    expect(service.create).not.toHaveBeenCalled()
    expect(service.update).not.toHaveBeenCalled()
    expect(report.failed).toEqual([
      expect.objectContaining({
        message: 'Multiple existing rows match "Hisham". Resolve duplicates before importing.',
      }),
    ])
  })

  it('validates required key fields before writing', async () => {
    const service = createService()

    const report = await upsertReferenceRows('vehicles', service, [{ Title: ' ' }], ['Title'])

    expect(service.getAll).not.toHaveBeenCalled()
    expect(service.create).not.toHaveBeenCalled()
    expect(report.failed).toEqual([
      expect.objectContaining({
        message: 'Row 1 is missing required field "Title".',
      }),
    ])
  })

  it('normalizes customer import rows to generated list fields', () => {
    expect(toCustomerRows([{
      Title: ' Acme ',
      field_1: ' HQ ',
      field_2: ' Tower A ',
      field_4: ' Keep ',
    }])).toEqual([{
      Title: 'Acme',
      field_1: 'HQ',
      field_2: 'Tower A',
      field_3: '',
      field_4: 'Keep',
      field_5: '',
      field_6: '',
      field_7: '',
      field_8: '',
      field_9: '',
    }])
  })
})
