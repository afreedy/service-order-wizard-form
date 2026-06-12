import type { Customer_area_data_clean_finalRead, Customer_area_data_clean_finalWrite } from '../generated/models/Customer_area_data_clean_finalModel'
import type { Drivers1Read, Drivers1Write } from '../generated/models/Drivers1Model'
import type { VehiclesRead, VehiclesWrite } from '../generated/models/VehiclesModel'
import { Customer_area_data_clean_finalService } from '../generated/services/Customer_area_data_clean_finalService'
import { Drivers1Service } from '../generated/services/Drivers1Service'
import { VehiclesService } from '../generated/services/VehiclesService'

type ReferenceRead = { ID?: number }
type ReferenceWrite = object

export interface ReferenceService<Read extends ReferenceRead, Write extends ReferenceWrite> {
  getAll: (options?: { top?: number }) => Promise<{ data?: Read[] }>
  create: (record: Omit<Write, 'ID'>) => Promise<{ data?: Read }>
  update: (id: string, changedFields: Partial<Omit<Write, 'ID'>>) => Promise<{ data?: Read }>
}

export type UpsertStatus = 'created' | 'updated' | 'skipped' | 'failed'

export interface UpsertLogEntry {
  status: UpsertStatus
  list: string
  key: string
  id?: number
  message: string
  changes?: Record<string, { from: string; to: string }>
}

export interface UpsertReport {
  created: UpsertLogEntry[]
  updated: UpsertLogEntry[]
  skipped: UpsertLogEntry[]
  failed: UpsertLogEntry[]
}

export interface ReferenceDataImportConfig {
  drivers?: Array<string | { Title?: string; name?: string }>
  vehicles?: Array<string | { Title?: string; vehicleNumber?: string; VehicleNumber?: string }>
  customers?: CustomerReferenceInput[]
}

export type CustomerReferenceInput = Partial<Pick<
  Customer_area_data_clean_finalWrite,
  'Title' | 'field_1' | 'field_2' | 'field_3' | 'field_4' | 'field_5' | 'field_6' | 'field_7' | 'field_8' | 'field_9'
>>

const CUSTOMER_FIELDS = ['Title', 'field_1', 'field_2', 'field_3', 'field_4', 'field_5', 'field_6', 'field_7', 'field_8', 'field_9'] as const
const CUSTOMER_KEY_FIELDS = ['Title', 'field_1', 'field_2'] as const

const createEmptyReport = (): UpsertReport => ({
  created: [],
  updated: [],
  skipped: [],
  failed: [],
})

const asText = (value: unknown) => String(value ?? '').trim()
const normalize = (value: unknown) => asText(value).replace(/\s+/g, ' ').toLowerCase()
const getField = (record: ReferenceWrite, field: string) => (record as Record<string, unknown>)[field]
const keyOf = (record: ReferenceWrite, fields: readonly string[]) => fields.map((field) => normalize(getField(record, field))).join(' | ')
const displayKeyOf = (record: ReferenceWrite, fields: readonly string[]) => fields.map((field) => asText(getField(record, field)) || '<blank>').join(' / ')

const pickChangedFields = <Write extends ReferenceWrite>(
  existing: ReferenceRead & ReferenceWrite,
  incoming: Write,
) => Object.entries(incoming).reduce<Record<string, { from: string; to: string }>>((changes, [field, value]) => {
  const current = asText(getField(existing, field))
  const next = asText(value)
  if (current !== next) changes[field] = { from: current, to: next }
  return changes
}, {})

const appendLog = (report: UpsertReport, entry: UpsertLogEntry) => {
  report[entry.status].push(entry)
}

const validateRows = <Write extends ReferenceWrite>(
  listName: string,
  rows: Write[],
  keyFields: readonly string[],
  report: UpsertReport,
) => {
  const seen = new Map<string, number>()
  return rows.filter((row, index) => {
    const missingKey = keyFields.find((field) => !asText(getField(row, field)))
    const displayKey = displayKeyOf(row, keyFields)
    if (missingKey) {
      appendLog(report, {
        status: 'failed',
        list: listName,
        key: displayKey,
        message: `Row ${index + 1} is missing required field "${missingKey}".`,
      })
      return false
    }

    const normalizedKey = keyOf(row, keyFields)
    const firstRow = seen.get(normalizedKey)
    if (firstRow !== undefined) {
      appendLog(report, {
        status: 'failed',
        list: listName,
        key: displayKey,
        message: `Row ${index + 1} duplicates import row ${firstRow + 1}.`,
      })
      return false
    }
    seen.set(normalizedKey, index)
    return true
  })
}

export const upsertReferenceRows = async <Read extends ReferenceRead, Write extends ReferenceWrite>(
  listName: string,
  service: ReferenceService<Read, Write>,
  rows: Write[],
  keyFields: readonly string[],
): Promise<UpsertReport> => {
  const report = createEmptyReport()
  const validRows = validateRows(listName, rows, keyFields, report)
  if (validRows.length === 0) return report

  const existingResult = await service.getAll({ top: 5000 })
  const existingRows = Array.isArray(existingResult.data) ? existingResult.data : []
  const existingByKey = existingRows.reduce<Map<string, Read[]>>((map, row) => {
    const key = keyOf(row, keyFields)
    if (!key.trim()) return map
    map.set(key, [...(map.get(key) ?? []), row])
    return map
  }, new Map())

  for (const row of validRows) {
    const normalizedKey = keyOf(row, keyFields)
    const displayKey = displayKeyOf(row, keyFields)
    const matches = existingByKey.get(normalizedKey) ?? []

    if (matches.length > 1) {
      appendLog(report, {
        status: 'failed',
        list: listName,
        key: displayKey,
        message: `Multiple existing rows match "${displayKey}". Resolve duplicates before importing.`,
      })
      continue
    }

    if (matches.length === 0) {
      const created = await service.create(row as Omit<Write, 'ID'>)
      appendLog(report, {
        status: 'created',
        list: listName,
        key: displayKey,
        id: created.data?.ID,
        message: `Created "${displayKey}".`,
      })
      continue
    }

    const existing = matches[0]
    const changes = pickChangedFields(existing, row)
    if (Object.keys(changes).length === 0) {
      appendLog(report, {
        status: 'skipped',
        list: listName,
        key: displayKey,
        id: existing.ID,
        message: `Skipped "${displayKey}" because it is already up to date.`,
      })
      continue
    }

    if (!existing.ID) {
      appendLog(report, {
        status: 'failed',
        list: listName,
        key: displayKey,
        message: `Existing row for "${displayKey}" has no SharePoint ID.`,
      })
      continue
    }

    const changedFields = Object.fromEntries(Object.entries(changes).map(([field, change]) => [field, change.to]))
    const updated = await service.update(String(existing.ID), changedFields as Partial<Omit<Write, 'ID'>>)
    appendLog(report, {
      status: 'updated',
      list: listName,
      key: displayKey,
      id: updated.data?.ID ?? existing.ID,
      changes,
      message: `Updated "${displayKey}".`,
    })
  }

  return report
}

export const toDriverRows = (drivers: ReferenceDataImportConfig['drivers'] = []) =>
  drivers.map((driver) => ({
    Title: typeof driver === 'string' ? asText(driver) : asText(driver.Title || driver.name),
  }))

export const toVehicleRows = (vehicles: ReferenceDataImportConfig['vehicles'] = []) =>
  vehicles.map((vehicle) => ({
    Title: typeof vehicle === 'string' ? asText(vehicle) : asText(vehicle.Title || vehicle.VehicleNumber || vehicle.vehicleNumber),
  }))

export const toCustomerRows = (customers: CustomerReferenceInput[] = []) =>
  customers.map((customer) => Object.fromEntries(
    CUSTOMER_FIELDS.map((field) => [field, asText(customer[field])]),
  ) as CustomerReferenceInput)

export const upsertDriversReferenceData = (
  drivers: ReferenceDataImportConfig['drivers'] = [],
  service: ReferenceService<Drivers1Read, Drivers1Write> = Drivers1Service,
) =>
  upsertReferenceRows<Drivers1Read, Drivers1Write>(
    'drivers1',
    service,
    toDriverRows(drivers) as Drivers1Write[],
    ['Title'],
  )

export const upsertVehiclesReferenceData = (
  vehicles: ReferenceDataImportConfig['vehicles'] = [],
  service: ReferenceService<VehiclesRead, VehiclesWrite> = VehiclesService,
) =>
  upsertReferenceRows<VehiclesRead, VehiclesWrite>(
    'vehicles',
    service,
    toVehicleRows(vehicles) as VehiclesWrite[],
    ['Title'],
  )

export const upsertCustomersReferenceData = (
  customers: CustomerReferenceInput[] = [],
  service: ReferenceService<Customer_area_data_clean_finalRead, Customer_area_data_clean_finalWrite> = Customer_area_data_clean_finalService,
) =>
  upsertReferenceRows<Customer_area_data_clean_finalRead, Customer_area_data_clean_finalWrite>(
    'customer-area-data-clean-final',
    service,
    toCustomerRows(customers) as Customer_area_data_clean_finalWrite[],
    CUSTOMER_KEY_FIELDS,
  )

export const upsertReferenceData = async (config: ReferenceDataImportConfig) => ({
  drivers: await upsertDriversReferenceData(config.drivers),
  vehicles: await upsertVehiclesReferenceData(config.vehicles),
  customers: await upsertCustomersReferenceData(config.customers),
})
