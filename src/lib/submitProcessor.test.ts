import { beforeEach, describe, expect, it, vi } from 'vitest'
import { makePersistedForm, makeQueuedSubmission, makeSubmissionPayload } from '../test/fixtures'

const ServiceOrderProofQueueService = {
  create: vi.fn(),
  delete: vi.fn(),
  getAll: vi.fn(),
}

const ServiceOrdersService = {
  create: vi.fn(),
  delete: vi.fn(),
  getAll: vi.fn(),
  update: vi.fn(),
}

const ServiceOrderWasteItemsService = {
  create: vi.fn(),
  delete: vi.fn(),
  getAll: vi.fn(),
}

vi.mock('../generated/services/ServiceOrderProofQueueService', () => ({
  ServiceOrderProofQueueService,
}))

vi.mock('../generated/services/ServiceOrdersService', () => ({
  ServiceOrdersService,
}))

vi.mock('../generated/services/ServiceOrderWasteItemsService', () => ({
  ServiceOrderWasteItemsService,
}))

describe('submitProcessor', () => {
  beforeEach(() => {
    vi.resetModules()
    vi.clearAllMocks()
  })

  it('creates missing rows and escapes OData filters for quoted titles', async () => {
    const title = "SO-O'Reilly-001"
    const queuedSubmission = makeQueuedSubmission({
      payload: makeSubmissionPayload({
        form: makePersistedForm({
          Title: title,
          WasteItems: [{ id: 'waste-1', WasteCategory: 'Metal', Tonnage: '12.5' }],
        }),
      }),
    })

    ServiceOrdersService.getAll
      .mockResolvedValueOnce({ data: [] })
      .mockResolvedValueOnce({ data: [{ ID: 101, Title: title, WasteCategory: 'Metal', Tonnage: 12.5 }] })
    ServiceOrdersService.create.mockResolvedValue({ data: { ID: 101, Title: title } })
    ServiceOrderWasteItemsService.getAll
      .mockResolvedValueOnce({ data: [] })
      .mockResolvedValueOnce({ data: [{ ID: 201, Title: title, WasteCategory: 'Metal', Tonnage: 12.5 }] })
    ServiceOrderWasteItemsService.create.mockResolvedValue({ data: { ID: 201, Title: title } })
    ServiceOrderProofQueueService.getAll.mockResolvedValue({ data: [] })
    ServiceOrderProofQueueService.create.mockResolvedValue({ data: { ID: 301, OrderTitle: title } })

    const { processQueuedSubmission } = await import('./submitProcessor')
    const result = await processQueuedSubmission(queuedSubmission)

    expect(ServiceOrdersService.getAll).toHaveBeenCalledWith(expect.objectContaining({
      filter: "Title eq 'SO-O''Reilly-001'",
      top: 200,
    }))
    expect(ServiceOrdersService.create).toHaveBeenCalledWith(expect.objectContaining({
      Title: title,
      WasteCategory: 'Metal',
      Tonnage: 12.5,
    }))
    expect(ServiceOrderWasteItemsService.create).toHaveBeenCalledWith(expect.objectContaining({
      Title: title,
      WasteCategory: 'Metal',
      Tonnage: 12.5,
    }))
    expect(ServiceOrderProofQueueService.create).toHaveBeenCalledWith(expect.objectContaining({
      Title: `${title}-proof`,
      ServiceOrderId: '101',
      OrderTitle: title,
    }))
    expect(result).toEqual({
      createdOrders: [{ ID: 101, Title: title, WasteCategory: 'Metal', Tonnage: 12.5 }],
      proofQueueItem: { ID: 301, OrderTitle: title },
    })
  })

  it('sorts matching orders and waste items by numeric id before returning them', async () => {
    const title = 'SO-SORT-001'
    const queuedSubmission = makeQueuedSubmission({
      payload: makeSubmissionPayload({
        form: makePersistedForm({
          Title: title,
          WasteItems: [{ id: 'waste-1', WasteCategory: 'Metal', Tonnage: '8' }],
        }),
      }),
    })

    ServiceOrdersService.getAll
      .mockResolvedValueOnce({
        data: [
          { ID: 9, Title: 'other', WasteCategory: 'Metal', Tonnage: 8 },
          { ID: 5, Title: title, WasteCategory: 'Metal', Tonnage: 8 },
          { ID: 2, Title: title, WasteCategory: 'Metal', Tonnage: 8 },
        ],
      })
      .mockResolvedValueOnce({
        data: [
          { ID: 5, Title: title, WasteCategory: 'Metal', Tonnage: 8 },
          { ID: 2, Title: title, WasteCategory: 'Metal', Tonnage: 8 },
        ],
      })
    ServiceOrderWasteItemsService.getAll
      .mockResolvedValueOnce({
        data: [
          { ID: 7, Title: title, WasteCategory: 'Metal', Tonnage: 8 },
          { ID: 3, Title: title, WasteCategory: 'Metal', Tonnage: 8 },
          { ID: 1, Title: 'other', WasteCategory: 'Metal', Tonnage: 8 },
        ],
      })
      .mockResolvedValueOnce({
        data: [
          { ID: 7, Title: title, WasteCategory: 'Metal', Tonnage: 8 },
          { ID: 3, Title: title, WasteCategory: 'Metal', Tonnage: 8 },
        ],
      })
    ServiceOrderProofQueueService.getAll.mockResolvedValue({
      data: [{ ID: 15, OrderTitle: title }],
    })

    const { processQueuedSubmission } = await import('./submitProcessor')
    const result = await processQueuedSubmission(queuedSubmission)

    expect(result.createdOrders.map((order) => order.ID)).toEqual([2, 5])
  })

  it('reuses an existing proof queue item and updates service orders with proof URLs', async () => {
    const title = 'SO-EXISTING-PROOF'
    const queuedSubmission = makeQueuedSubmission({
      payload: makeSubmissionPayload({
        form: makePersistedForm({
          Title: title,
          WasteItems: [{ id: 'waste-1', WasteCategory: '', Tonnage: '' }],
        }),
      }),
    })

    ServiceOrdersService.getAll
      .mockResolvedValueOnce({ data: [] })
      .mockResolvedValueOnce({ data: [{ ID: 401, Title: title }] })
    ServiceOrdersService.create.mockResolvedValue({ data: { ID: 401, Title: title } })
    ServiceOrdersService.update.mockResolvedValue({
      data: {
        ID: 401,
        Title: title,
        SignatureUrl: 'sig-url',
        BeforePhotoUrl: 'before-url',
        AfterPhotoUrl: 'after-url',
      },
    })
    ServiceOrderProofQueueService.getAll.mockResolvedValue({
      data: [{
        ID: 501,
        OrderTitle: title,
        SignatureUrl: 'sig-url',
        BeforePhotoUrl: 'before-url',
        AfterPhotoUrl: 'after-url',
      }],
    })
    ServiceOrderWasteItemsService.getAll.mockResolvedValue({ data: [] })

    const { processQueuedSubmission } = await import('./submitProcessor')
    const result = await processQueuedSubmission(queuedSubmission)

    expect(ServiceOrderProofQueueService.create).not.toHaveBeenCalled()
    expect(ServiceOrdersService.update).toHaveBeenCalledWith('401', {
      SignatureUrl: 'sig-url',
      BeforePhotoUrl: 'before-url',
      AfterPhotoUrl: 'after-url',
    })
    expect(result.proofQueueItem).toEqual(expect.objectContaining({ ID: 501 }))
  })

  it('rolls back newly created records when confirmation fails later in the pipeline', async () => {
    const title = 'SO-ROLLBACK-001'
    const queuedSubmission = makeQueuedSubmission({
      payload: makeSubmissionPayload({
        form: makePersistedForm({
          Title: title,
          WasteItems: [{ id: 'waste-1', WasteCategory: 'Metal', Tonnage: '4.25' }],
        }),
      }),
    })

    ServiceOrdersService.getAll
      .mockResolvedValueOnce({ data: [] })
      .mockResolvedValueOnce({ data: [{ ID: 701, Title: title, WasteCategory: 'Metal', Tonnage: 4.25 }] })
    ServiceOrdersService.create.mockResolvedValue({ data: { ID: 701, Title: title } })
    ServiceOrderWasteItemsService.getAll
      .mockResolvedValueOnce({ data: [] })
      .mockResolvedValueOnce({ data: [] })
    ServiceOrderWasteItemsService.create.mockResolvedValue({ data: { ID: 801, Title: title } })

    const { processQueuedSubmission } = await import('./submitProcessor')

    await expect(processQueuedSubmission(queuedSubmission)).rejects.toThrow(
      `Could not confirm the waste line rows for "${title}".`,
    )
    expect(ServiceOrderWasteItemsService.delete).toHaveBeenCalledWith('801')
    expect(ServiceOrdersService.delete).toHaveBeenCalledWith('701')
  })

  it('uses fallback proof confirmation polling and uploads before/after images', async () => {
    vi.useFakeTimers()

    try {
      const title = 'SO-PROOF-POLL-001'
      const beforeBlob = new Blob(['before-bytes'], { type: 'image/jpeg' })
      const afterBlob = new Blob(['after-bytes'], { type: 'image/jpeg' })
      const queuedSubmission = makeQueuedSubmission({
        payload: makeSubmissionPayload({
          form: makePersistedForm({
            Title: title,
            WasteItems: [{ id: 'waste-1', WasteCategory: '', Tonnage: '' }],
          }),
          beforePhoto: {
            blob: beforeBlob,
            name: 'Before Photo!!.png',
            type: 'image/jpeg',
            lastModified: 111,
          },
          afterPhoto: {
            dataUrl: 'data:image/jpeg;base64,YWZ0ZXI=',
            name: 'After Photo!!.png',
            type: 'image/jpeg',
            lastModified: 222,
          },
        }),
      })

      ServiceOrdersService.getAll
        .mockResolvedValueOnce({ data: [] })
        .mockResolvedValueOnce({ data: [{ ID: 901, Title: title }] })
      ServiceOrdersService.create.mockResolvedValue({ data: { ID: 901, Title: title } })
      ServiceOrderWasteItemsService.getAll.mockResolvedValue({ data: [] })
      ServiceOrderProofQueueService.getAll
        .mockResolvedValueOnce({ data: [] })
        .mockResolvedValueOnce({ data: [] })
        .mockResolvedValueOnce({
          data: [{
            ID: 902,
            OrderTitle: title,
          }],
        })
      ServiceOrderProofQueueService.create.mockResolvedValue({ data: {} })

      const { processQueuedSubmission } = await import('./submitProcessor')
      const resultPromise = processQueuedSubmission(queuedSubmission)
      await vi.advanceTimersByTimeAsync(800)
      const result = await resultPromise

      expect(ServiceOrderProofQueueService.create).toHaveBeenCalledWith(expect.objectContaining({
        BeforePhotoFileName: `${title}-before-Before-Photo---png.jpg`,
        BeforePhotoBase64: btoa('before-bytes'),
        AfterPhotoFileName: `${title}-after-After-Photo---png.jpg`,
        AfterPhotoBase64: 'YWZ0ZXI=',
      }))
      expect(result).toEqual({
        createdOrders: [{ ID: 901, Title: title }],
        proofQueueItem: { ID: 902, OrderTitle: title },
      })
    } finally {
      vi.useRealTimers()
    }
  })

  it('avoids duplicate order and waste item creation when rows already exist', async () => {
    const title = 'SO-DEDUPE-001'
    const existingOrder = { ID: 1001, Title: title, WasteCategory: 'Metal', Tonnage: 8 }
    const existingWaste = { ID: 1002, Title: title, WasteCategory: 'Metal', Tonnage: 8 }
    const queuedSubmission = makeQueuedSubmission({
      payload: makeSubmissionPayload({
        form: makePersistedForm({
          Title: title,
          WasteItems: [{ id: 'waste-1', WasteCategory: 'Metal', Tonnage: '8' }],
        }),
      }),
    })

    ServiceOrdersService.getAll
      .mockResolvedValueOnce({ data: [existingOrder] })
      .mockResolvedValueOnce({ data: [existingOrder] })
    ServiceOrderWasteItemsService.getAll
      .mockResolvedValueOnce({ data: [existingWaste] })
      .mockResolvedValueOnce({ data: [existingWaste] })
    ServiceOrderProofQueueService.getAll.mockResolvedValue({
      data: [{ ID: 1003, OrderTitle: title }],
    })

    const { processQueuedSubmission } = await import('./submitProcessor')
    const result = await processQueuedSubmission(queuedSubmission)

    expect(ServiceOrdersService.create).not.toHaveBeenCalled()
    expect(ServiceOrderWasteItemsService.create).not.toHaveBeenCalled()
    expect(result.createdOrders).toEqual([existingOrder])
  })

  it('throws when created orders cannot be confirmed', async () => {
    const title = 'SO-NO-CONFIRM-001'
    const queuedSubmission = makeQueuedSubmission({
      payload: makeSubmissionPayload({
        form: makePersistedForm({
          Title: title,
          WasteItems: [{ id: 'waste-1', WasteCategory: '', Tonnage: '' }],
        }),
      }),
    })

    ServiceOrdersService.getAll
      .mockResolvedValueOnce({ data: [] })
      .mockResolvedValueOnce({ data: [] })
    ServiceOrdersService.create.mockResolvedValue({ data: { ID: 1101, Title: title } })

    const { processQueuedSubmission } = await import('./submitProcessor')

    await expect(processQueuedSubmission(queuedSubmission)).rejects.toThrow(
      `Could not confirm the created service order rows for "${title}".`,
    )
    expect(ServiceOrdersService.delete).toHaveBeenCalledWith('1101')
  })

  it('throws a proof confirmation error and appends rollback cleanup failures', async () => {
    vi.useFakeTimers()

    try {
      const title = 'SO-PROOF-FAIL-001'
      const queuedSubmission = makeQueuedSubmission({
        payload: makeSubmissionPayload({
          form: makePersistedForm({
            Title: title,
            WasteItems: [{ id: 'waste-1', WasteCategory: 'Metal', Tonnage: '2' }],
          }),
        }),
      })

      ServiceOrdersService.getAll
        .mockResolvedValueOnce({ data: [] })
        .mockResolvedValueOnce({ data: [{ ID: 1201, Title: title, WasteCategory: 'Metal', Tonnage: 2 }] })
      ServiceOrdersService.create.mockResolvedValue({ data: { ID: 1201, Title: title } })
      ServiceOrderWasteItemsService.getAll
        .mockResolvedValueOnce({ data: [] })
        .mockResolvedValueOnce({ data: [{ ID: 1202, Title: title, WasteCategory: 'Metal', Tonnage: 2 }] })
      ServiceOrderWasteItemsService.create.mockResolvedValue({ data: { ID: 1202, Title: title } })
      ServiceOrderProofQueueService.getAll.mockResolvedValue({ data: [] })
      ServiceOrderProofQueueService.create.mockResolvedValue({ data: {} })
      ServiceOrderWasteItemsService.delete.mockRejectedValue(new Error('waste cleanup failed'))
      ServiceOrdersService.delete.mockRejectedValue(new Error('order cleanup failed'))

      const { processQueuedSubmission } = await import('./submitProcessor')
      const resultPromise = processQueuedSubmission(queuedSubmission)
      const assertion = expect(resultPromise).rejects.toThrow(
        'Could not confirm the proof media queue item. Rollback needs manual cleanup. waste item 1202: waste cleanup failed service order 1201: order cleanup failed',
      )
      await vi.advanceTimersByTimeAsync(2000)

      await assertion
    } finally {
      vi.useRealTimers()
    }
  })

  it('rolls back the proof queue item when a later order update fails', async () => {
    const title = 'SO-PROOF-ROLLBACK-001'
    const queuedSubmission = makeQueuedSubmission({
      payload: makeSubmissionPayload({
        form: makePersistedForm({
          Title: title,
          WasteItems: [{ id: 'waste-1', WasteCategory: '', Tonnage: '' }],
        }),
      }),
    })

    ServiceOrdersService.getAll
      .mockResolvedValueOnce({ data: [] })
      .mockResolvedValueOnce({ data: [{ ID: 1301, Title: title }] })
    ServiceOrdersService.create.mockResolvedValue({ data: { ID: 1301, Title: title } })
    ServiceOrdersService.update.mockRejectedValue(new Error('update failed'))
    ServiceOrderWasteItemsService.getAll.mockResolvedValue({ data: [] })
    ServiceOrderProofQueueService.getAll.mockResolvedValue({
      data: [{
        ID: 1302,
        OrderTitle: title,
        SignatureUrl: 'sig-url',
      }],
    })

    const { processQueuedSubmission } = await import('./submitProcessor')

    await expect(processQueuedSubmission(queuedSubmission)).rejects.toThrow('update failed')
    expect(ServiceOrderProofQueueService.delete).toHaveBeenCalledWith('1302')
  })
})
