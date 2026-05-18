import { Effect } from "effect"

export const PDF_RENDERED_MIME = "image/png" as const
export const DEFAULT_PDF_PAGE_LIMIT = 5
export const MAX_PDF_PAGE_LIMIT = 10
export const PDF_MAX_RENDER_DIMENSION = 2000
const PDF_RENDER_SCALE = 2

export interface RenderedPdfPage {
  page: number
  mime: typeof PDF_RENDERED_MIME
  bytes: Uint8Array
}

export interface RenderPdfPagesResult {
  totalPages: number
  pages: RenderedPdfPage[]
  truncated: boolean
  nextOffset?: number
}

export interface RenderPdfPagesOptions {
  data: Uint8Array
  filename?: string
  offset?: number
  limit?: number
}

async function loadCanvas() {
  const canvas = await import("@napi-rs/canvas")
  const global = globalThis as Record<string, unknown>
  global.DOMMatrix ??= canvas.DOMMatrix
  global.ImageData ??= canvas.ImageData
  global.Path2D ??= canvas.Path2D
  return canvas
}

function message(cause: unknown) {
  return cause instanceof Error ? cause.message : String(cause)
}

export const renderPdfPages = Effect.fn("PdfRender.renderPdfPages")(function* (opts: RenderPdfPagesOptions) {
  return yield* Effect.tryPromise({
    try: async () => {
      const canvasModule = await loadCanvas()
      const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs")
      const loadingTask = pdfjs.getDocument({
        data: new Uint8Array(opts.data),
        disableWorker: true,
        useSystemFonts: true,
      } as never)

      const document = await loadingTask.promise
      try {
        const totalPages = document.numPages
        const offset = Math.max(1, opts.offset ?? 1)
        if (offset > totalPages) throw new Error(`Offset ${offset} is out of range for this PDF (${totalPages} pages)`)

        const requestedLimit = Math.max(1, opts.limit ?? DEFAULT_PDF_PAGE_LIMIT)
        const pageLimit = Math.min(requestedLimit, MAX_PDF_PAGE_LIMIT)
        const endPage = Math.min(totalPages, offset + pageLimit - 1)
        const pages: RenderedPdfPage[] = []

        for (let pageNumber = offset; pageNumber <= endPage; pageNumber++) {
          const page = await document.getPage(pageNumber)
          const baseViewport = page.getViewport({ scale: 1 })
          const scale = Math.max(
            0.1,
            Math.min(
              PDF_RENDER_SCALE,
              PDF_MAX_RENDER_DIMENSION / baseViewport.width,
              PDF_MAX_RENDER_DIMENSION / baseViewport.height,
            ),
          )
          const viewport = page.getViewport({ scale })
          const canvas = canvasModule.createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height))
          const canvasContext = canvas.getContext("2d")

          await page.render({ canvas: canvas as never, canvasContext: canvasContext as never, viewport }).promise
          pages.push({ page: pageNumber, mime: PDF_RENDERED_MIME, bytes: new Uint8Array(canvas.toBuffer("image/png")) })
          page.cleanup()
        }

        const nextOffset = endPage < totalPages ? endPage + 1 : undefined
        return {
          totalPages,
          pages,
          truncated: requestedLimit > pageLimit || nextOffset !== undefined,
          ...(nextOffset ? { nextOffset } : {}),
        }
      } finally {
        await document.destroy()
      }
    },
    catch: (cause) => new Error(`Failed to render PDF${opts.filename ? ` ${opts.filename}` : ""}: ${message(cause)}`),
  })
})
