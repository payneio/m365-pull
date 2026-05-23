/** Sanitize a path-derived suggested filename to a safe, flat name. */
function sanitize(name: string): string {
  return name
    .replace(/[/\\]/g, "-") // path separators
    .replace(/[^a-zA-Z0-9._-]/g, "-") // anything else not safe
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 200)
}

export interface SaveResult {
  saved: boolean
  reason?: "cancelled" | "unsupported" | string
}

/** Save a JS value as JSON via the browser's save-file dialog. */
export async function saveAsJson(
  suggestedName: string,
  data: unknown,
): Promise<SaveResult> {
  const cleaned = sanitize(suggestedName)
  const safeName = cleaned.endsWith(".json") ? cleaned : cleaned + ".json"

  const w = window as unknown as {
    showSaveFilePicker?: (opts: {
      suggestedName?: string
      types?: { description?: string; accept: Record<string, string[]> }[]
    }) => Promise<{
      createWritable: () => Promise<{
        write: (data: string) => Promise<void>
        close: () => Promise<void>
      }>
    }>
  }

  if (!w.showSaveFilePicker) {
    return { saved: false, reason: "unsupported" }
  }

  try {
    const handle = await w.showSaveFilePicker({
      suggestedName: safeName,
      types: [
        {
          description: "JSON",
          accept: { "application/json": [".json"] },
        },
      ],
    })
    const writable = await handle.createWritable()
    await writable.write(JSON.stringify(data, null, 2))
    await writable.close()
    return { saved: true }
  } catch (err) {
    if ((err as Error).name === "AbortError") {
      return { saved: false, reason: "cancelled" }
    }
    return { saved: false, reason: (err as Error).message }
  }
}
