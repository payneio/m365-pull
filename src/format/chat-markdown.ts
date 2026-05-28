// Chat-thread markdown renderer.
//
// Takes a list of Teams chat messages and produces an LLM-friendly markdown
// thread: messages grouped by date, HTML stripped to plain text, sender +
// time + body block per message.
//
// Pure function. No Graph calls.

import type { TeamsChatMessage } from "../sources/teams-chats"

const HTML_ENTITIES: Record<string, string> = {
  "&amp;": "&",
  "&lt;": "<",
  "&gt;": ">",
  "&quot;": '"',
  "&apos;": "'",
  "&nbsp;": " ",
  "&#39;": "'",
}

function decodeHtmlEntities(text: string): string {
  // Named entities + a couple of common numeric ones.
  for (const k of Object.keys(HTML_ENTITIES)) {
    text = text.split(k).join(HTML_ENTITIES[k])
  }
  return text.replace(/&#(\d+);/g, (_m, n) => String.fromCharCode(Number(n)))
}

/** Strip HTML to plain text, preserving paragraph breaks. */
function htmlToPlainText(html: string): string {
  let text = html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n\n")
    .replace(/<\/div>/gi, "\n")
    // Render attachment placeholders so they're visible in the output
    .replace(/<attachment[^>]*\/?>(?:<\/attachment>)?/gi, "[attachment]")
  text = text.replace(/<[^>]+>/g, "")
  text = decodeHtmlEntities(text)
  text = text
    .split("\n")
    .map((line) => line.replace(/[ \t]+/g, " ").trim())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim()
  return text
}

export interface ChatMarkdownOptions {
  /** Display name for the chat (1:1 person, group title, etc.). */
  title: string
  /** Raw Teams chat ID. */
  chatId: string
  /** "oneOnOne" | "group" | "meeting" -- rendered as a friendly label. */
  chatType: string
  /** ISO timestamp this download happened. */
  fetchedAt: string
  /** Lookback selection (e.g., "all", "since-last-download", "30"). */
  lookback: string
  /** Time floor used for this fetch, when one was applied. */
  sinceIso: string | null
}

const KIND_LABELS: Record<string, string> = {
  oneOnOne: "1:1",
  group: "Group",
  meeting: "Meeting",
}

/** Render a list of chat messages as a single markdown thread. */
export function renderChatMarkdown(
  messages: TeamsChatMessage[],
  options: ChatMarkdownOptions,
): string {
  // Newest-first listing is unhelpful for reading; emit oldest-first.
  const sorted = [...messages].sort((a, b) =>
    a.createdDateTime.localeCompare(b.createdDateTime),
  )

  const fetchedDate = new Date(options.fetchedAt)
  const fetchedShort = fetchedDate.toISOString().slice(0, 10)
  const kindLabel = KIND_LABELS[options.chatType] ?? options.chatType

  const out: string[] = []
  out.push(`# Chat: ${options.title || "Untitled chat"}`)
  out.push("")
  out.push(`Chat type: ${kindLabel}`)
  out.push(`Chat ID: ${options.chatId}`)
  out.push(`Downloaded: ${fetchedShort}`)
  const lookbackLabel = options.sinceIso
    ? `${options.lookback} (since ${options.sinceIso.slice(0, 10)})`
    : options.lookback
  out.push(`Lookback: ${lookbackLabel}`)
  out.push(`Messages: ${sorted.length}`)
  out.push("")
  out.push("---")
  out.push("")

  if (sorted.length === 0) {
    out.push("_(no messages in this window)_")
    return out.join("\n").trimEnd() + "\n"
  }

  let lastDate = ""
  for (const m of sorted) {
    const d = new Date(m.createdDateTime)
    const dateLabel = d.toISOString().slice(0, 10) // YYYY-MM-DD
    const timeLabel = d
      .toISOString()
      .slice(11, 16) // HH:MM
    if (dateLabel !== lastDate) {
      out.push(`## ${dateLabel}`)
      out.push("")
      lastDate = dateLabel
    }
    const sender = m.from?.user?.displayName?.trim() || "Unknown"
    const rawBody = m.body?.content ?? ""
    const isHtml = (m.body?.contentType ?? "text").toLowerCase() === "html"
    let bodyText = isHtml ? htmlToPlainText(rawBody) : decodeHtmlEntities(rawBody)
    const attachmentCount = Array.isArray(m.attachments) ? m.attachments.length : 0
    if (!bodyText && attachmentCount > 0) {
      bodyText = `_(${attachmentCount} attachment${attachmentCount === 1 ? "" : "s"})_`
    }
    bodyText = bodyText || "_(empty)_"
    out.push(`[${timeLabel}] **${sender}**`)
    // Indent body with > so it reads as a quoted block; preserves line breaks
    const quoted = bodyText
      .split("\n")
      .map((line) => (line.length > 0 ? `> ${line}` : ">"))
      .join("\n")
    out.push(quoted)
    if (attachmentCount > 0 && bodyText !== `_(${attachmentCount} attachment${attachmentCount === 1 ? "" : "s"})_`) {
      out.push(`> _(+${attachmentCount} attachment${attachmentCount === 1 ? "" : "s"})_`)
    }
    out.push("")
  }

  return out.join("\n").trimEnd() + "\n"
}
