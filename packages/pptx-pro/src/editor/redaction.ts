/**
 * Label helpers shared by the manual redaction UI.
 *
 * The pptx redaction feature keeps the reader's real text/media in the file and
 * only writes a label marker; these helpers clean that label. They are kept out
 * of the removed AI module so the editor no longer depends on agent-core.
 */
const LABEL_STRIP = /[{}<>="']/g
const MAX_LABEL_LENGTH = 40

/** Clean a label for storage and display. */
export function sanitizeLabel(raw: string): string {
  return raw
    .replace(LABEL_STRIP, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, MAX_LABEL_LENGTH)
    .trim()
}

/** The label of a withheld picture/video/audio element, or null. */
export function mediaLabelOf(node: { redact?: string }): string | null {
  return typeof node.redact === 'string' && node.redact ? node.redact : null
}
