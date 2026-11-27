// Stable editor-side block identity.
//
// `docxIndex` is a file coordinate: it changes after every save and must never
// be used as undo identity. Top-level blocks carry `blockId` instead; the page
// keeps a separate blockId -> docxIndex map for the save planner.

let sequence = 0

export function newBlockId(): string {
  const cryptoApi = globalThis.crypto
  if (typeof cryptoApi?.randomUUID === 'function')
    return `b-${cryptoApi.randomUUID()}`
  sequence += 1
  return `b-${Date.now().toString(36)}-${sequence.toString(36)}`
}
