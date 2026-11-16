// ─── 工具声明（由 docs 编辑器 fork 的 AGENT_TOOLS 机械转换而来） ────────
//
// 来源：packages/docx/src/editor/ai/tools.ts 的 AGENT_TOOLS（纯元数据，无执行体）。
// 本文件在**主进程**求值：不得 import 任何编辑器模块（tiptap / react / docx-engine）。
//
// 转换规则（scripts/gen-declarations.mjs）：
//   name/description 原样；inputSchema → parameters；kinds 固定 docx 三扩展名；
//   新鲜度令牌固定 docRev+epoch（位置型命令的坐标来自页面模型或块索引）。
//
// 读类工具（READ_TOOLS）由执行点 read/plan 提供——宿主只把 readOnly 工具交给
// 执行点，页面代理不承载读，所以这些必须能从字节侧独立回答。
// 其余（WRITE_TOOLS）是页面内工具，由编辑器的 executeTool 执行；无页面时
// 执行点回 command.page-unavailable（契约 §7.2，不回退第二份写模型）。
//
// 已移除的宿主型工具：web_search / image_search / analyze_media /
//   generate_image / create_document（归助手内置能力，不属于文档语义）。

import type { ToolDecl } from './types'

export const KINDS = ['docx', 'docm', 'dotx']

/** 只读工具：执行点侧实现（模型能拿到文本结果）。 */
export const READ_TOOLS: ToolDecl[] = [
  {
    name: "get_document_context",
    label: "Get Document Context",
    kinds: KINDS,
    readOnly: true,
    requiresBaseVersion: false,
    tokens: 'docRev+epoch',
    description: "Get the latest state of the current document: block list (index|type|content preview), full-text stats (word/character counts) and the current selection. Block indexes change after modifications; call this when you need up-to-date indexes.",
    parameters: {
  "type": "object",
  "properties": {},
  "required": []
},
  },
  {
    name: "read_blocks",
    label: "Read Blocks",
    kinds: KINDS,
    readOnly: true,
    requiresBaseVersion: false,
    tokens: 'docRev+epoch',
    description: "Read the full content of a block range (restricted HTML). Previews in the block list are truncated; you must read the full original text with this tool before rewriting. Long ranges are paged: a truncated result says which offset to continue from; concatenate the slices in order to get the full HTML.",
    parameters: {
  "type": "object",
  "properties": {
    "startBlockIndex": {
      "type": "integer",
      "description": "start block index (0-based, inclusive)"
    },
    "endBlockIndex": {
      "type": "integer",
      "description": "end block index (inclusive)"
    },
    "offset": {
      "type": "integer",
      "description": "character offset to continue a truncated read (default 0); use the offset given in the previous truncation notice"
    }
  },
  "required": [
    "startBlockIndex",
    "endBlockIndex"
  ]
},
  },
  {
    name: "read_revisions",
    label: "Read Revisions",
    kinds: KINDS,
    readOnly: true,
    requiresBaseVersion: false,
    tokens: 'docRev+epoch',
    description: "List every pending tracked revision (insertions, deletions, formatting/move/table changes) with its id, type, author, date, block index, the affected text and, for formatting changes, what changed. Ids are positional (r1 = first in document order) and renumber after every edit, so read again before accept_changes / reject_changes.",
    parameters: {
  "type": "object",
  "properties": {},
  "required": []
},
  },
  {
    name: "read_comments",
    label: "Read Comments",
    kinds: KINDS,
    readOnly: true,
    requiresBaseVersion: false,
    tokens: 'docRev+epoch',
    description: "List all comment threads (including resolved ones) with ids, authors, anchored block indexes and anchor text. Unresolved threads already ride along in the message context; use this for the full picture.",
    parameters: {
  "type": "object",
  "properties": {},
  "required": []
},
  },
  {
    name: "read_notes",
    label: "Read Notes",
    kinds: KINDS,
    readOnly: true,
    requiresBaseVersion: false,
    tokens: 'docRev+epoch',
    description: "List every footnote and endnote with its id, number, the block holding its reference mark and its text.",
    parameters: {
  "type": "object",
  "properties": {},
  "required": []
},
  },
  {
    name: "list_styles",
    label: "List Styles",
    kinds: KINDS,
    readOnly: true,
    requiresBaseVersion: false,
    tokens: 'docRev+epoch',
    description: "List the styles defined in the document (id, name, type, parent, heading level) for define_style / applyStyle.",
    parameters: {
  "type": "object",
  "properties": {},
  "required": []
},
  },
]

/** 写工具：页面内 executeTool 执行；无页面 → page-unavailable。 */
export const WRITE_TOOLS: ToolDecl[] = [
  {
    name: "insert_content",
    label: "Insert Content",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Insert new content at a given position (restricted HTML, may contain multiple blocks). For writing/continuing/generating new content; to rewrite existing content use replace_blocks.",
    parameters: {
  "type": "object",
  "properties": {
    "html": {
      "type": "string",
      "description": "restricted HTML fragment to insert; <span data-tab></span> is a tab character (literal tabs collapse to spaces)"
    },
    "afterBlockIndex": {
      "type": "integer",
      "description": "insert after this block index; -1 = start of document; omitted = after the block containing the cursor"
    }
  },
  "required": [
    "html"
  ]
},
  },
  {
    name: "write_document",
    label: "Write Document",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "[For long new content: a whole document, a chapter, a full report, article or translation] Hands the writing to the system writer, which streams restricted HTML straight into the document while the user watches; you never write the text yourself. Give a concrete plan (title, section outline with the key points of each, tone, target length) and put every fact, figure, name and quote the text must use into context — the writer sees only the plan and context, not the conversation. Omit afterBlockIndex on a blank document; on a document with content, pass afterBlockIndex to insert after that block, or replaceDocument=true when the user asked to rewrite everything. Short additions (a paragraph or two) use insert_content instead.",
    parameters: {
  "type": "object",
  "properties": {
    "plan": {
      "type": "string",
      "description": "title, section outline with key points per section, tone, target length"
    },
    "title": {
      "type": "string",
      "description": "document/chapter title, when there is one"
    },
    "context": {
      "type": "string",
      "description": "reference material: facts, figures, quotes, sources gathered from the conversation, attachments and web_search"
    },
    "afterBlockIndex": {
      "type": "integer",
      "description": "insert after this block index (-1 = document start); omitted = whole document"
    },
    "replaceDocument": {
      "type": "boolean",
      "description": "true replaces all existing content (only when the user asked for a rewrite)"
    }
  },
  "required": [
    "plan"
  ]
},
  },
  {
    name: "replace_blocks",
    label: "Replace Blocks",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Replace a block range with new content (restricted HTML). For rewriting/translating/condensing/expanding existing content; the new block count may differ from the old. New blocks inherit the replaced blocks' paragraph and text formatting (font, size, color, indent, spacing, alignment) automatically, and a rewritten <table> keeps the old table's column widths, borders, shading and cell formatting (unchanged cells keep their content); never try to restore formatting afterwards.",
    parameters: {
  "type": "object",
  "properties": {
    "startBlockIndex": {
      "type": "integer",
      "description": "start block index (0-based, inclusive)"
    },
    "endBlockIndex": {
      "type": "integer",
      "description": "end block index (inclusive)"
    },
    "html": {
      "type": "string",
      "description": "replacement restricted HTML fragment; <span data-tab></span> is a tab character"
    }
  },
  "required": [
    "startBlockIndex",
    "endBlockIndex",
    "html"
  ]
},
  },
  {
    name: "replace_selection",
    label: "Replace Selection",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Replace exactly the user's selected text (the <sel>…</sel> span in the context) with new inline content, leaving the rest of the block untouched. For rewording/translating/correcting a selected phrase or sentence inside a paragraph. The new text inherits the selection's formatting unless the fragment styles it. Requires a range selection inside one paragraph/heading/list item; for whole blocks or several blocks use replace_blocks.",
    parameters: {
  "type": "object",
  "properties": {
    "html": {
      "type": "string",
      "description": "replacement inline content: plain text or restricted inline HTML (strong em u s a br formula)"
    }
  },
  "required": [
    "html"
  ]
},
  },
  {
    name: "apply_ops",
    label: "Apply Ops",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Run a list of formatting/structure ops as one atomic transaction (see the apply_ops guide in the system prompt): setFont, setMatchedFont, setParagraphFormat, setHeadingLevel, findReplace, deleteBlocks, moveBlocks, setList, clearList, setImageProperties, insertToc, insertField, insertBookmark, updateFields, insertTableRow, deleteTableRow, insertTableColumn, deleteTableColumn, mergeTableCells, splitTableCell, setTableCellFormat, setTableStyle, applyStyle. Each op is a flat { op, target?, ...fields } object; fields are patches (present = set, null = clear, absent = untouched). Any invalid op rejects the whole batch with its usage line.",
    parameters: {
  "type": "object",
  "properties": {
    "ops": {
      "type": "array",
      "description": "ops executed in order, e.g. [{\"op\":\"setFont\",\"target\":{\"nodeType\":\"docHeading\"},\"color\":\"#FF0000\"}]",
      "items": {
        "type": "object"
      }
    },
    "dryRun": {
      "type": "boolean",
      "description": "validate and return the plan without changing the document"
    }
  },
  "required": [
    "ops"
  ]
},
  },
  {
    name: "accept_changes",
    label: "Accept Changes",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Accept pending tracked changes: insertions become plain text, deleted text disappears, new formatting stays. Select with all: true, ids from read_revisions, or any combination of author / type / blockIndex / blockRange / before. Only when the user asks to accept changes.",
    parameters: {
  "type": "object",
  "properties": {
    "all": {
      "type": "boolean",
      "description": "true = every pending change"
    },
    "ids": {
      "type": "array",
      "items": {
        "type": "string"
      },
      "description": "revision ids from read_revisions (r1, r2, …); positional, so re-read after any edit"
    },
    "author": {
      "type": "string",
      "description": "only changes by this author (case-insensitive)"
    },
    "type": {
      "type": "string",
      "enum": [
        "insertion",
        "deletion",
        "formatting",
        "move"
      ],
      "description": "only changes of this type"
    },
    "blockIndex": {
      "type": "integer",
      "description": "only changes inside this block"
    },
    "blockRange": {
      "type": "array",
      "items": {
        "type": "integer"
      },
      "minItems": 2,
      "maxItems": 2,
      "description": "[from, to] block indexes, inclusive"
    },
    "before": {
      "type": "string",
      "description": "ISO date; only changes recorded before it"
    }
  },
  "required": []
},
  },
  {
    name: "reject_changes",
    label: "Reject Changes",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Reject pending tracked changes: inserted text disappears, deleted text is restored, formatting goes back to what it was. Same selector as accept_changes. Only when the user asks to reject changes.",
    parameters: {
  "type": "object",
  "properties": {
    "all": {
      "type": "boolean",
      "description": "true = every pending change"
    },
    "ids": {
      "type": "array",
      "items": {
        "type": "string"
      },
      "description": "revision ids from read_revisions (r1, r2, …); positional, so re-read after any edit"
    },
    "author": {
      "type": "string",
      "description": "only changes by this author (case-insensitive)"
    },
    "type": {
      "type": "string",
      "enum": [
        "insertion",
        "deletion",
        "formatting",
        "move"
      ],
      "description": "only changes of this type"
    },
    "blockIndex": {
      "type": "integer",
      "description": "only changes inside this block"
    },
    "blockRange": {
      "type": "array",
      "items": {
        "type": "integer"
      },
      "minItems": 2,
      "maxItems": 2,
      "description": "[from, to] block indexes, inclusive"
    },
    "before": {
      "type": "string",
      "description": "ISO date; only changes recorded before it"
    }
  },
  "required": []
},
  },
  {
    name: "reply_comment",
    label: "Reply Comment",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Add a reply to a comment thread: after completing a requested change (summarize what changed), or to answer/ask back when the comment is a question or is ambiguous.",
    parameters: {
  "type": "object",
  "properties": {
    "parentId": {
      "type": "string",
      "description": "id of the comment thread to reply to"
    },
    "text": {
      "type": "string",
      "description": "reply text"
    }
  },
  "required": [
    "parentId",
    "text"
  ]
},
  },
  {
    name: "resolve_comment",
    label: "Resolve Comment",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Mark a comment thread as resolved. Only after the requested change was applied (reply first), or when the user explicitly asked to resolve.",
    parameters: {
  "type": "object",
  "properties": {
    "id": {
      "type": "string",
      "description": "id of the comment thread to resolve"
    }
  },
  "required": [
    "id"
  ]
},
  },
  {
    name: "insert_footnote",
    label: "Insert Footnote",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Add a footnote: a superscript reference mark goes into the block (right after afterText, else at its end) and the note text prints at the bottom of that page. Use for sources, asides and clarifications the user asks to footnote.",
    parameters: {
  "type": "object",
  "properties": {
    "blockIndex": {
      "type": "integer",
      "description": "block that gets the reference mark"
    },
    "afterText": {
      "type": "string",
      "description": "exact text in that block the mark follows; omitted = end of the block"
    },
    "text": {
      "type": "string",
      "description": "the note text (\\n separates paragraphs)"
    }
  },
  "required": [
    "blockIndex",
    "text"
  ]
},
  },
  {
    name: "insert_endnote",
    label: "Insert Endnote",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Add an endnote (same as insert_footnote, but the text collects at the end of the document). Use when the document already uses endnotes or the user asks for them.",
    parameters: {
  "type": "object",
  "properties": {
    "blockIndex": {
      "type": "integer",
      "description": "block that gets the reference mark"
    },
    "afterText": {
      "type": "string",
      "description": "exact text in that block the mark follows; omitted = end of the block"
    },
    "text": {
      "type": "string",
      "description": "the note text (\\n separates paragraphs)"
    }
  },
  "required": [
    "blockIndex",
    "text"
  ]
},
  },
  {
    name: "delete_note",
    label: "Delete Note",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Remove a footnote or endnote together with its reference mark; ids from read_notes.",
    parameters: {
  "type": "object",
  "properties": {
    "kind": {
      "type": "string",
      "enum": [
        "footnote",
        "endnote"
      ]
    },
    "id": {
      "type": "string",
      "description": "note id from read_notes"
    }
  },
  "required": [
    "kind",
    "id"
  ]
},
  },
  {
    name: "edit_note",
    label: "Edit Note",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Change the text of an existing footnote or endnote in place (findReplace inside the note): the note keeps its id, reference mark and formatting. Use for a typo or a citation fix instead of delete_note + insert_footnote.",
    parameters: {
  "type": "object",
  "properties": {
    "kind": {
      "type": "string",
      "enum": [
        "footnote",
        "endnote"
      ],
      "description": "omit when the id is unambiguous across footnotes and endnotes"
    },
    "id": {
      "type": "string",
      "description": "note id from read_notes"
    },
    "find": {
      "type": "string",
      "description": "exact text inside the note to replace"
    },
    "replace": {
      "type": "string",
      "description": "replacement text (\"\" deletes the match)"
    },
    "matchCase": {
      "type": "boolean",
      "description": "defaults to true"
    }
  },
  "required": [
    "id",
    "find",
    "replace"
  ]
},
  },
  {
    name: "add_comment",
    label: "Add Comment",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Start a new comment thread on a block or on an exact text span inside it, without changing the document: review notes, questions, suggestions the user should decide on.",
    parameters: {
  "type": "object",
  "properties": {
    "blockIndex": {
      "type": "integer",
      "description": "block to annotate"
    },
    "text": {
      "type": "string",
      "description": "exact text inside the block to anchor the comment to; omitted = the whole block"
    },
    "occurrence": {
      "type": "integer",
      "description": "which match to anchor when text occurs more than once (1 = first)"
    },
    "comment": {
      "type": "string",
      "description": "comment body; \\n starts a new paragraph"
    },
    "author": {
      "type": "string",
      "description": "author name; default: the AI author"
    },
    "initials": {
      "type": "string",
      "description": "author initials shown in the margin"
    }
  },
  "required": [
    "blockIndex",
    "comment"
  ]
},
  },
  {
    name: "delete_comment",
    label: "Delete Comment",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Delete a comment. A thread root with replies is refused unless withReplies is true; a reply id deletes just that reply.",
    parameters: {
  "type": "object",
  "properties": {
    "id": {
      "type": "string",
      "description": "id of the comment to delete"
    },
    "withReplies": {
      "type": "boolean",
      "description": "also delete the replies of a thread root (default false)"
    }
  },
  "required": [
    "id"
  ]
},
  },
  {
    name: "insert_image",
    label: "Insert Image",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Download a direct image link (an imageUrl from image_search) and insert it into the document (at the cursor / end of document).",
    parameters: {
  "type": "object",
  "properties": {
    "url": {
      "type": "string",
      "description": "direct image link"
    },
    "maxWidthPx": {
      "type": "integer",
      "description": "maximum width (px), default 480"
    }
  },
  "required": [
    "url"
  ]
},
  },
  {
    name: "insert_chart",
    label: "Insert Chart",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Insert a chart (saved as a native Word chart). Data must be real: from the document content or web_search results — do not make up numbers.",
    parameters: {
  "type": "object",
  "properties": {
    "kind": {
      "type": "string",
      "enum": [
        "bar",
        "line",
        "pie"
      ],
      "description": "chart type"
    },
    "title": {
      "type": "string",
      "description": "chart title"
    },
    "categories": {
      "type": "array",
      "items": {
        "type": "string"
      },
      "description": "category (x axis / sector) labels"
    },
    "series": {
      "type": "array",
      "description": "data series; values has the same length as categories, use null for missing data. Pie charts use only the first series",
      "items": {
        "type": "object",
        "properties": {
          "name": {
            "type": "string"
          },
          "values": {
            "type": "array",
            "items": {
              "type": [
                "number",
                "null"
              ]
            }
          }
        },
        "required": [
          "values"
        ]
      }
    },
    "afterBlockIndex": {
      "type": "integer",
      "description": "insert after this block index; -1 = start of document; omitted = after the block containing the cursor"
    }
  },
  "required": [
    "kind",
    "categories",
    "series"
  ]
},
  },
  {
    name: "edit_chart",
    label: "Edit Chart",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Edit the data of an existing chart in the document (title/category labels/series names/values). Chart blocks in the block list can be edited; the category count and the number of values per series must match the original chart (data points cannot be added or removed).",
    parameters: {
  "type": "object",
  "properties": {
    "blockIndex": {
      "type": "integer",
      "description": "block index of the chart"
    },
    "title": {
      "type": "string",
      "description": "new title (omit to keep)"
    },
    "categories": {
      "type": "array",
      "items": {
        "type": [
          "string",
          "null"
        ]
      },
      "description": "new category labels, same length as the original; pass null for positions to keep"
    },
    "series": {
      "type": "array",
      "description": "series to change",
      "items": {
        "type": "object",
        "properties": {
          "index": {
            "type": "integer",
            "description": "series index (0-based)"
          },
          "name": {
            "type": "string",
            "description": "new series name (omit to keep)"
          },
          "values": {
            "type": "array",
            "items": {
              "type": [
                "number",
                "null"
              ]
            },
            "description": "new values, same length as the original series; pass null for positions to keep"
          }
        },
        "required": [
          "index"
        ]
      }
    }
  },
  "required": [
    "blockIndex"
  ]
},
  },
  {
    name: "set_header_footer",
    label: "Set Header Footer",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Set the page header or footer text (the current contents are listed in the message context). Plain text; \\n separates lines; the tokens {PAGE} and {NUMPAGES} become live page-number fields; an empty string clears the text. Per-line alignment/styling of the existing header/footer is preserved; images in it are untouched. view \"first\"/\"even\" writes the different-first-page / even-page variant (enabling that setting if needed).",
    parameters: {
  "type": "object",
  "properties": {
    "kind": {
      "type": "string",
      "enum": [
        "header",
        "footer"
      ]
    },
    "text": {
      "type": "string",
      "description": "new text; \\n between lines; may contain {PAGE} / {NUMPAGES}; \"\" clears"
    },
    "view": {
      "type": "string",
      "enum": [
        "default",
        "first",
        "even"
      ],
      "description": "which variant to write (default when omitted)"
    }
  },
  "required": [
    "kind",
    "text"
  ]
},
  },
  {
    name: "set_page_setup",
    label: "Set Page Setup",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Change paper size, orientation, margins, text columns, the different-first-page flag or page numbering of the document or of one section (the message context lists the sections). Only the fields given change. Lengths take a unit suffix (cm, mm, in, pt) or twips. Setting the orientation alone swaps the current width and height.",
    parameters: {
  "type": "object",
  "properties": {
    "section": {
      "type": "integer",
      "description": "section index (0-based) to change; omit with blockIndex omitted = every section"
    },
    "blockIndex": {
      "type": "integer",
      "description": "change the section that contains this block index"
    },
    "paper": {
      "type": "string",
      "description": "named size: A3, A4, A5, B5, Letter, Legal, Tabloid"
    },
    "width": {
      "type": [
        "string",
        "number"
      ],
      "description": "custom page width, a number of twips or a string with a unit: \"2.54cm\", \"1in\", \"72pt\", \"10mm\""
    },
    "height": {
      "type": [
        "string",
        "number"
      ],
      "description": "custom page height, a number of twips or a string with a unit: \"2.54cm\", \"1in\", \"72pt\", \"10mm\""
    },
    "orientation": {
      "type": "string",
      "enum": [
        "portrait",
        "landscape"
      ]
    },
    "margins": {
      "type": "object",
      "description": "any of top, right, bottom, left, header, footer, gutter; each a number of twips or a string with a unit: \"2.54cm\", \"1in\", \"72pt\", \"10mm\"",
      "properties": {
        "top": {
          "type": [
            "string",
            "number"
          ]
        },
        "right": {
          "type": [
            "string",
            "number"
          ]
        },
        "bottom": {
          "type": [
            "string",
            "number"
          ]
        },
        "left": {
          "type": [
            "string",
            "number"
          ]
        },
        "header": {
          "type": [
            "string",
            "number"
          ],
          "description": "header distance from the page top"
        },
        "footer": {
          "type": [
            "string",
            "number"
          ],
          "description": "footer distance from the page bottom"
        },
        "gutter": {
          "type": [
            "string",
            "number"
          ],
          "description": "extra binding margin"
        }
      }
    },
    "columns": {
      "type": "object",
      "description": "text columns of the section",
      "properties": {
        "count": {
          "type": "integer",
          "description": "1-12"
        },
        "spacing": {
          "type": [
            "string",
            "number"
          ],
          "description": "gap between columns, a number of twips or a string with a unit: \"2.54cm\", \"1in\", \"72pt\", \"10mm\""
        }
      },
      "required": [
        "count"
      ]
    },
    "titlePg": {
      "type": "boolean",
      "description": "different first page (own header/footer on the first page of the section)"
    },
    "pageNumberStart": {
      "type": [
        "integer",
        "null"
      ],
      "description": "restart page numbering at this value; null = continue from the previous section"
    },
    "pageNumberFormat": {
      "type": "string",
      "description": "decimal, lowerRoman, upperRoman, lowerLetter, upperLetter, ..."
    }
  }
},
  },
  {
    name: "insert_section_break",
    label: "Insert Section Break",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Insert a section break after a block so the blocks after it form a new section with its own page setup, headers and footers (the new section starts as a copy of the current one; change it with set_page_setup afterwards). type nextPage starts the new section on a new page, continuous keeps the flow (needed for column changes mid-page), evenPage / oddPage start on the next even / odd page.",
    parameters: {
  "type": "object",
  "properties": {
    "afterBlockIndex": {
      "type": "integer",
      "description": "the break goes after this block index; -1 = before the first block"
    },
    "type": {
      "type": "string",
      "enum": [
        "nextPage",
        "continuous",
        "evenPage",
        "oddPage"
      ],
      "description": "default nextPage"
    }
  },
  "required": [
    "afterBlockIndex"
  ]
},
  },
  {
    name: "define_style",
    label: "Define Style",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Create a style or change the definition of an existing one (paragraph, character or table style in styles.xml). Only the fields given change; everything else the style already defines is kept. Paragraphs using the style pick the change up automatically. Use apply_style (or the applyStyle op) to put a style on paragraphs. Existing style ids are listed by list_styles / in the tool output.",
    parameters: {
  "type": "object",
  "properties": {
    "styleId": {
      "type": "string",
      "description": "style id (letters, digits, - and _; e.g. \"Heading1\", \"Quote\", \"MyCallout\")"
    },
    "name": {
      "type": "string",
      "description": "display name (defaults to the id when creating)"
    },
    "type": {
      "type": "string",
      "enum": [
        "paragraph",
        "character",
        "table"
      ],
      "description": "when creating (default paragraph); an existing style keeps its type"
    },
    "basedOn": {
      "type": [
        "string",
        "null"
      ],
      "description": "parent style id to inherit from; null removes the inheritance"
    },
    "next": {
      "type": [
        "string",
        "null"
      ],
      "description": "style of the paragraph created by Enter at the end of a paragraph in this style"
    },
    "quickFormat": {
      "type": "boolean",
      "description": "show in the quick style gallery"
    },
    "paragraph": {
      "type": "object",
      "description": "paragraph formatting (paragraph and table styles)",
      "properties": {
        "align": {
          "type": [
            "string",
            "null"
          ],
          "enum": [
            "left",
            "center",
            "right",
            "justify",
            null
          ]
        },
        "spaceBefore": {
          "type": [
            "number",
            "string",
            "null"
          ],
          "description": "space before, a length like \"0.5in\", \"1.27cm\", \"12pt\" (bare numbers are points), or null to clear"
        },
        "spaceAfter": {
          "type": [
            "number",
            "string",
            "null"
          ],
          "description": "space after, a length like \"0.5in\", \"1.27cm\", \"12pt\" (bare numbers are points), or null to clear"
        },
        "lineSpacing": {
          "type": [
            "number",
            "null"
          ],
          "description": "line spacing as a multiple of single (1, 1.15, 1.5, 2)"
        },
        "indentLeft": {
          "type": [
            "number",
            "string",
            "null"
          ],
          "description": "a length like \"0.5in\", \"1.27cm\", \"12pt\" (bare numbers are points), or null to clear"
        },
        "indentRight": {
          "type": [
            "number",
            "string",
            "null"
          ],
          "description": "a length like \"0.5in\", \"1.27cm\", \"12pt\" (bare numbers are points), or null to clear"
        },
        "firstLineIndent": {
          "type": [
            "number",
            "string",
            "null"
          ],
          "description": "first-line indent, negative = hanging indent, a length like \"0.5in\", \"1.27cm\", \"12pt\" (bare numbers are points), or null to clear"
        },
        "keepNext": {
          "type": "boolean",
          "description": "keep with next paragraph"
        },
        "keepLines": {
          "type": "boolean",
          "description": "keep lines together"
        },
        "pageBreakBefore": {
          "type": "boolean"
        },
        "outlineLevel": {
          "type": [
            "integer",
            "null"
          ],
          "description": "1-9: the style becomes a heading of that level in the navigation pane / TOC"
        },
        "tabs": {
          "type": [
            "array",
            "null"
          ],
          "description": "tab stops of the style (replaces its w:tabs; [] or null removes them): [{ pos: twips or \"6.5in\", align?: left|center|right|decimal|bar, leader?: none|dot|hyphen|underscore|middleDot, clear?: true }]",
          "items": {
            "type": "object"
          }
        }
      }
    },
    "run": {
      "type": "object",
      "description": "text formatting (all style types)",
      "properties": {
        "bold": {
          "type": "boolean"
        },
        "italic": {
          "type": "boolean"
        },
        "underline": {
          "type": "boolean"
        },
        "strike": {
          "type": "boolean"
        },
        "caps": {
          "type": "boolean"
        },
        "smallCaps": {
          "type": "boolean"
        },
        "color": {
          "type": [
            "string",
            "null"
          ],
          "description": "\"#RRGGBB\" or null to clear"
        },
        "highlight": {
          "type": [
            "string",
            "null"
          ],
          "description": "highlight color name (yellow, green, cyan, magenta, blue, red, ...) or null"
        },
        "fontSize": {
          "type": [
            "number",
            "null"
          ],
          "description": "points"
        },
        "fontFamily": {
          "type": [
            "string",
            "null"
          ],
          "description": "Latin font face"
        },
        "eastAsiaFontFamily": {
          "type": [
            "string",
            "null"
          ],
          "description": "CJK font face"
        }
      }
    }
  },
  "required": [
    "styleId"
  ]
},
  },
  {
    name: "set_watermark",
    label: "Set Watermark",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Set, restyle or remove the page watermark: a gray text like DRAFT or CONFIDENTIAL, or a washed-out picture (a logo), drawn behind the body on every page (written into the page header, where Word shows it in Design > Watermark). Give text or image, not both; text null or image null removes whichever watermark the document has.",
    parameters: {
  "type": "object",
  "properties": {
    "text": {
      "type": [
        "string",
        "null"
      ],
      "description": "watermark text; null removes the watermark"
    },
    "fontFamily": {
      "type": "string",
      "description": "font face (default the document theme face)"
    },
    "color": {
      "type": "string",
      "description": "\"#RRGGBB\" (default silver)"
    },
    "opacity": {
      "type": "number",
      "description": "0-1 (default 0.5)"
    },
    "diagonal": {
      "type": "boolean",
      "description": "diagonal (default) or horizontal layout"
    },
    "bold": {
      "type": "boolean"
    },
    "italic": {
      "type": "boolean"
    },
    "image": {
      "type": [
        "string",
        "null"
      ],
      "description": "picture watermark: http(s) or data: URL of a png/jpg/gif, centered behind the body; null removes the watermark"
    },
    "scale": {
      "type": "number",
      "description": "picture size as a percent of its natural size (default: fit the page inside the margins)"
    },
    "washout": {
      "type": "boolean",
      "description": "fade the picture like Word's washout (default true)"
    }
  }
},
  },
  {
    name: "insert_text_box",
    label: "Insert Text Box",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Insert a floating text box (a positioned rectangle holding text, like a callout or sidebar) anchored to a block. Body text either flows above/below it or ignores it.",
    parameters: {
  "type": "object",
  "properties": {
    "afterBlockIndex": {
      "type": "integer",
      "description": "anchor after this block (-1 = start of document); default: end of document"
    },
    "text": {
      "type": "string",
      "description": "box text; \\n separates paragraphs"
    },
    "width": {
      "type": [
        "number",
        "string"
      ],
      "description": "box width, \"2.54cm\", \"1in\", \"72pt\", \"96px\" or a bare number of points"
    },
    "height": {
      "type": [
        "number",
        "string"
      ],
      "description": "box height, \"2.54cm\", \"1in\", \"72pt\", \"96px\" or a bare number of points"
    },
    "x": {
      "type": [
        "number",
        "string"
      ],
      "description": "horizontal offset, \"2.54cm\", \"1in\", \"72pt\", \"96px\" or a bare number of points"
    },
    "y": {
      "type": [
        "number",
        "string"
      ],
      "description": "vertical offset, \"2.54cm\", \"1in\", \"72pt\", \"96px\" or a bare number of points"
    },
    "anchor": {
      "type": "string",
      "enum": [
        "paragraph",
        "page"
      ],
      "description": "offsets measure from the anchor paragraph (default) or the page corner"
    },
    "wrap": {
      "type": "string",
      "enum": [
        "topAndBottom",
        "none"
      ],
      "description": "body text resumes below the box (default for paragraph anchors) or ignores it (default for page anchors)"
    },
    "fill": {
      "type": [
        "string",
        "null"
      ],
      "description": "\"#RRGGBB\" background (default white); null = transparent"
    },
    "borderColor": {
      "type": [
        "string",
        "null"
      ],
      "description": "\"#RRGGBB\" outline (default black); null = no outline"
    },
    "fontSize": {
      "type": [
        "number",
        "string"
      ],
      "description": "text size in points (a bare number, or a length like \"12pt\")"
    },
    "align": {
      "type": "string",
      "enum": [
        "left",
        "center",
        "right",
        "justify"
      ]
    },
    "bold": {
      "type": "boolean"
    },
    "color": {
      "type": "string",
      "description": "\"#RRGGBB\" text color"
    }
  },
  "required": [
    "text",
    "width",
    "height",
    "x",
    "y"
  ]
},
  },
  {
    name: "insert_picture",
    label: "Insert Picture",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Insert a picture from a direct image link at a block position, at a given size, inline or floating (positioned on the page, with text wrapping around/behind/in front). insert_image is the simpler cursor-position variant.",
    parameters: {
  "type": "object",
  "properties": {
    "url": {
      "type": "string",
      "description": "direct image link (png/jpg/gif)"
    },
    "afterBlockIndex": {
      "type": "integer",
      "description": "insert after this block (-1 = start of document); default: end of document"
    },
    "width": {
      "type": [
        "number",
        "string"
      ],
      "description": "display width, \"2.54cm\", \"1in\", \"72pt\", \"96px\" or a bare number of points; default natural size capped at 480px"
    },
    "height": {
      "type": [
        "number",
        "string"
      ],
      "description": "display height; omit to keep the aspect ratio"
    },
    "float": {
      "type": "object",
      "description": "omit for an inline picture in its own paragraph; give it to float the picture at a position",
      "properties": {
        "anchor": {
          "type": "string",
          "enum": [
            "paragraph",
            "page",
            "margin"
          ],
          "description": "what x/y measure from: the anchor paragraph (x from the text column, y from the paragraph top), the page corner, or the margin box corner"
        },
        "x": {
          "type": [
            "number",
            "string"
          ],
          "description": "horizontal offset, \"2.54cm\", \"1in\", \"72pt\", \"96px\" or a bare number of points"
        },
        "y": {
          "type": [
            "number",
            "string"
          ],
          "description": "vertical offset, \"2.54cm\", \"1in\", \"72pt\", \"96px\" or a bare number of points"
        },
        "wrap": {
          "type": "string",
          "enum": [
            "square",
            "tight",
            "topAndBottom",
            "behind",
            "inFront"
          ],
          "description": "how body text flows around it (default square)"
        }
      },
      "required": [
        "x",
        "y"
      ]
    },
    "altText": {
      "type": "string",
      "description": "alternative text (accessibility)"
    }
  },
  "required": [
    "url"
  ]
},
  },
  {
    name: "generate_mindmap",
    label: "Generate Mindmap",
    kinds: KINDS,
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    description: "Render an interactive mind map of the document (or of one focused aspect) as a diagram card in the chat. Call it when the user asks for a mind map / knowledge structure / overview diagram of the document or a topic. Pass the entire tree in this single call: the root is the document title or the focus topic, its children are the main sections, and so on. Labels must be short phrases in the document's own language and wording — never invent topics the document does not cover. Cite the block index each node summarizes in `block` so the user can jump to the passage. This tool does not modify the document; after it returns, reply with one short sentence.",
    parameters: {
  "type": "object",
  "properties": {
    "focus": {
      "type": "string",
      "description": "optional: the aspect the map should center on (e.g. \"risk controls\", \"for beginners\"); omit for a whole-document map"
    },
    "root": {
      "type": "object",
      "description": "the mind map tree: root node with nested children",
      "properties": {
        "label": {
          "type": "string",
          "description": "short node label (2-8 words) in the document's language, taken from its headings or key phrases"
        },
        "block": {
          "type": "integer",
          "description": "0-based block index this node summarizes (from the current block list); omit when the node spans several blocks"
        },
        "children": {
          "type": "array",
          "description": "sub-topics; 2-5 per node reads best",
          "items": {
            "type": "object",
            "properties": {
              "label": {
                "type": "string",
                "description": "short node label (2-8 words) in the document's language, taken from its headings or key phrases"
              },
              "block": {
                "type": "integer",
                "description": "0-based block index this node summarizes (from the current block list); omit when the node spans several blocks"
              },
              "children": {
                "type": "array",
                "description": "sub-topics; 2-5 per node reads best",
                "items": {
                  "type": "object",
                  "properties": {
                    "label": {
                      "type": "string",
                      "description": "short node label (2-8 words) in the document's language, taken from its headings or key phrases"
                    },
                    "block": {
                      "type": "integer",
                      "description": "0-based block index this node summarizes (from the current block list); omit when the node spans several blocks"
                    },
                    "children": {
                      "type": "array",
                      "description": "sub-topics; 2-5 per node reads best",
                      "items": {
                        "type": "object",
                        "properties": {
                          "label": {
                            "type": "string",
                            "description": "short node label (2-8 words) in the document's language, taken from its headings or key phrases"
                          },
                          "block": {
                            "type": "integer",
                            "description": "0-based block index this node summarizes (from the current block list); omit when the node spans several blocks"
                          },
                          "children": {
                            "type": "array",
                            "description": "sub-topics; 2-5 per node reads best",
                            "items": {
                              "type": "object",
                              "properties": {
                                "label": {
                                  "type": "string",
                                  "description": "short node label (2-8 words) in the document's language, taken from its headings or key phrases"
                                },
                                "block": {
                                  "type": "integer",
                                  "description": "0-based block index this node summarizes (from the current block list); omit when the node spans several blocks"
                                },
                                "children": {
                                  "type": "array",
                                  "description": "sub-topics; 2-5 per node reads best",
                                  "items": {
                                    "type": "object",
                                    "properties": {
                                      "label": {
                                        "type": "string",
                                        "description": "short node label (2-8 words) in the document's language, taken from its headings or key phrases"
                                      },
                                      "block": {
                                        "type": "integer",
                                        "description": "0-based block index this node summarizes (from the current block list); omit when the node spans several blocks"
                                      },
                                      "children": {
                                        "type": "array",
                                        "description": "sub-topics; 2-5 per node reads best",
                                        "items": {
                                          "type": "object",
                                          "properties": {
                                            "label": {
                                              "type": "string",
                                              "description": "short node label (2-8 words) in the document's language"
                                            }
                                          },
                                          "required": [
                                            "label"
                                          ]
                                        }
                                      }
                                    },
                                    "required": [
                                      "label"
                                    ]
                                  }
                                }
                              },
                              "required": [
                                "label"
                              ]
                            }
                          }
                        },
                        "required": [
                          "label"
                        ]
                      }
                    }
                  },
                  "required": [
                    "label"
                  ]
                }
              }
            },
            "required": [
              "label"
            ]
          }
        }
      },
      "required": [
        "label"
      ]
    }
  },
  "required": [
    "root"
  ]
},
  },
]

export const GENERATED_TOOLS: ToolDecl[] = [...READ_TOOLS, ...WRITE_TOOLS]
