// ─── docx_apply 的 action 命令（由编辑器 AGENT_TOOLS 生成） ──────────
//
// 这些是「不是 op、也不是块命令」的专用动作：参数是 HTML 片段 / 图表数据 /
// 修订范围，不能塞进 apply_ops 的扁平 op 结构。
//
// 只收录在 docx 页面里**真的可用**的动作——它们的执行体不依赖 App 私有
// 状态（批注 / 脚注 / 页眉页脚 / 页面设置 / 样式 / 水印 / 文档生成的 accessor
// 还没接进页面桥）。未接线的动作列在 UNAVAILABLE_ACTIONS，不进目录、不可调用。

export interface ActionCommand {
  name: string
  description: string
  parameters: unknown
}

export const ACTION_COMMANDS: ActionCommand[] = [
  {
    name: "insert_content",
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
    name: "replace_blocks",
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
    name: "accept_changes",
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
    name: "insert_image",
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
    name: "insert_text_box",
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
]

export const ACTION_NAMES = new Set(ACTION_COMMANDS.map((action) => action.name))

/** 未接线：需要 App 会话内的 accessor（批注 / 脚注 / 页眉页脚 / 页面设置 / 样式 / 水印 / 文档生成）。 */
export const UNAVAILABLE_ACTIONS = ["write_document","reply_comment","resolve_comment","insert_footnote","insert_endnote","delete_note","edit_note","add_comment","delete_comment","set_header_footer","set_page_setup","insert_section_break","define_style","set_watermark","generate_mindmap"] as const
