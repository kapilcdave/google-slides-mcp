#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import {
  PREDEFINED_LAYOUTS,
  addImage,
  addShape,
  addSlide,
  addTable,
  addTextBox,
  batchUpdate,
  createDeckFromOutline,
  createPresentation,
  deleteElement,
  deleteSlide,
  duplicateSlide,
  exportDeck,
  formatText,
  getPresentation,
  getPresentationRaw,
  getSlide,
  listSlides,
  moveSlide,
  replaceAllText,
  reorderSlides,
  setSlideBackground,
  setSpeakerNotes,
  updateText,
} from "./presentations.js";
import { run } from "./util.js";

const server = new McpServer(
  {
    name: "google-slides-mcp",
    version: "0.1.0",
  },
  {
    instructions: [
      "You are connected to Google Slides.",
      "Typical flow: create_deck_from_outline for a new deck, or create_presentation + add_slide.",
      "Use get_presentation / get_slide to discover object IDs, then edit text, images, shapes, tables and speaker notes.",
      "Coordinates are in inches on a 13.333 x 7.5 in widescreen slide unless the deck uses another size.",
      "Always return the presentation URL to the user when done.",
    ].join("\n"),
  },
);

const presentationId = z
  .string()
  .min(1)
  .describe(
    "Presentation ID. Found in the Google Slides URL: docs.google.com/presentation/d/<ID>/edit",
  );

const slideId = z
  .string()
  .min(1)
  .describe("Slide (page) object ID from get_presentation or list_slides");

const objectId = z
  .string()
  .min(1)
  .describe("Page element object ID from get_slide");

const theme = z
  .enum(["midnight", "light", "sunset", "forest"])
  .optional()
  .describe("Color theme for generated/appended slides");

const layout = z
  .enum(PREDEFINED_LAYOUTS)
  .optional()
  .describe("Predefined Google Slides layout (default TITLE_AND_BODY)");

const alignment = z
  .enum(["START", "CENTER", "END", "JUSTIFIED"])
  .optional()
  .describe("Paragraph alignment");

const position = {
  leftInches: z.number().nonnegative().default(0.75).describe("Left, inches"),
  topInches: z.number().nonnegative().default(1.6).describe("Top, inches"),
  widthInches: z.number().positive().default(6.5).describe("Width, inches"),
  heightInches: z.number().positive().default(2.5).describe("Height, inches"),
};

server.registerTool(
  "create_presentation",
  {
    title: "Create presentation",
    description:
      "Create a new, empty Google Slides presentation and return its ID and URL.",
    inputSchema: {
      title: z.string().min(1).describe("Presentation title"),
    },
    annotations: {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    },
  },
  ({ title }) => run(() => createPresentation(title)),
);

server.registerTool(
  "create_deck_from_outline",
  {
    title: "Create deck from outline",
    description:
      "Create a complete, themed Google Slides deck from a structured outline in one call. The first slide defaults to a title slide; the rest default to title + bullets. Use this for most new decks, then refine with other tools.",
    inputSchema: {
      title: z.string().min(1).describe("Deck title"),
      subtitle: z
        .string()
        .optional()
        .describe("Subtitle for the title slide"),
      theme,
      fontFamily: z
        .string()
        .optional()
        .describe("Font family, e.g. Arial, Georgia, Roboto"),
      slides: z
        .array(
          z.object({
            title: z.string().describe("Slide title"),
            bullets: z
              .array(z.string())
              .optional()
              .describe("Bullet points for the slide body"),
            notes: z.string().optional().describe("Speaker notes"),
            layout: z
              .enum(["title", "section", "content", "blank"])
              .optional()
              .describe(
                "Slide kind. Defaults to title for slide 1 and content for the rest.",
              ),
          }),
        )
        .min(1)
        .describe("Ordered slides in the deck"),
    },
    annotations: {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    },
  },
  (args) => run(() => createDeckFromOutline(args)),
);

server.registerTool(
  "get_presentation",
  {
    title: "Get presentation",
    description:
      "Get a compact summary of a presentation: title, size, slides, their object IDs, element IDs, text, and speaker notes. Set raw=true for the full Google API response.",
    inputSchema: {
      presentationId,
      raw: z.boolean().default(false).describe("Return the raw API response"),
    },
    annotations: { readOnlyHint: true },
  },
  ({ presentationId, raw }) => run(() => (raw ? getPresentationRaw(presentationId) : getPresentation(presentationId))),
);

server.registerTool(
  "list_slides",
  {
    title: "List slides",
    description:
      "List every slide in a deck with its index, object ID, title and layout ID.",
    inputSchema: { presentationId },
    annotations: { readOnlyHint: true },
  },
  ({ presentationId }) => run(() => listSlides(presentationId)),
);

server.registerTool(
  "get_slide",
  {
    title: "Get slide",
    description:
      "Get one slide with element IDs, types, text, positions, sizes and speaker notes. Use this before editing elements.",
    inputSchema: { presentationId, slideId },
    annotations: { readOnlyHint: true },
  },
  ({ presentationId, slideId }) =>
    run(() => getSlide(presentationId, slideId)),
);

server.registerTool(
  "add_slide",
  {
    title: "Add slide",
    description:
      "Add a slide using a Google Slides layout. Optionally fill the title, body or bullet list, apply a theme, and add speaker notes.",
    inputSchema: {
      presentationId,
      layout,
      title: z.string().optional().describe("Title placeholder text"),
      body: z.string().optional().describe("Body placeholder text"),
      bullets: z
        .array(z.string())
        .optional()
        .describe("Bullet list for the body (overrides body)"),
      notes: z.string().optional().describe("Speaker notes"),
      theme,
      insertionIndex: z
        .number()
        .int()
        .nonnegative()
        .optional()
        .describe("Zero-based position to insert the slide"),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  (args) => run(() => addSlide(args)),
);

server.registerTool(
  "delete_slide",
  {
    title: "Delete slide",
    description: "Delete a slide (page) from a presentation.",
    inputSchema: { presentationId, slideId },
    annotations: {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: true,
    },
  },
  ({ presentationId, slideId }) =>
    run(() => deleteSlide(presentationId, slideId)),
);

server.registerTool(
  "duplicate_slide",
  {
    title: "Duplicate slide",
    description: "Duplicate a slide, optionally moving the copy to an index.",
    inputSchema: {
      presentationId,
      slideId,
      insertionIndex: z
        .number()
        .int()
        .nonnegative()
        .optional()
        .describe("Where the duplicate should be placed"),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  ({ presentationId, slideId, insertionIndex }) =>
    run(() => duplicateSlide(presentationId, slideId, insertionIndex)),
);

server.registerTool(
  "move_slide",
  {
    title: "Move slide",
    description: "Move a slide to a new zero-based index.",
    inputSchema: {
      presentationId,
      slideId,
      insertionIndex: z
        .number()
        .int()
        .nonnegative()
        .describe("Target zero-based index"),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  ({ presentationId, slideId, insertionIndex }) =>
    run(() => moveSlide(presentationId, slideId, insertionIndex)),
);

server.registerTool(
  "reorder_slides",
  {
    title: "Reorder slides",
    description:
      "Set the complete slide order. Pass every slide object ID exactly once, in the desired order.",
    inputSchema: {
      presentationId,
      orderedSlideIds: z
        .array(z.string().min(1))
        .min(1)
        .describe("All slide object IDs in the new order"),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  ({ presentationId, orderedSlideIds }) =>
    run(() => reorderSlides(presentationId, orderedSlideIds)),
);

server.registerTool(
  "add_text_box",
  {
    title: "Add text box",
    description:
      "Add a styled text box to a slide. Coordinates are in inches. Set bulleted=true to render the lines as a bullet list.",
    inputSchema: {
      presentationId,
      slideId,
      text: z.string().describe("Text content; use \\n for new lines"),
      ...position,
      fontSize: z.number().positive().default(18).describe("Font size in pt"),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      color: z.string().optional().describe("Text color as hex, e.g. #FFFFFF"),
      fontFamily: z.string().optional().describe("Font family, e.g. Arial"),
      alignment,
      bulleted: z.boolean().optional().describe("Render lines as bullets"),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  (args) => run(() => addTextBox(args)),
);

server.registerTool(
  "add_bullets",
  {
    title: "Add bullet list",
    description:
      "Add a bullet list text box to a slide. Convenience wrapper around add_text_box.",
    inputSchema: {
      presentationId,
      slideId,
      bullets: z.array(z.string()).min(1).describe("Bullet items"),
      ...position,
      fontSize: z.number().positive().default(18).describe("Font size in pt"),
      color: z.string().optional().describe("Text color as hex"),
      fontFamily: z.string().optional(),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  ({ bullets, ...args }) =>
    run(() =>
      addTextBox({
        ...args,
        text: bullets.join("\n"),
        bulleted: true,
      }),
    ),
);

server.registerTool(
  "update_text",
  {
    title: "Update text",
    description:
      "Replace all text inside a shape or text box. Get the object ID from get_slide.",
    inputSchema: {
      presentationId,
      objectId,
      text: z.string().describe("New text content"),
    },
    annotations: { readOnlyHint: false, destructiveHint: true },
  },
  ({ presentationId, objectId, text }) =>
    run(() => updateText(presentationId, objectId, text)),
);

server.registerTool(
  "replace_all_text",
  {
    title: "Replace all text",
    description:
      "Find and replace text across the whole deck, like Find & Replace in Slides.",
    inputSchema: {
      presentationId,
      findText: z.string().min(1).describe("Text to find"),
      replaceText: z.string().describe("Replacement text"),
      matchCase: z.boolean().default(false).describe("Case-sensitive match"),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  ({ presentationId, findText, replaceText, matchCase }) =>
    run(() => replaceAllText(presentationId, findText, replaceText, matchCase)),
);

server.registerTool(
  "format_text",
  {
    title: "Format text",
    description:
      "Style a character range inside a shape or text box. Indices are zero-based UTF-16 offsets. Use get_slide to find text first.",
    inputSchema: {
      presentationId,
      objectId,
      startIndex: z.number().int().nonnegative().describe("Start index"),
      endIndex: z.number().int().positive().describe("End index"),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      fontSize: z.number().positive().optional().describe("Font size in pt"),
      fontFamily: z.string().optional(),
      color: z.string().optional().describe("Text color as hex"),
      backgroundColor: z
        .string()
        .optional()
        .describe("Highlight color as hex"),
      linkUrl: z.string().url().optional().describe("Attach a hyperlink"),
      alignment,
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  (args) => run(() => formatText(args)),
);

server.registerTool(
  "add_image",
  {
    title: "Add image",
    description:
      "Insert an image from a publicly accessible URL into a slide at a given position and size (inches).",
    inputSchema: {
      presentationId,
      slideId,
      imageUrl: z.string().url().describe("Public image URL"),
      leftInches: z.number().nonnegative().default(0.75),
      topInches: z.number().nonnegative().default(1.8),
      widthInches: z.number().positive().default(6),
      heightInches: z.number().positive().default(3.8),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  (args) => run(() => addImage(args)),
);

server.registerTool(
  "add_shape",
  {
    title: "Add shape",
    description:
      "Add a shape (RECTANGLE, ROUND_RECTANGLE, ELLIPSE, TRIANGLE, DIAMOND, STAR_5, ARROW_RIGHT, etc.) with optional fill, outline and text.",
    inputSchema: {
      presentationId,
      slideId,
      shapeType: z
        .string()
        .min(1)
        .describe("Google Slides ShapeType, e.g. ROUND_RECTANGLE"),
      ...position,
      fillColor: z.string().optional().describe("Fill color as hex"),
      outlineColor: z.string().optional().describe("Outline color as hex"),
      text: z.string().optional().describe("Text inside the shape"),
      textColor: z.string().optional().describe("Text color as hex"),
      fontSize: z.number().positive().default(18),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  (args) => run(() => addShape(args)),
);

server.registerTool(
  "add_table",
  {
    title: "Add table",
    description:
      "Add a table to a slide, optionally filled with a 2D array of cell strings (row-major).",
    inputSchema: {
      presentationId,
      slideId,
      rows: z.number().int().positive(),
      columns: z.number().int().positive(),
      leftInches: z.number().nonnegative().default(0.75),
      topInches: z.number().nonnegative().default(1.8),
      widthInches: z.number().positive().default(8),
      heightInches: z.number().positive().default(3.5),
      data: z
        .array(z.array(z.string()))
        .optional()
        .describe("Row-major cell contents, e.g. [[\"A\",\"B\"],[\"1\",\"2\"]]"),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  (args) => run(() => addTable(args)),
);

server.registerTool(
  "set_slide_background",
  {
    title: "Set slide background",
    description: "Set a slide's background to a solid color.",
    inputSchema: {
      presentationId,
      slideId,
      color: z.string().describe("Background color as hex, e.g. #0B1220"),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  },
  ({ presentationId, slideId, color }) =>
    run(() => setSlideBackground(presentationId, slideId, color)),
);

server.registerTool(
  "set_speaker_notes",
  {
    title: "Set speaker notes",
    description: "Replace a slide's speaker notes.",
    inputSchema: {
      presentationId,
      slideId,
      notes: z.string().describe("Speaker notes text"),
    },
    annotations: { readOnlyHint: false, destructiveHint: true },
  },
  ({ presentationId, slideId, notes }) =>
    run(() => setSpeakerNotes(presentationId, slideId, notes)),
);

server.registerTool(
  "delete_element",
  {
    title: "Delete element",
    description:
      "Delete a shape, image, table or other page element. Use delete_slide to remove whole slides.",
    inputSchema: { presentationId, objectId },
    annotations: {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: true,
    },
  },
  ({ presentationId, objectId }) =>
    run(() => deleteElement(presentationId, objectId)),
);

server.registerTool(
  "export_deck",
  {
    title: "Export deck",
    description:
      "Export the presentation to PDF or PPTX with Google Drive and write it to the local filesystem.",
    inputSchema: {
      presentationId,
      format: z.enum(["pdf", "pptx"]).default("pdf"),
      outputPath: z
        .string()
        .optional()
        .describe("Absolute output path; defaults to ./<title>.<format>"),
    },
    annotations: { readOnlyHint: true, destructiveHint: false },
  },
  ({ presentationId, format, outputPath }) =>
    run(() => exportDeck(presentationId, format, outputPath)),
);

server.registerTool(
  "batch_update",
  {
    title: "Raw batch update",
    description:
      "Escape hatch: send raw Google Slides API batchUpdate requests. Use for advanced edits not covered by the other tools. See the Slides API request reference.",
    inputSchema: {
      presentationId,
      requests: z
        .array(z.record(z.string(), z.unknown()))
        .min(1)
        .describe("Array of Slides API Request objects"),
    },
    annotations: { readOnlyHint: false, destructiveHint: true },
  },
  ({ presentationId, requests }) =>
    run(() =>
      batchUpdate(
        presentationId,
        requests as Parameters<typeof batchUpdate>[1],
      ),
    ),
);

server.registerPrompt(
  "build_deck",
  {
    title: "Build a deck",
    description:
      "Plan and generate a polished Google Slides deck on a topic.",
    argsSchema: {
      topic: z.string().describe("What the deck is about"),
      audience: z.string().optional().describe("Who will see it"),
      slideCount: z.string().optional().describe("Approximate number of slides"),
      style: z.string().optional().describe("Tone or visual style"),
    },
  },
  ({ topic, audience, slideCount, style }) => ({
    messages: [
      {
        role: "user",
        content: {
          type: "text",
          text: [
            `Build a Google Slides deck about: ${topic}`,
            audience ? `Audience: ${audience}` : "",
            slideCount ? `Target slide count: ${slideCount}` : "",
            style ? `Style: ${style}` : "",
            "",
            "First outline the story, then call create_deck_from_outline with a strong title, punchy slide titles, 3-5 concise bullets per slide, and speaker notes.",
            "After it is created, fetch the deck with get_presentation to verify it, and return the URL.",
          ]
            .filter(Boolean)
            .join("\n"),
        },
      },
    ],
  }),
);

const transport = new StdioServerTransport();
await server.connect(transport);
console.error("google-slides-mcp running on stdio");
