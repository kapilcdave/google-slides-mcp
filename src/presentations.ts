import fs from "node:fs/promises";
import path from "node:path";
import { randomBytes } from "node:crypto";
import { google, type slides_v1 } from "googleapis";
import { createAuth } from "./auth.js";
import { themeFor, THEMES, type ThemeName } from "./theme.js";
import {
  elementText,
  inchesToEmu,
  pageTitle,
  presentationUrl,
  rgbColor,
  summarizeElement,
} from "./util.js";

type Request = slides_v1.Schema$Request;

let clientsCache:
  | {
      slides: slides_v1.Slides;
      drive: ReturnType<typeof google.drive>;
    }
  | undefined;

function clients() {
  if (!clientsCache) {
    const auth = createAuth();
    clientsCache = {
      slides: google.slides({ version: "v1", auth }),
      drive: google.drive({ version: "v3", auth }),
    };
  }
  return clientsCache;
}

function shortId(): string {
  return randomBytes(6).toString("hex");
}

async function batch(
  presentationId: string,
  requests: Request[],
): Promise<slides_v1.Schema$Response[]> {
  if (requests.length === 0) return [];
  const { slides } = clients();
  const response = await slides.presentations.batchUpdate({
    presentationId,
    requestBody: { requests },
  });
  return response.data.replies ?? [];
}

type TextStyleOptions = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  backgroundColor?: string;
  linkUrl?: string;
};

function updateTextStyle(
  objectId: string,
  endIndex: number,
  style: TextStyleOptions,
): Request {
  const textStyle: slides_v1.Schema$TextStyle = {};
  const fields: string[] = [];

  if (style.bold !== undefined) {
    textStyle.bold = style.bold;
    fields.push("bold");
  }
  if (style.italic !== undefined) {
    textStyle.italic = style.italic;
    fields.push("italic");
  }
  if (style.underline !== undefined) {
    textStyle.underline = style.underline;
    fields.push("underline");
  }
  if (style.fontSize !== undefined) {
    textStyle.fontSize = { magnitude: style.fontSize, unit: "PT" };
    fields.push("fontSize");
  }
  if (style.fontFamily !== undefined) {
    textStyle.fontFamily = style.fontFamily;
    fields.push("fontFamily");
  }
  if (style.color !== undefined) {
    textStyle.foregroundColor = {
      opaqueColor: { rgbColor: rgbColor(style.color) },
    };
    fields.push("foregroundColor");
  }
  if (style.backgroundColor !== undefined) {
    textStyle.backgroundColor = {
      opaqueColor: { rgbColor: rgbColor(style.backgroundColor) },
    };
    fields.push("backgroundColor");
  }
  if (style.linkUrl !== undefined) {
    textStyle.link = { url: style.linkUrl };
    fields.push("link");
  }

  return {
    updateTextStyle: {
      objectId,
      textRange: {
        type: "FIXED_RANGE",
        startIndex: 0,
        endIndex,
      },
      style: textStyle,
      fields: fields.join(","),
    },
  };
}

function backgroundRequest(slideId: string, hex: string): Request {
  return {
    updatePageProperties: {
      objectId: slideId,
      pageProperties: {
        pageBackgroundFill: {
          propertyState: "RENDERED",
          solidFill: { color: { rgbColor: rgbColor(hex) } },
        },
      },
      fields:
        "pageBackgroundFill.solidFill.color,pageBackgroundFill.propertyState",
    },
  };
}

function accentBarRequests(slideId: string, hex: string): Request[] {
  const objectId = `accent_${shortId()}`;
  return [
    {
      createShape: {
        objectId,
        shapeType: "RECTANGLE",
        elementProperties: {
          pageObjectId: slideId,
          size: {
            width: { magnitude: inchesToEmu(1.2), unit: "EMU" },
            height: { magnitude: inchesToEmu(0.06), unit: "EMU" },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: inchesToEmu(0.6),
            translateY: inchesToEmu(6.6),
            unit: "EMU",
          },
        },
      },
    },
    {
      updateShapeProperties: {
        objectId,
        shapeProperties: {
          shapeBackgroundFill: {
            solidFill: { color: { rgbColor: rgbColor(hex) } },
          },
          outline: { propertyState: "NOT_RENDERED" },
        },
        fields:
          "shapeBackgroundFill.solidFill.color,outline.propertyState",
      },
    },
  ];
}

export const PREDEFINED_LAYOUTS = [
  "BLANK",
  "CAPTION_ONLY",
  "TITLE",
  "TITLE_AND_BODY",
  "TITLE_AND_TWO_COLUMNS",
  "TITLE_ONLY",
  "SECTION_HEADER",
  "SECTION_TITLE_AND_DESCRIPTION",
  "ONE_COLUMN_TEXT",
  "MAIN_POINT",
  "BIG_NUMBER",
] as const;

export type PredefinedLayout = (typeof PREDEFINED_LAYOUTS)[number];

type PlaceholderPlan = {
  key: "title" | "body" | "subtitle" | "body2";
  type: string;
  index: number;
};

const LAYOUT_PLANS: Record<PredefinedLayout, PlaceholderPlan[]> = {
  BLANK: [],
  CAPTION_ONLY: [{ key: "body", type: "BODY", index: 0 }],
  TITLE: [
    { key: "title", type: "TITLE", index: 0 },
    { key: "subtitle", type: "SUBTITLE", index: 0 },
  ],
  TITLE_AND_BODY: [
    { key: "title", type: "TITLE", index: 0 },
    { key: "body", type: "BODY", index: 0 },
  ],
  TITLE_AND_TWO_COLUMNS: [
    { key: "title", type: "TITLE", index: 0 },
    { key: "body", type: "BODY", index: 0 },
    { key: "body2", type: "BODY", index: 1 },
  ],
  TITLE_ONLY: [{ key: "title", type: "TITLE", index: 0 }],
  SECTION_HEADER: [
    { key: "title", type: "TITLE", index: 0 },
    { key: "subtitle", type: "SUBTITLE", index: 0 },
  ],
  SECTION_TITLE_AND_DESCRIPTION: [
    { key: "title", type: "TITLE", index: 0 },
    { key: "body", type: "BODY", index: 0 },
  ],
  ONE_COLUMN_TEXT: [
    { key: "title", type: "TITLE", index: 0 },
    { key: "body", type: "BODY", index: 0 },
  ],
  MAIN_POINT: [{ key: "title", type: "TITLE", index: 0 }],
  BIG_NUMBER: [{ key: "title", type: "TITLE", index: 0 }],
};

export async function createPresentation(title: string) {
  const { slides } = clients();
  const response = await slides.presentations.create({
    requestBody: { title },
  });
  const presentationId = response.data.presentationId;
  if (!presentationId) {
    throw new Error("Google did not return a presentation ID.");
  }
  return {
    presentationId,
    title: response.data.title ?? title,
    url: presentationUrl(presentationId),
  };
}

export async function getPresentationRaw(presentationId: string) {
  const { slides } = clients();
  const response = await slides.presentations.get({ presentationId });
  return response.data;
}

export async function getPresentation(presentationId: string) {
  const { slides } = clients();
  const response = await slides.presentations.get({ presentationId });
  const presentation = response.data;

  return {
    presentationId,
    title: presentation.title,
    url: presentationUrl(presentationId),
    slideSize: presentation.pageSize,
    slides: (presentation.slides ?? []).map((slide, index) => ({
      index,
      objectId: slide.objectId,
      title: pageTitle(slide),
      layoutId: slide.slideProperties?.layoutObjectId,
      elements: (slide.pageElements ?? []).map(summarizeElement),
      speakerNotes: (
        slide.slideProperties?.notesPage?.pageElements ?? []
      )
        .filter(
          (el) =>
            el.objectId ===
            slide.slideProperties?.notesPage?.notesProperties
              ?.speakerNotesObjectId,
        )
        .map((el) => elementText(el).trim())
        .filter(Boolean)
        .join("\n"),
    })),
  };
}

export async function listSlides(presentationId: string) {
  const presentation = await getPresentation(presentationId);
  return {
    presentationId,
    title: presentation.title,
    url: presentation.url,
    slideCount: presentation.slides.length,
    slides: presentation.slides.map((slide) => ({
      index: slide.index,
      objectId: slide.objectId,
      title: slide.title,
      layoutId: slide.layoutId,
    })),
  };
}

export async function getSlide(
  presentationId: string,
  slideId: string,
) {
  const { slides } = clients();
  const response = await slides.presentations.pages.get({
    presentationId,
    pageObjectId: slideId,
  });
  const page = response.data;

  return {
    presentationId,
    objectId: page.objectId,
    title: pageTitle(page),
    layoutId: page.slideProperties?.layoutObjectId,
    notes: (page.slideProperties?.notesPage?.pageElements ?? [])
      .filter(
        (el) =>
          el.objectId ===
          page.slideProperties?.notesPage?.notesProperties
            ?.speakerNotesObjectId,
      )
      .map((el) => elementText(el))
      .join("\n"),
    elements: (page.pageElements ?? []).map((element) => ({
      ...summarizeElement(element),
      transform: element.transform,
      size: element.size,
    })),
  };
}

export async function addSlide(options: {
  presentationId: string;
  layout?: PredefinedLayout;
  title?: string;
  body?: string;
  bullets?: string[];
  notes?: string;
  theme?: ThemeName;
  insertionIndex?: number;
}) {
  const {
    presentationId,
    layout = "TITLE_AND_BODY",
    title,
    body,
    bullets,
    notes,
    theme: themeName,
    insertionIndex,
  } = options;

  const plan = LAYOUT_PLANS[layout];
  const uid = shortId();
  const slideId = `slide_${uid}`;
  const ids: Partial<Record<PlaceholderPlan["key"], string>> = {};
  const mappings: slides_v1.Schema$LayoutPlaceholderIdMapping[] = plan.map(
    (entry) => {
      const objectId = `${entry.key}_${uid}`;
      ids[entry.key] = objectId;
      return {
        layoutPlaceholder: { type: entry.type, index: entry.index },
        objectId,
      };
    },
  );

  const requests: Request[] = [
    {
      createSlide: {
        objectId: slideId,
        ...(insertionIndex !== undefined ? { insertionIndex } : {}),
        slideLayoutReference: { predefinedLayout: layout },
        placeholderIdMappings: mappings,
      },
    },
  ];

  const theme = themeName ? themeFor(themeName) : undefined;
  if (theme) {
    requests.push(backgroundRequest(slideId, theme.background));
  }

  if (ids.title && title) {
    requests.push({
      insertText: { objectId: ids.title, text: title, insertionIndex: 0 },
    });
    requests.push(
      updateTextStyle(ids.title, title.length, {
        bold: true,
        fontSize:
          layout === "TITLE" ? 40 : layout === "SECTION_HEADER" ? 28 : 30,
        fontFamily: theme?.fontFamily,
        color: theme?.title,
      }),
    );
  }

  const bodyText =
    bullets && bullets.length > 0 ? bullets.join("\n") : body;

  const bodyTarget = ids.body ?? ids.subtitle;
  if (bodyTarget && bodyText) {
    requests.push({
      insertText: {
        objectId: bodyTarget,
        text: bodyText,
        insertionIndex: 0,
      },
    });
    requests.push(
      updateTextStyle(bodyTarget, bodyText.length, {
        fontSize: 16,
        fontFamily: theme?.fontFamily,
        color: theme?.body,
      }),
    );
    if (bullets && bullets.length > 0) {
      requests.push({
        createParagraphBullets: {
          objectId: bodyTarget,
          textRange: {
            type: "FIXED_RANGE",
            startIndex: 0,
            endIndex: bodyText.length,
          },
          bulletPreset: "BULLET_DISC_CIRCLE_SQUARE",
        },
      });
    }
  }

  if (theme && layout !== "BLANK") {
    requests.push(...accentBarRequests(slideId, theme.accent));
  }

  await batch(presentationId, requests);

  if (notes) {
    await setSpeakerNotes(presentationId, slideId, notes);
  }

  return {
    presentationId,
    slideId,
    layout,
    url: presentationUrl(presentationId),
    placeholders: ids,
  };
}

export async function deleteSlide(
  presentationId: string,
  slideId: string,
) {
  await batch(presentationId, [{ deleteObject: { objectId: slideId } }]);
  return { presentationId, deletedSlideId: slideId };
}

export async function duplicateSlide(
  presentationId: string,
  slideId: string,
  insertionIndex?: number,
) {
  const requests: Request[] = [
    { duplicateObject: { objectId: slideId } },
  ];
  const replies = await batch(presentationId, requests);
  const newSlideId = replies[0]?.duplicateObject?.objectId;

  if (insertionIndex !== undefined && newSlideId) {
    await batch(presentationId, [
      {
        updateSlidesPosition: {
          slideObjectIds: [newSlideId],
          insertionIndex,
        },
      },
    ]);
  }

  return { presentationId, sourceSlideId: slideId, newSlideId };
}

export async function moveSlide(
  presentationId: string,
  slideId: string,
  insertionIndex: number,
) {
  await batch(presentationId, [
    {
      updateSlidesPosition: {
        slideObjectIds: [slideId],
        insertionIndex,
      },
    },
  ]);
  return { presentationId, slideId, insertionIndex };
}

export async function reorderSlides(
  presentationId: string,
  orderedSlideIds: string[],
) {
  const presentation = await getPresentation(presentationId);
  const existingIds = presentation.slides
    .map((slide) => slide.objectId)
    .filter((id): id is string => Boolean(id));

  if (
    orderedSlideIds.length !== existingIds.length ||
    new Set(orderedSlideIds).size !== orderedSlideIds.length ||
    orderedSlideIds.some((id) => !existingIds.includes(id))
  ) {
    throw new Error(
      `orderedSlideIds must contain every slide exactly once. Slides in deck: ${existingIds.join(", ")}`,
    );
  }

  await batch(presentationId, [
    {
      updateSlidesPosition: {
        slideObjectIds: orderedSlideIds,
        insertionIndex: 0,
      },
    },
  ]);

  return { presentationId, order: orderedSlideIds };
}

export async function addTextBox(options: {
  presentationId: string;
  slideId: string;
  text: string;
  leftInches?: number;
  topInches?: number;
  widthInches?: number;
  heightInches?: number;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  fontFamily?: string;
  alignment?: "START" | "CENTER" | "END" | "JUSTIFIED";
  bulleted?: boolean;
}) {
  const {
    presentationId,
    slideId,
    text,
    leftInches = 0.75,
    topInches = 1.6,
    widthInches = 6.5,
    heightInches = 2.5,
    fontSize = 18,
    bold,
    italic,
    underline,
    color,
    fontFamily,
    alignment,
    bulleted,
  } = options;

  const objectId = `text_${shortId()}`;
  const requests: Request[] = [
    {
      createShape: {
        objectId,
        shapeType: "TEXT_BOX",
        elementProperties: {
          pageObjectId: slideId,
          size: {
            width: { magnitude: inchesToEmu(widthInches), unit: "EMU" },
            height: { magnitude: inchesToEmu(heightInches), unit: "EMU" },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: inchesToEmu(leftInches),
            translateY: inchesToEmu(topInches),
            unit: "EMU",
          },
        },
      },
    },
    {
      insertText: { objectId, text, insertionIndex: 0 },
    },
    updateTextStyle(objectId, text.length, {
      fontSize,
      bold,
      italic,
      underline,
      fontFamily,
      color,
    }),
  ];

  if (alignment) {
    requests.push({
      updateParagraphStyle: {
        objectId,
        textRange: {
          type: "FIXED_RANGE",
          startIndex: 0,
          endIndex: text.length,
        },
        style: { alignment },
        fields: "alignment",
      },
    });
  }

  if (bulleted) {
    requests.push({
      createParagraphBullets: {
        objectId,
        textRange: {
          type: "FIXED_RANGE",
          startIndex: 0,
          endIndex: text.length,
        },
        bulletPreset: "BULLET_DISC_CIRCLE_SQUARE",
      },
    });
  }

  await batch(presentationId, requests);

  return {
    presentationId,
    slideId,
    objectId,
    url: presentationUrl(presentationId),
  };
}

export async function updateText(
  presentationId: string,
  objectId: string,
  text: string,
) {
  const requests: Request[] = [
    { deleteText: { objectId, textRange: { type: "ALL" } } },
    { insertText: { objectId, text, insertionIndex: 0 } },
  ];
  await batch(presentationId, requests);
  return { presentationId, objectId, text };
}

export async function replaceAllText(
  presentationId: string,
  findText: string,
  replaceText: string,
  matchCase = false,
) {
  const replies = await batch(presentationId, [
    {
      replaceAllText: {
        containsText: { text: findText, matchCase },
        replaceText,
      },
    },
  ]);
  return {
    presentationId,
    occurrencesChanged:
      replies[0]?.replaceAllText?.occurrencesChanged ?? 0,
  };
}

export async function formatText(options: {
  presentationId: string;
  objectId: string;
  startIndex: number;
  endIndex: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  backgroundColor?: string;
  linkUrl?: string;
  alignment?: "START" | "CENTER" | "END" | "JUSTIFIED";
}) {
  const { presentationId, objectId, startIndex, endIndex } = options;
  const styleRequest = updateTextStyle(objectId, endIndex, {
    bold: options.bold,
    italic: options.italic,
    underline: options.underline,
    fontSize: options.fontSize,
    fontFamily: options.fontFamily,
    color: options.color,
    backgroundColor: options.backgroundColor,
    linkUrl: options.linkUrl,
  });

  styleRequest.updateTextStyle!.textRange = {
    type: "FIXED_RANGE",
    startIndex,
    endIndex,
  };

  const requests: Request[] = [styleRequest];

  if (options.alignment) {
    requests.push({
      updateParagraphStyle: {
        objectId,
        textRange: {
          type: "FIXED_RANGE",
          startIndex,
          endIndex,
        },
        style: { alignment: options.alignment },
        fields: "alignment",
      },
    });
  }

  await batch(presentationId, requests);
  return { presentationId, objectId, startIndex, endIndex };
}

export async function addImage(options: {
  presentationId: string;
  slideId: string;
  imageUrl: string;
  leftInches?: number;
  topInches?: number;
  widthInches?: number;
  heightInches?: number;
}) {
  const {
    presentationId,
    slideId,
    imageUrl,
    leftInches = 0.75,
    topInches = 1.8,
    widthInches = 6,
    heightInches = 3.8,
  } = options;

  const objectId = `image_${shortId()}`;
  await batch(presentationId, [
    {
      createImage: {
        objectId,
        url: imageUrl,
        elementProperties: {
          pageObjectId: slideId,
          size: {
            width: { magnitude: inchesToEmu(widthInches), unit: "EMU" },
            height: { magnitude: inchesToEmu(heightInches), unit: "EMU" },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: inchesToEmu(leftInches),
            translateY: inchesToEmu(topInches),
            unit: "EMU",
          },
        },
      },
    },
  ]);

  return { presentationId, slideId, objectId };
}

export async function addShape(options: {
  presentationId: string;
  slideId: string;
  shapeType: string;
  leftInches?: number;
  topInches?: number;
  widthInches?: number;
  heightInches?: number;
  fillColor?: string;
  outlineColor?: string;
  text?: string;
  textColor?: string;
  fontSize?: number;
}) {
  const {
    presentationId,
    slideId,
    shapeType,
    leftInches = 0.75,
    topInches = 1.8,
    widthInches = 3,
    heightInches = 2,
    fillColor,
    outlineColor,
    text,
    textColor,
    fontSize = 18,
  } = options;

  const objectId = `shape_${shortId()}`;
  const shapeProperties: slides_v1.Schema$ShapeProperties = {};
  const fields: string[] = [];

  if (fillColor) {
    shapeProperties.shapeBackgroundFill = {
      solidFill: { color: { rgbColor: rgbColor(fillColor) } },
    };
    fields.push("shapeBackgroundFill.solidFill.color");
  }
  if (outlineColor) {
    shapeProperties.outline = {
      outlineFill: {
        solidFill: { color: { rgbColor: rgbColor(outlineColor) } },
      },
    };
    fields.push("outline.outlineFill.solidFill.color");
  }

  const requests: Request[] = [
    {
      createShape: {
        objectId,
        shapeType,
        elementProperties: {
          pageObjectId: slideId,
          size: {
            width: { magnitude: inchesToEmu(widthInches), unit: "EMU" },
            height: { magnitude: inchesToEmu(heightInches), unit: "EMU" },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: inchesToEmu(leftInches),
            translateY: inchesToEmu(topInches),
            unit: "EMU",
          },
        },
      },
    },
  ];

  if (fields.length > 0) {
    requests.push({
      updateShapeProperties: {
        objectId,
        shapeProperties,
        fields: fields.join(","),
      },
    });
  }

  if (text) {
    requests.push({
      insertText: { objectId, text, insertionIndex: 0 },
    });
    requests.push(
      updateTextStyle(objectId, text.length, {
        fontSize,
        color: textColor,
      }),
    );
  }

  await batch(presentationId, requests);
  return { presentationId, slideId, objectId };
}

export async function addTable(options: {
  presentationId: string;
  slideId: string;
  rows: number;
  columns: number;
  leftInches?: number;
  topInches?: number;
  widthInches?: number;
  heightInches?: number;
  data?: string[][];
}) {
  const {
    presentationId,
    slideId,
    rows,
    columns,
    leftInches = 0.75,
    topInches = 1.8,
    widthInches = 8,
    heightInches = 3.5,
    data,
  } = options;

  if (data && data.length > rows) {
    throw new Error("data has more rows than the table.");
  }
  if (data?.some((row) => row.length > columns)) {
    throw new Error("data has more columns than the table.");
  }

  const objectId = `table_${shortId()}`;
  const requests: Request[] = [
    {
      createTable: {
        objectId,
        rows,
        columns,
        elementProperties: {
          pageObjectId: slideId,
          size: {
            width: { magnitude: inchesToEmu(widthInches), unit: "EMU" },
            height: { magnitude: inchesToEmu(heightInches), unit: "EMU" },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: inchesToEmu(leftInches),
            translateY: inchesToEmu(topInches),
            unit: "EMU",
          },
        },
      },
    },
  ];

  if (data) {
    data.forEach((row, rowIndex) => {
      row.forEach((value, columnIndex) => {
        if (!value) return;
        requests.push({
          insertText: {
            objectId,
            cellLocation: { rowIndex, columnIndex },
            insertionIndex: 0,
            text: value,
          },
        });
      });
    });
  }

  await batch(presentationId, requests);
  return { presentationId, slideId, objectId, rows, columns };
}

export async function setSlideBackground(
  presentationId: string,
  slideId: string,
  color: string,
) {
  await batch(presentationId, [backgroundRequest(slideId, color)]);
  return { presentationId, slideId, color };
}

export async function setSpeakerNotes(
  presentationId: string,
  slideId: string,
  notes: string,
) {
  const { slides } = clients();
  const page = await slides.presentations.pages.get({
    presentationId,
    pageObjectId: slideId,
  });
  const notesId =
    page.data.slideProperties?.notesPage?.notesProperties?.speakerNotesObjectId;

  if (!notesId) {
    throw new Error("This slide has no speaker notes object.");
  }

  const existing = page.data.slideProperties?.notesPage?.pageElements?.find(
    (el) => el.objectId === notesId,
  );
  const requests: Request[] = [];

  if (existing && elementText(existing).length > 0) {
    requests.push({
      deleteText: { objectId: notesId, textRange: { type: "ALL" } },
    });
  }

  requests.push({
    insertText: { objectId: notesId, text: notes, insertionIndex: 0 },
  });

  await batch(presentationId, requests);
  return { presentationId, slideId, notes };
}

export async function deleteElement(
  presentationId: string,
  objectId: string,
) {
  await batch(presentationId, [{ deleteObject: { objectId } }]);
  return { presentationId, deletedObjectId: objectId };
}

export async function batchUpdate(
  presentationId: string,
  requests: Request[],
) {
  const replies = await batch(presentationId, requests);
  return { presentationId, replies };
}

export async function exportDeck(
  presentationId: string,
  format: "pdf" | "pptx" = "pdf",
  outputPath?: string,
) {
  const { slides, drive } = clients();
  const presentation = await slides.presentations.get({
    presentationId,
    fields: "title",
  });
  const title = presentation.data.title ?? presentationId;
  const mimeType =
    format === "pptx"
      ? "application/vnd.openxmlformats-officedocument.presentationml.presentation"
      : "application/pdf";

  const safeTitle =
    title.replace(/[^a-zA-Z0-9-_ ]/g, "").trim() || presentationId;
  const target =
    outputPath ??
    path.resolve(process.cwd(), `${safeTitle}.${format}`);

  const response = await drive.files.export(
    { fileId: presentationId, mimeType },
    { responseType: "arraybuffer" },
  );

  const buffer = Buffer.from(response.data as ArrayBuffer);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, buffer);

  return {
    presentationId,
    title,
    format,
    path: target,
    bytes: buffer.byteLength,
  };
}

export type OutlineSlide = {
  title: string;
  bullets?: string[];
  notes?: string;
  layout?: "title" | "section" | "content" | "blank";
};

export async function createDeckFromOutline(options: {
  title: string;
  subtitle?: string;
  slides: OutlineSlide[];
  theme?: ThemeName;
  fontFamily?: string;
}) {
  const {
    title,
    subtitle,
    slides: outline,
    theme: themeName = "midnight",
    fontFamily,
  } = options;

  if (outline.length === 0) {
    throw new Error("slides must contain at least one slide.");
  }

  const { slides } = clients();
  const created = await slides.presentations.create({
    requestBody: { title },
  });
  const presentationId = created.data.presentationId;
  if (!presentationId) {
    throw new Error("Google did not return a presentation ID.");
  }

  const theme = themeFor(themeName);
  const existing = await slides.presentations.get({ presentationId });
  const existingSlideIds = (existing.data.slides ?? [])
    .map((slide) => slide.objectId)
    .filter((id): id is string => Boolean(id));

  const requests: Request[] = [];
  const createdSlides: {
    index: number;
    slideId: string;
    layout: string;
    title: string;
  }[] = [];

  outline.forEach((item, index) => {
    const uid = shortId();
    const kind = item.layout ?? (index === 0 ? "title" : "content");

    let layout: PredefinedLayout;
    let plan: PlaceholderPlan[];
    if (kind === "title") {
      layout = "TITLE";
      plan = LAYOUT_PLANS.TITLE;
    } else if (kind === "section") {
      layout = "SECTION_HEADER";
      plan = LAYOUT_PLANS.SECTION_HEADER;
    } else if (kind === "blank") {
      layout = "BLANK";
      plan = [];
    } else {
      layout = "TITLE_AND_BODY";
      plan = LAYOUT_PLANS.TITLE_AND_BODY;
    }

    const slideId = `slide_${uid}`;
    const ids: Partial<Record<PlaceholderPlan["key"], string>> = {};
    const mappings = plan.map((entry) => {
      const objectId = `${entry.key}_${uid}`;
      ids[entry.key] = objectId;
      return {
        layoutPlaceholder: { type: entry.type, index: entry.index },
        objectId,
      };
    });

    requests.push({
      createSlide: {
        objectId: slideId,
        slideLayoutReference: { predefinedLayout: layout },
        placeholderIdMappings: mappings,
      },
    });
    requests.push(backgroundRequest(slideId, theme.background));

    const titleText = item.title ?? "";
    if (ids.title && titleText) {
      requests.push({
        insertText: {
          objectId: ids.title,
          text: titleText,
          insertionIndex: 0,
        },
      });
      requests.push(
        updateTextStyle(ids.title, titleText.length, {
          bold: true,
          fontSize:
            kind === "title" ? 40 : kind === "section" ? 28 : 30,
          fontFamily: fontFamily ?? theme.fontFamily,
          color: theme.title,
        }),
      );
    }

    let secondary = "";
    if (kind === "title") {
      secondary = item.bullets?.join("\n") || subtitle || "";
    } else if (kind === "section") {
      secondary = item.bullets?.join("\n") ?? "";
    }

    if (ids.subtitle && secondary) {
      requests.push({
        insertText: {
          objectId: ids.subtitle,
          text: secondary,
          insertionIndex: 0,
        },
      });
      requests.push(
        updateTextStyle(ids.subtitle, secondary.length, {
          fontSize: 18,
          fontFamily: fontFamily ?? theme.fontFamily,
          color: theme.body,
        }),
      );
    }

    const bodyText = item.bullets?.length
      ? item.bullets.join("\n")
      : "";

    if (ids.body && bodyText) {
      requests.push({
        insertText: {
          objectId: ids.body,
          text: bodyText,
          insertionIndex: 0,
        },
      });
      requests.push(
        updateTextStyle(ids.body, bodyText.length, {
          fontSize: 16,
          fontFamily: fontFamily ?? theme.fontFamily,
          color: theme.body,
        }),
      );
      if (item.bullets?.length) {
        requests.push({
          createParagraphBullets: {
            objectId: ids.body,
            textRange: {
              type: "FIXED_RANGE",
              startIndex: 0,
              endIndex: bodyText.length,
            },
            bulletPreset: "BULLET_DISC_CIRCLE_SQUARE",
          },
        });
      }
    }

    if (kind !== "blank") {
      requests.push(...accentBarRequests(slideId, theme.accent));
    }

    createdSlides.push({
      index,
      slideId,
      layout,
      title: titleText,
    });
  });

  for (const id of existingSlideIds) {
    requests.push({ deleteObject: { objectId: id } });
  }

  await batch(presentationId, requests);

  const withNotes = outline
    .map((item, index) => ({ item, slide: createdSlides[index] }))
    .filter(({ item }) => item.notes);

  if (withNotes.length > 0) {
    const fresh = await slides.presentations.get({ presentationId });
    const notesRequests: Request[] = [];

    for (const { item, slide } of withNotes) {
      const page = (fresh.data.slides ?? []).find(
        (candidate) => candidate.objectId === slide.slideId,
      );
      const notesId =
        page?.slideProperties?.notesPage?.notesProperties?.speakerNotesObjectId;
      if (notesId && item.notes) {
        notesRequests.push({
          insertText: {
            objectId: notesId,
            text: item.notes,
            insertionIndex: 0,
          },
        });
      }
    }

    await batch(presentationId, notesRequests);
  }

  return {
    presentationId,
    title,
    url: presentationUrl(presentationId),
    theme: themeName,
    slideCount: createdSlides.length,
    slides: createdSlides,
  };
}

export { THEMES };
