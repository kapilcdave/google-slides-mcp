import type { slides_v1 } from "googleapis";

export const EMU_PER_INCH = 914400;

export function inchesToEmu(inches: number): number {
  return Math.round(inches * EMU_PER_INCH);
}

export function rgbColor(hex: string): {
  red: number;
  green: number;
  blue: number;
} {
  const value = hex.replace(/^#/, "");
  const full =
    value.length === 3
      ? value
          .split("")
          .map((c) => c + c)
          .join("")
      : value;

  if (!/^[0-9a-fA-F]{6}$/.test(full)) {
    throw new Error(`Invalid hex color: ${hex}`);
  }

  return {
    red: parseInt(full.slice(0, 2), 16) / 255,
    green: parseInt(full.slice(2, 4), 16) / 255,
    blue: parseInt(full.slice(4, 6), 16) / 255,
  };
}

export function elementText(
  element: slides_v1.Schema$PageElement | undefined,
): string {
  if (!element) return "";

  const shapeText = element.shape?.text?.textElements
    ?.map((te) => te.textRun?.content ?? "")
    .join("");
  if (shapeText) return shapeText;

  const cellText =
    element.table?.tableRows
      ?.flatMap((row) => row.tableCells ?? [])
      .map((cell) =>
        (cell.text?.textElements ?? [])
          .map((te) => te.textRun?.content ?? "")
          .join(""),
      )
      .join(" | ") ?? "";
  if (cellText) return cellText;

  return element.wordArt?.renderedText ?? "";
}

export function pageTitle(page: slides_v1.Schema$Page): string {
  const elements = page.pageElements ?? [];
  const titleElement =
    elements.find((el) => el.shape?.placeholder?.type === "TITLE") ??
    elements.find((el) => el.shape?.text);
  return elementText(titleElement).trim();
}

export function elementType(
  element: slides_v1.Schema$PageElement,
): string {
  if (element.shape) return "shape";
  if (element.image) return "image";
  if (element.table) return "table";
  if (element.line) return "line";
  if (element.video) return "video";
  if (element.sheetsChart) return "sheets_chart";
  if (element.wordArt) return "word_art";
  if (element.elementGroup) return "group";
  return "unknown";
}

export function summarizeElement(
  element: slides_v1.Schema$PageElement,
): Record<string, unknown> {
  const text = elementText(element).trim();
  return {
    objectId: element.objectId,
    type: elementType(element),
    ...(element.shape?.shapeType
      ? { shapeType: element.shape.shapeType }
      : {}),
    ...(element.shape?.placeholder?.type
      ? { placeholder: element.shape.placeholder.type }
      : {}),
    ...(text ? { text } : {}),
    ...(element.image ? { image: true } : {}),
    ...(element.table
      ? {
          rows: element.table.rows,
          columns: element.table.columns,
        }
      : {}),
  };
}

export function presentationUrl(presentationId: string): string {
  return `https://docs.google.com/presentation/d/${presentationId}/edit`;
}

export function jsonResult(data: unknown) {
  return {
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(data, null, 2),
      },
    ],
  };
}

export async function run(
  fn: () => Promise<unknown>,
): Promise<
  | ReturnType<typeof jsonResult>
  | {
      content: { type: "text"; text: string }[];
      isError: true;
    }
> {
  try {
    return jsonResult(await fn());
  } catch (error) {
    const message =
      error instanceof Error ? error.message : String(error);
    return {
      content: [{ type: "text" as const, text: `Error: ${message}` }],
      isError: true,
    };
  }
}
