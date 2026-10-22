export type ThemeName = "midnight" | "light" | "sunset" | "forest";

export type Theme = {
  background: string;
  title: string;
  body: string;
  accent: string;
  fontFamily: string;
};

export const THEMES: Record<ThemeName, Theme> = {
  midnight: {
    background: "#0B1220",
    title: "#FFFFFF",
    body: "#C7D2FE",
    accent: "#6366F1",
    fontFamily: "Arial",
  },
  light: {
    background: "#FFFFFF",
    title: "#0F172A",
    body: "#334155",
    accent: "#2563EB",
    fontFamily: "Arial",
  },
  sunset: {
    background: "#1A0B2E",
    title: "#FDE68A",
    body: "#FBCFE8",
    accent: "#F97316",
    fontFamily: "Arial",
  },
  forest: {
    background: "#052E1F",
    title: "#ECFDF5",
    body: "#A7F3D0",
    accent: "#10B981",
    fontFamily: "Arial",
  },
};

export function themeFor(name: ThemeName | undefined): Theme {
  return THEMES[name ?? "midnight"];
}
