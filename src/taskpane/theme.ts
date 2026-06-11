import { webLightTheme, Theme } from "@fluentui/react-components";

// Brand palette distilled from the main web app's design tokens, so the add-in
// reads as the same product. Purple is the primary action color; orange is a
// single highlight accent (mirrors the app's `main-color` mixin usage).
export const brand = {
  purple: "#352676", // --brand-darker
  purpleHover: "#2c1f64",
  purplePressed: "#241a52",
  orange: "#ff6201", // app `main-color` highlight
  middle: "#f1edff", // --brand-middle (pastel)
  bgGrey: "#f9fafb", // --background-grey
  border: "#ddd", // --border-color
  positivePastel: "#d9f7e1", // --positive-pastel
};

// Stock Fluent light theme with brand/neutral/radius/font tokens overridden.
// Pragmatic for a POC vs. hand-rolling a 16-shade createLightTheme ramp — Fluent
// cascades these through every component automatically.
export const silksTheme: Theme = {
  ...webLightTheme,
  fontFamilyBase: "'Satoshi Variable', 'Satoshi', 'Helvetica', 'Arial', sans-serif",
  // radii — app uses 8px boxes/buttons
  borderRadiusSmall: "6px",
  borderRadiusMedium: "8px",
  borderRadiusLarge: "8px",
  // brand → primary buttons, focus rings, links
  colorBrandBackground: brand.purple,
  colorBrandBackgroundHover: brand.purpleHover,
  colorBrandBackgroundPressed: brand.purplePressed,
  colorBrandBackgroundSelected: brand.purpleHover,
  colorCompoundBrandBackground: brand.purple,
  colorCompoundBrandBackgroundHover: brand.purpleHover,
  colorCompoundBrandBackgroundPressed: brand.purplePressed,
  colorBrandForeground1: brand.purple,
  colorBrandForeground2: brand.purpleHover,
  colorCompoundBrandForeground1: brand.purple,
  colorCompoundBrandForeground1Hover: brand.purpleHover,
  colorBrandStroke1: brand.purple,
  colorBrandStroke2: brand.middle,
  colorNeutralStroke1: brand.border,
  // status colors to match the app palette
  colorPaletteRedForeground1: "#dc3545", // --error (also drives attachment "error" icon)
  colorPaletteGreenForeground1: "#66af65", // --success (attachment "ready" check icon)
  // softer hairline borders like the app's --border-color
  colorNeutralStroke2: brand.border,
};
