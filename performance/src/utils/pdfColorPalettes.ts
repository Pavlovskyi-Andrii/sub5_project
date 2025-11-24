export interface PdfColorPalette {
  name: string;
  primary: string;
  background: string;
  border: string;
  accent: string;
  text: string;
  lightBackground: string;
}

export const PDF_COLOR_PALETTES: Record<string, PdfColorPalette> = {
  green: {
    name: 'Verde Standard',
    primary: '#84cc16',
    background: '#f0f9ff',
    border: '#84cc16',
    accent: '#65a30d',
    text: '#333333',
    lightBackground: '#f7fee7'
  },
  blue: {
    name: 'Blu Professionale',
    primary: '#2563eb',
    background: '#eff6ff',
    border: '#2563eb',
    accent: '#1d4ed8',
    text: '#333333',
    lightBackground: '#dbeafe'
  },
  red: {
    name: 'Rosso Dinamico',
    primary: '#dc2626',
    background: '#fef2f2',
    border: '#dc2626',
    accent: '#b91c1c',
    text: '#333333',
    lightBackground: '#fee2e2'
  },
  orange: {
    name: 'Arancione Energetico',
    primary: '#ea580c',
    background: '#fff7ed',
    border: '#ea580c',
    accent: '#c2410c',
    text: '#333333',
    lightBackground: '#fed7aa'
  },
  dark: {
    name: 'Scuro Elegante',
    primary: '#374151',
    background: '#f9fafb',
    border: '#374151',
    accent: '#4b5563',
    text: '#333333',
    lightBackground: '#f3f4f6'
  }
};

export const getPalettePreviewColors = (paletteKey: string) => {
  const palette = PDF_COLOR_PALETTES[paletteKey];
  if (!palette) return PDF_COLOR_PALETTES.green;
  return palette;
};