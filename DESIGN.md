# AuditWorkpaper Pro — Design System & Tokens

Dokumentasi spesifikasi desain antarmuka, variabel CSS, tipografi, dan state komponen untuk aplikasi **AuditWorkpaper Pro**.

---

## Color Palette & Tokens

Sistem warna dirancang khusus untuk memenuhi standar keterbacaan tinggi dalam tugas audit profesional, dengan rasio kontras $\ge 4.5:1$ (WCAG AA) pada mode terang (*Light Mode*) dan mode gelap (*Dark Mode*).

### Light Mode Tokens
```css
:root {
  /* Brand / Primary */
  --color-primary: #1d4ed8;         /* Blue 700 - A11y compliant on light bg */
  --color-primary-hover: #1e40af;   /* Blue 800 */
  --color-primary-subtle: #eff6ff;  /* Blue 50 */
  --color-primary-border: #bfdbfe;  /* Blue 200 */

  /* Neutrals & Surfaces */
  --color-bg-canvas: #f8fafc;       /* Slate 50 */
  --color-bg-surface: #ffffff;      /* Pure white */
  --color-bg-subtle: #f1f5f9;       /* Slate 100 */
  --color-bg-muted: #e2e8f0;        /* Slate 200 */

  /* Text & Ink */
  --color-text-primary: #0f172a;    /* Slate 900 */
  --color-text-secondary: #334155;  /* Slate 700 - 9.5:1 contrast */
  --color-text-muted: #475569;      /* Slate 600 - 5.5:1 contrast (WCAG AA pass) */

  /* Borders */
  --color-border: #cbd5e1;          /* Slate 300 */
  --color-border-subtle: #e2e8f0;   /* Slate 200 */
  --color-border-focus: #2563eb;    /* Blue 600 */

  /* Domain Status & Accents */
  --color-excel: #15803d;           /* Green 700 */
  --color-excel-hover: #166534;     /* Green 800 */
  --color-excel-bg: #f0fdf4;        /* Green 50 */
  --color-pdf: #b91c1c;             /* Red 700 */
  --color-pdf-hover: #991b1b;       /* Red 800 */
  --color-pdf-bg: #fef2f2;          /* Red 50 */
  --color-warning: #b45309;         /* Amber 700 */
  --color-warning-bg: #fffbeb;      /* Amber 50 */

  /* Elevation Shadows */
  --shadow-xs: 0 1px 2px 0 rgba(15, 23, 42, 0.05);
  --shadow-sm: 0 1px 3px 0 rgba(15, 23, 42, 0.08), 0 1px 2px -1px rgba(15, 23, 42, 0.08);
  --shadow-md: 0 4px 6px -1px rgba(15, 23, 42, 0.08), 0 2px 4px -2px rgba(15, 23, 42, 0.06);
  --shadow-lg: 0 10px 15px -3px rgba(15, 23, 42, 0.08), 0 4px 6px -4px rgba(15, 23, 42, 0.04);
}
```

### Dark Mode Tokens
```css
[data-theme="dark"] {
  --color-primary: #3b82f6;         /* Blue 500 */
  --color-primary-hover: #60a5fa;   /* Blue 400 */
  --color-primary-subtle: #1e293b;  /* Slate 800 */
  --color-primary-border: #1e3a8a;  /* Blue 900 */

  --color-bg-canvas: #090d16;       /* Slate 950 Deep */
  --color-bg-surface: #111827;      /* Slate 900 */
  --color-bg-subtle: #1e293b;       /* Slate 800 */
  --color-bg-muted: #334155;        /* Slate 700 */

  --color-text-primary: #f8fafc;    /* Slate 50 */
  --color-text-secondary: #cbd5e1;  /* Slate 300 */
  --color-text-muted: #94a3b8;      /* Slate 400 */

  --color-border: #334155;          /* Slate 700 */
  --color-border-subtle: #1e293b;   /* Slate 800 */
  --color-border-focus: #60a5fa;    /* Blue 400 */

  --color-excel: #22c55e;
  --color-excel-hover: #4ade80;
  --color-excel-bg: #052e16;
  --color-pdf: #ef4444;
  --color-pdf-hover: #f87171;
  --color-pdf-bg: #450a0a;
  --color-warning: #f59e0b;
  --color-warning-bg: #451a03;

  --shadow-xs: 0 1px 2px 0 rgba(0, 0, 0, 0.3);
  --shadow-sm: 0 1px 3px 0 rgba(0, 0, 0, 0.4), 0 1px 2px -1px rgba(0, 0, 0, 0.4);
  --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.4), 0 2px 4px -2px rgba(0, 0, 0, 0.3);
  --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.5), 0 4px 6px -4px rgba(0, 0, 0, 0.3);
}
```

---

## Typography
- **Primary Body Font**: `'Public Sans', system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif`
- **Headings & Branding**: `'IBM Plex Sans', 'Public Sans', system-ui, sans-serif`
- **Numeric & Tabular Amounts**: `'JetBrains Mono', monospace`
- **Font Scale** (Rasio Modular 1.25):
  - `Display / H1`: `1.75rem` (28px) — Weight 700, Letter-spacing `-0.02em`
  - `Card Header / H2`: `1.25rem` (20px) — Weight 700, Letter-spacing `-0.01em`
  - `Section / H3`: `1.125rem` (18px) — Weight 600
  - `Body / Form Labels`: `0.9375rem` (15px) — Weight 400 / 600
  - `Subtext / Helpers`: `0.8125rem` (13px) — Weight 400
  - `Badge / Micro`: `0.75rem` (12px) — Weight 600, Letter-spacing `0.02em`

---

## Layout & Spacing
- **Border Radii**:
  - Small elements (Badges, Buttons): `8px` (`0.5rem`)
  - Inputs & Card controls: `10px` (`0.625rem`)
  - Cards & Dropzones: `16px` (`1rem`)
- **Touch Target Size**:
  - Minimal `44px × 44px` untuk semua target interaktif di perangkat sentuh.

---

## Component Interaction States
1. **Dropzone File**:
   - *Default*: Border 2px dashed `--color-border`, background `--color-bg-subtle`
   - *Hover / Drag Over / Active*: Border 2px dashed `--color-primary`, background `--color-primary-subtle`
   - *Focus-Visible*: Outline ring `3px solid var(--color-border-focus)`, offset `2px`
   - *File Ready State*: Border `--color-excel`, icon checklist hijau
2. **Interactive Radio Cards**:
   - *Default*: Border 1.5px solid `--color-border-subtle`, background `--color-bg-surface`
   - *Checked*: Border 2px solid `--color-primary`, background `--color-primary-subtle`, label highlighted
   - *Focus-Visible*: Ring outline terfokus
3. **Action Buttons**:
   - `Export Excel`: Hijau akuntansi (`--color-excel`), hover brightness/elevasi, active transform `scale(0.98)`
   - `Export PDF`: Merah dokumen (`--color-pdf`), hover brightness/elevasi, active transform `scale(0.98)`
