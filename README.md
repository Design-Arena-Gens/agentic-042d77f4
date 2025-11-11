## Word Format Studio

Word Format Studio transforms rough Microsoft Word documents into clean, publication-ready deliverables. Upload a `.docx` file, pick a preset, toggle finishing touches, preview the results instantly, and export a polished `.docx` in seconds.

### Feature Highlights
- **Curated presets** — Classic Report, Modern Minimal, and Executive Brief ensure consistent fonts, heading hierarchy, and spacing.
- **Smart tidy options** — Justify body copy, normalize blank space, title‑case headings, auto-number H1–H3, and convert straight quotes to smart quotes.
- **Live preview** — Rendered with Tailwind Typography so you can inspect headings, tables, and lists before downloading.
- **Safe formatting pipeline** — Mammoth converts `.docx` to HTML, Cheerio cleans structure, DOMPurify sanitizes, and `html-docx-js` rebuilds the final Word file.

### Local Development
```bash
npm install
npm run dev
```
Visit `http://localhost:3000` and interact with the studio UI.

### Production Build
```bash
npm run build
npm run start
```

### Scripts
- `npm run dev` – start the development server
- `npm run build` – produce an optimized production build
- `npm run start` – serve the production build
- `npm run lint` – lint the codebase with ESLint

### Architecture Overview
- **Front end:** Next.js App Router with Tailwind CSS v4 styling
- **Formatter API:** `/api/format` converts the uploaded `.docx` to HTML (Mammoth), applies preset styles and tidying rules (Cheerio), sanitizes (DOMPurify), and exports a brand-new `.docx` (html-docx-js)
- **TypeScript:** Strict mode across the project plus a custom type definition for `html-docx-js`

### Project Layout
```
app/
  api/format/route.ts     # Document formatting pipeline
  layout.tsx              # Global layout + metadata
  page.tsx                # Upload UI, options, and preview
  globals.css             # Tailwind + typography adjustments
public/                   # Static assets
types/                    # Ambient type declarations
```

### License
MIT
