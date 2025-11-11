import { NextRequest, NextResponse } from "next/server";
import { load } from "cheerio";
import mammoth from "mammoth";
import DOMPurify from "isomorphic-dompurify";
import htmlDocx from "html-docx-js";

type StylePresetId = "classic" | "modern" | "executive";

const PRESET_CONFIG: Record<
  StylePresetId,
  {
    fontFamily: string;
    headingFont: string;
    bodySize: number;
    headingWeight: number;
    lineHeight: number;
    paragraphSpacing: number;
    accent: string;
    pageMargin: string;
    background: string;
  }
> = {
  classic: {
    fontFamily: '"Times New Roman", Times, serif',
    headingFont: '"Times New Roman", Times, serif',
    bodySize: 12,
    headingWeight: 600,
    lineHeight: 1.5,
    paragraphSpacing: 14,
    accent: "#1d4ed8",
    pageMargin: "1in",
    background: "#f8fafc",
  },
  modern: {
    fontFamily: '"Calibri", "Segoe UI", sans-serif',
    headingFont: '"Source Sans Pro", "Calibri", "Segoe UI", sans-serif',
    bodySize: 11,
    headingWeight: 600,
    lineHeight: 1.65,
    paragraphSpacing: 16,
    accent: "#0f766e",
    pageMargin: "0.9in",
    background: "#f0fdfa",
  },
  executive: {
    fontFamily: '"Helvetica Neue", Arial, sans-serif',
    headingFont: "Georgia, 'Times New Roman', serif",
    bodySize: 11,
    headingWeight: 700,
    lineHeight: 1.58,
    paragraphSpacing: 15,
    accent: "#a16207",
    pageMargin: "1in 1.15in 1in 1.15in",
    background: "#fffbeb",
  },
};

type FormatterOptions = {
  justify: boolean;
  tidySpacing: boolean;
  titleCaseHeadings: boolean;
  autoNumberHeadings: boolean;
  convertQuotes: boolean;
};

const parseBoolean = (value: FormDataEntryValue | null, fallback = false) => {
  if (typeof value !== "string") return fallback;
  return value === "true" || value === "1";
};

const toTitleCase = (input: string) => {
  const lower = input.toLowerCase();
  const words = lower.split(/\s+/);
  const smallWords = new Set(["and", "or", "the", "of", "for", "in", "on", "to", "a", "an", "with"]);
  return words
    .map((word, index) => {
      if (word.length === 0) return "";
      if (word.includes("-")) {
        return word
          .split("-")
          .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
          .join("-");
      }
      if (index !== 0 && index !== words.length - 1 && smallWords.has(word)) {
        return word;
      }
      return word.charAt(0).toUpperCase() + word.slice(1);
    })
    .join(" ");
};

const convertSmartQuotes = (value: string) => {
  let result = value;
  result = result.replace(/(^|[\s([{<])"/g, "$1“");
  result = result.replace(/"/g, "”");
  result = result.replace(/(^|[\s([{<])'/g, "$1‘");
  result = result.replace(/'/g, "’");
  return result;
};

const normalizeSpacing = ($: ReturnType<typeof load>) => {
  $("p").each((_, element) => {
    const el = $(element);
    const html = el.html() ?? "";
    const normalized = html
      .replace(/&nbsp;/g, " ")
      .replace(/\s{2,}/g, " ")
      .replace(/\s+([.,;:!?])/g, "$1")
      .trim();
    if (!normalized.length) {
      const previous = el.prev("p");
      if (previous.length === 0 || previous.text().trim().length === 0) {
        el.remove();
      } else {
        el.html("");
      }
    } else {
      el.html(normalized);
    }
  });

  $("p:empty").remove();
};

const applyTitleCase = ($: ReturnType<typeof load>) => {
  $("h1, h2, h3").each((_, element) => {
    const el = $(element);
    const current = el.text();
    el.text(toTitleCase(current));
  });
};

const applySmartQuotes = ($: ReturnType<typeof load>) => {
  $("body")
    .find("*")
    .contents()
    .each((_, node) => {
      if (node.type === "text" && node.nodeValue) {
        node.nodeValue = convertSmartQuotes(node.nodeValue);
      }
    });
};

const buildStyleSheet = (preset: (typeof PRESET_CONFIG)[StylePresetId], options: FormatterOptions) => {
  const numbering = options.autoNumberHeadings
    ? `
  .docx-content {
    counter-reset: h1;
  }
  .docx-content h1 {
    counter-reset: h2;
  }
  .docx-content h1::before {
    counter-increment: h1;
    content: counter(h1) ". ";
  }
  .docx-content h2 {
    counter-reset: h3;
  }
  .docx-content h2::before {
    counter-increment: h2;
    content: counter(h1) "." counter(h2) " ";
  }
  .docx-content h3::before {
    counter-increment: h3;
    content: counter(h1) "." counter(h2) "." counter(h3) " ";
  }
  .docx-content h1::before,
  .docx-content h2::before,
  .docx-content h3::before {
    font-weight: ${preset.headingWeight};
    color: ${preset.accent};
  }
`
    : "";

  return `
  @page {
    margin: ${preset.pageMargin};
  }

  body {
    font-family: ${preset.fontFamily};
    font-size: ${preset.bodySize}pt;
    line-height: ${preset.lineHeight};
    color: #1f2937;
    background: #ffffff;
    margin: 0;
  }

  .docx-shell {
    min-height: 100vh;
    background: ${preset.background};
    padding: 48px 0;
  }

  .docx-content {
    background: #ffffff;
    margin: 0 auto;
    max-width: 7in;
    padding: 1in 1.1in;
    box-shadow: 0 20px 45px -28px rgba(15, 23, 42, 0.45);
    border-radius: 12px;
  }

  .docx-content p {
    margin: ${preset.paragraphSpacing}pt 0;
    ${options.justify ? "text-align: justify;" : "text-align: left;"}
  }

  .docx-content h1,
  .docx-content h2,
  .docx-content h3,
  .docx-content h4,
  .docx-content h5,
  .docx-content h6 {
    font-family: ${preset.headingFont};
    font-weight: ${preset.headingWeight};
    letter-spacing: -0.02em;
    color: ${preset.accent};
    margin-top: 1.6em;
    margin-bottom: 0.5em;
  }

  .docx-content h1 { font-size: 28px; }
  .docx-content h2 { font-size: 22px; }
  .docx-content h3 { font-size: 18px; }

  .docx-content ul,
  .docx-content ol {
    margin: ${preset.paragraphSpacing}pt 0 ${preset.paragraphSpacing}pt 24px;
    padding: 0;
  }

  .docx-content ul li {
    margin-bottom: 6px;
  }

  .docx-content table {
    width: 100%;
    border-collapse: collapse;
    margin: 18px 0;
  }

  .docx-content table th,
  .docx-content table td {
    border: 1px solid #e2e8f0;
    padding: 8px 12px;
  }

  .docx-content blockquote {
    border-left: 4px solid ${preset.accent};
    padding-left: 18px;
    margin: ${preset.paragraphSpacing}pt 0;
    font-style: italic;
    color: #475569;
  }

  ${numbering}
`;
};

const buildHtmlDocument = (bodyMarkup: string, preset: StylePresetId, options: FormatterOptions) => {
  const config = PRESET_CONFIG[preset];
  const stylesheet = buildStyleSheet(config, options);
  const sanitizedContent = DOMPurify.sanitize(bodyMarkup);
  const wrapped = `
    <!doctype html>
    <html lang="en">
      <head>
        <meta charset="utf-8" />
        <title>Formatted Word Document</title>
        <style>${stylesheet}</style>
      </head>
      <body>
        <div class="docx-shell">
          <div class="docx-content">
            ${sanitizedContent}
          </div>
        </div>
      </body>
    </html>
  `;

  const preview = `
    <style>${stylesheet}</style>
    <div class="docx-content">
      ${sanitizedContent}
    </div>
  `;

  return { documentHtml: wrapped, previewHtml: preview };
};

export async function POST(request: NextRequest) {
  const formData = await request.formData();
  const file = formData.get("file");

  if (!(file instanceof File)) {
    return NextResponse.json({ error: "Upload a valid .docx file." }, { status: 400 });
  }

  const presetId = (formData.get("preset") as StylePresetId) || "classic";
  const chosenPreset: StylePresetId = PRESET_CONFIG[presetId] ? presetId : "classic";

  const options: FormatterOptions = {
    justify: parseBoolean(formData.get("justify"), true),
    tidySpacing: parseBoolean(formData.get("tidySpacing"), true),
    titleCaseHeadings: parseBoolean(formData.get("titleCaseHeadings"), true),
    autoNumberHeadings: parseBoolean(formData.get("autoNumberHeadings"), false),
    convertQuotes: parseBoolean(formData.get("convertQuotes"), true),
  };

  try {
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const { value: rawHtml } = await mammoth.convertToHtml({ buffer }, { includeDefaultStyleMap: true });
    const $ = load(`<div class="docx-root">${rawHtml}</div>`);

    if (options.tidySpacing) {
      normalizeSpacing($);
    }

    if (options.titleCaseHeadings) {
      applyTitleCase($);
    }

    if (options.convertQuotes) {
      applySmartQuotes($);
    }

    const bodyMarkup = $(".docx-root").html() ?? "";

    const { documentHtml, previewHtml } = buildHtmlDocument(bodyMarkup, chosenPreset, options);
    const docxBlob = htmlDocx.asBlob(documentHtml);
    const docxBuffer = Buffer.from(await docxBlob.arrayBuffer());
    const base64 = docxBuffer.toString("base64");

    return NextResponse.json({
      base64,
      previewHtml,
      appliedPreset: chosenPreset,
    });
  } catch (error) {
    console.error("[format-api]", error);
    return NextResponse.json(
      { error: "We could not format that document. Please ensure it is a valid .docx exported from Word." },
      { status: 500 },
    );
  }
}
