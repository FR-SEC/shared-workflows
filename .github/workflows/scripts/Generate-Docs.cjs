/**
 * =============================================================================
 *  Generate-Docs.js — AI-powered Word document generator for Confluence
 *
 *  Author        : Frederick Barton
 *  Version       : 2.2.0
 *  Last Updated  : 2026-03-30
 *  Environment   : Node.js 20+ / GitHub Actions
 *
 *  Change Log    :
 *    2.2.0 - 2026-03-30
 *        - Improved documentation quality:
 *          * max_tokens raised from 4000 to 16000 (was truncating output)
 *          * Now reads .md files (CHANGELOG, CLAUDE, ROADMAP, APP_INFO) as
 *            first-class sources, not just script code
 *          * Per-file read cap raised from 50KB to 80KB
 *          * Prompt rewritten: demands specific, grounded content from all
 *            project files; added Architecture and Roadmap sections;
 *            Change Log now uses CHANGELOG.md as authoritative source
 *          * Added Architecture and Roadmap sections to document structure
 *
 *    2.1.0 - 2026-03-30
 *        - Applied Sectra SPX / DOC_STYLE.md brand tokens to all styles:
 *          Blue-500 (#3C73BB) for H1, Asphalt-500 (#1E3A5F) for H2/cover
 *          Times New Roman 11pt body, Silver palette for tables/borders
 *        - Cover page: Asphalt-500 full-width colour bar with Silver-50 text
 *        - Table headers: Blue-500 background, Silver-50 text
 *        - Code blocks: Silver-100 background, Silver-400 borders
 *        - Footer/header borders updated to Silver-400
 *
 *    2.0.0 - 2026-03-21
 *        - Replaced header parsing with Claude API codebase analysis
 *        - Generates full Confluence-ready Word document per project
 *        - Reads all scripts in repo and sends to Claude for analysis
 *        - Produces Purpose, Usage, Examples, Configuration, Change Log sections
 *    1.0.0 - 2026-03-21
 *        - Initial release with header parsing approach
 *
 * -----------------------------------------------------------------------------
 *  PURPOSE:
 *    Triggered by GitHub Actions on minor/major version tag pushes.
 *    Reads all scripts in the repo, sends them to the Claude API for analysis,
 *    and generates a Confluence-ready Word document saved to docs/.
 *
 *  USAGE:
 *    node Generate-Docs.js --version 1.2 --repo FR-SEC/HL7-Parser --tag v1.2
 *
 * =============================================================================
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, Footer, Header,
} = require('docx');
const fs   = require('fs');
const path = require('path');

// ---------------------------------------------------------------------------
// CLI args
// ---------------------------------------------------------------------------
const args = {};
process.argv.slice(2).forEach((a, i, arr) => {
  if (a.startsWith('--')) args[a.slice(2)] = arr[i + 1];
});
const VERSION       = args.version || '1.0';
const REPO          = args.repo    || '';
const TAG           = args.tag     || `v${VERSION}`;
const API_KEY       = process.env.ANTHROPIC_API_KEY || '';
const PROJECT_NAME  = REPO.split('/').pop() || 'Project';

// ---------------------------------------------------------------------------
// Find all scripts in repo
// ---------------------------------------------------------------------------
function findScripts(root) {
  const results = [];
  const exts    = ['.ps1', '.py', '.js', '.md', '.txt'];
  const skip    = new Set(['node_modules', '.git', 'docs', '.github', '__pycache__', 'dist', 'build', 'Output', 'icons']);

  function walk(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      if (skip.has(entry.name)) continue;
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) { walk(full); continue; }
      if (exts.includes(path.extname(entry.name).toLowerCase())) {
        results.push(full);
      }
    }
  }
  walk(root);
  return results;
}

// ---------------------------------------------------------------------------
// Read script contents (cap at 50KB per file to stay within API limits)
// ---------------------------------------------------------------------------
function readScripts(scriptPaths, root) {
  const MAX_BYTES = 80000;
  return scriptPaths.map(p => {
    const rel     = path.relative(root, p);
    const content = fs.readFileSync(p, 'utf8').slice(0, MAX_BYTES);
    return `=== FILE: ${rel} ===\n${content}`;
  }).join('\n\n');
}

// ---------------------------------------------------------------------------
// Call Claude API to analyze codebase and generate documentation
// ---------------------------------------------------------------------------
async function generateDocContent(codeContent, projectName, version, repo) {
  const prompt = `You are a senior technical writer producing professional Confluence documentation for "${projectName}" (version ${version}, repository: ${repo}).

You have been given the complete project source code AND supporting documentation files (APP_INFO.md, CHANGELOG.md, CLAUDE.md, ROADMAP.md, etc.). Use ALL of these sources — the markdown files contain authoritative information about the project's purpose, architecture, workflow, and change history that may not be obvious from the code alone.

Generate documentation with EXACTLY these sections in this order, using these exact headings:

# Overview
What this project does, who it is for, and why it exists. Write 3-5 substantive paragraphs covering the problem it solves, how it works at a high level, and what makes it valuable. Use specific details from APP_INFO.md and CLAUDE.md — do not write generic filler.

# Key Features
Bullet points with bold feature names and concise descriptions. Derive from actual code capabilities and APP_INFO.md feature list.

# Architecture
How the application is structured internally. Cover the major components, data flow, and key design decisions. Reference CLAUDE.md for architectural details. Include a description of the output format/structure if the application produces files.

# Requirements
System requirements, dependencies, and prerequisites. Be specific about versions.

# Installation & Setup
Step-by-step instructions from first download through first successful run. Include execution policy, certificate, and credential setup if applicable.

# Usage
Complete usage instructions with actual command examples. Cover first run, subsequent runs, all parameters and options. If there is a multi-step workflow (e.g. run on multiple servers then compare), document the full end-to-end process.

# Configuration
All configuration options with descriptions, default values, and valid ranges. Organize by category if applicable.

# Examples
3-5 real-world usage scenarios with actual commands and expected output. Not generic — use real parameter values and realistic context.

# Known Limitations
Specific, honest limitations derived from the code. Not generic caveats.

# Change Log
Use the CHANGELOG.md file (if present) as the authoritative source. Include all versions with their changes, not just the current version. Format each version as a subsection.

# Roadmap
If ROADMAP.md exists, include planned features. If empty or absent, state "No planned features at this time."

Write in a clear, factual, understated tone. This is internal technical documentation, not a product brochure. State what the tool does plainly — do not use superlatives, marketing language, or promotional phrasing. Avoid words like "powerful", "comprehensive", "robust", "seamless", "excels", "cutting-edge", "smart", "intelligent", or "optimized". Do not editorialize about the tool's value or quality — let the reader decide. Every statement must be grounded in the actual source material provided. Do not invent features, fabricate examples, or include placeholder text. If a section has limited information, keep it short rather than padding with generic content.

Here is the project content:

${codeContent}`;

  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type':         'application/json',
      'x-api-key':            API_KEY,
      'anthropic-version':    '2023-06-01',
    },
    body: JSON.stringify({
      model:      'claude-sonnet-4-20250514',
      max_tokens: 16000,
      messages:   [{ role: 'user', content: prompt }],
    }),
  });

  if (!response.ok) {
    const err = await response.text();
    throw new Error(`Claude API error ${response.status}: ${err}`);
  }

  const data = await response.json();
  return data.content[0].text;
}

// ---------------------------------------------------------------------------
// Parse Claude's markdown response into sections
// ---------------------------------------------------------------------------
function parseMarkdownSections(markdown) {
  const sections = {};
  let currentSection = null;
  let currentLines   = [];

  for (const line of markdown.split('\n')) {
    const h1 = line.match(/^#\s+(.+)/);
    if (h1) {
      if (currentSection) sections[currentSection] = currentLines.join('\n').trim();
      currentSection = h1[1].trim();
      currentLines   = [];
    } else {
      if (currentSection) currentLines.push(line);
    }
  }
  if (currentSection) sections[currentSection] = currentLines.join('\n').trim();
  return sections;
}

// ---------------------------------------------------------------------------
// Shared styles — Sectra SPX / DOC_STYLE.md tokens
// ---------------------------------------------------------------------------
const BLUE_500     = '3C73BB';  // Heading 1, table headers, hyperlinks, accent rules
const ASPHALT_500  = '1E3A5F';  // Heading 2, cover page bar, footer background
const ASPHALT_900  = '071326';  // Body text, darkest text
const SILVER_50    = 'F7F9FC';  // Text on dark backgrounds (cover page, table headers)
const SILVER_100   = 'EEF2F7';  // Table even rows, code block background
const SILVER_400   = 'C9D3DE';  // Table borders, horizontal rules
const border       = { style: BorderStyle.SINGLE, size: 1, color: SILVER_400 };
const borders      = { top: border, bottom: border, left: border, right: border };

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, size: 36, font: 'Arial', color: BLUE_500 })],
    spacing: { before: 480, after: 240 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE_500, space: 1 } },
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, size: 28, font: 'Arial', color: ASPHALT_500 })],
    spacing: { before: 360, after: 180 },
  });
}

function bodyText(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, font: 'Times New Roman', color: ASPHALT_900 })],
    spacing: { after: 160, line: 276 },  // 1.15 line spacing = 276 twips
  });
}

function bulletPara(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'bullets', level },
    children:  [new TextRun({ text, size: 22, font: 'Times New Roman', color: ASPHALT_900 })],
    spacing:   { after: 60 },
  });
}

function numberedPara(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'numbers', level },
    children:  [new TextRun({ text, size: 22, font: 'Times New Roman', color: ASPHALT_900 })],
    spacing:   { after: 60 },
  });
}

function codeBlock(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 20, font: 'Courier New', color: ASPHALT_900 })],
    spacing:  { before: 40, after: 40 },
    shading:  { fill: SILVER_100, type: ShadingType.CLEAR },
    indent:   { left: 720 },
    border: {
      top:    { style: BorderStyle.SINGLE, size: 1, color: SILVER_400, space: 1 },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: SILVER_400, space: 1 },
    },
  });
}

// ---------------------------------------------------------------------------
// Convert markdown text block to Word paragraphs
// ---------------------------------------------------------------------------
function markdownToParas(text) {
  if (!text) return [bodyText('—')];
  const paras   = [];
  let inCode    = false;
  let codeLines = [];

  for (const line of text.split('\n')) {
    // Code fence
    if (line.trim().startsWith('```')) {
      if (inCode) {
        codeLines.forEach(l => paras.push(codeBlock(l)));
        codeLines = [];
        inCode    = false;
      } else {
        inCode = true;
      }
      continue;
    }
    if (inCode) { codeLines.push(line); continue; }

    // Blank line
    if (!line.trim()) { paras.push(new Paragraph({ spacing: { after: 40 } })); continue; }

    // Numbered list
    const numMatch = line.match(/^\s*(\d+)\.\s+(.+)/);
    if (numMatch) { paras.push(numberedPara(numMatch[2].trim())); continue; }

    // Bullet list (-, *, •)
    const bulMatch = line.match(/^\s*[-*•]\s+(.+)/);
    if (bulMatch) {
      const indent = line.match(/^\s*/)[0].length;
      paras.push(bulletPara(bulMatch[1].trim(), indent >= 4 ? 1 : 0));
      continue;
    }

    // Inline code — strip backticks for Word
    const cleaned = line.replace(/`([^`]+)`/g, '$1').trim();
    // Bold — strip ** for Word
    const stripped = cleaned.replace(/\*\*([^*]+)\*\*/g, '$1');
    paras.push(bodyText(stripped));
  }

  // Flush unclosed code block
  if (codeLines.length) codeLines.forEach(l => paras.push(codeBlock(l)));

  return paras;
}

// ---------------------------------------------------------------------------
// Cover page
// ---------------------------------------------------------------------------
function coverPage(projectName, version, repo, tag) {
  return [
    // Asphalt-500 colour bar with title and subtitle in Silver-50
    new Paragraph({
      alignment: AlignmentType.CENTER,
      shading:   { fill: ASPHALT_500, type: ShadingType.CLEAR },
      children:  [new TextRun({ text: ' ', size: 24 })],
      spacing:   { before: 0, after: 0 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      shading:   { fill: ASPHALT_500, type: ShadingType.CLEAR },
      children:  [new TextRun({ text: projectName, bold: true, size: 48, font: 'Arial', color: SILVER_50 })],
      spacing:   { before: 480, after: 160 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      shading:   { fill: ASPHALT_500, type: ShadingType.CLEAR },
      children:  [new TextRun({ text: 'Application Documentation', size: 28, font: 'Arial', color: SILVER_50 })],
      spacing:   { after: 480 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      shading:   { fill: ASPHALT_500, type: ShadingType.CLEAR },
      children:  [new TextRun({ text: ' ', size: 24 })],
      spacing:   { after: 0 },
    }),
    new Paragraph({ spacing: { before: 480, after: 80 } }),
    // Metadata table
    new Table({
      width:        { size: 6480, type: WidthType.DXA },
      columnWidths: [2160, 4320],
      rows: [
        ['Document Title', `${projectName} Documentation`],
        ['Version',        version],
        ['Date',           new Date().toISOString().slice(0, 10)],
        ['Author',         'Frederick Barton'],
        ['Repository',     repo],
        ['Classification', 'Internal Use'],
      ].map(([label, value], idx) =>
        new TableRow({
          children: [
            new TableCell({
              borders, width: { size: 2160, type: WidthType.DXA },
              shading:  { fill: SILVER_100, type: ShadingType.CLEAR },
              margins:  { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 20, font: 'Arial', color: ASPHALT_900 })] })],
            }),
            new TableCell({
              borders, width: { size: 4320, type: WidthType.DXA },
              shading:  { fill: idx % 2 === 0 ? 'FFFFFF' : SILVER_100, type: ShadingType.CLEAR },
              margins:  { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: value, size: 20, font: 'Times New Roman', color: ASPHALT_900 })] })],
            }),
          ],
        })
      ),
    }),
    new Paragraph({ spacing: { after: 200 } }),
  ];
}

// ---------------------------------------------------------------------------
// Build full document
// ---------------------------------------------------------------------------
function buildDocument(projectName, version, repo, tag, sections) {
  const sectionOrder = [
    'Overview', 'Key Features', 'Architecture',
    'Requirements', 'Installation & Setup', 'Usage',
    'Configuration', 'Examples', 'Known Limitations',
    'Change Log', 'Roadmap',
  ];

  const children = [
    ...coverPage(projectName, version, repo, tag),
    new Paragraph({ pageBreakBefore: true, spacing: { after: 0 } }),
    ...sectionOrder.flatMap(title => {
      const content = sections[title];
      return [
        heading1(title),
        ...markdownToParas(content || ''),
        new Paragraph({ spacing: { after: 160 } }),
      ];
    }),
  ];

  return new Document({
    numbering: {
      config: [
        { reference: 'bullets', levels: [
          { level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: '\u25E6', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
        ]},
        { reference: 'numbers', levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        ]},
      ],
    },
    styles: {
      default: { document: { run: { font: 'Times New Roman', size: 22, color: ASPHALT_900 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 36, bold: true, font: 'Arial', color: BLUE_500 },
          paragraph: { spacing: { before: 480, after: 240 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 28, bold: true, font: 'Arial', color: ASPHALT_500 },
          paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 1 } },
      ],
    },
    sections: [{
      properties: {
        page: {
          size:   { width: 12240, height: 15840 },  // US Letter
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },  // 1 inch
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            children: [
              new TextRun({ text: `${projectName}  `, size: 18, font: 'Arial', color: SILVER_400 }),
              new TextRun({ text: `v${version}`, size: 18, font: 'Arial', color: BLUE_500 }),
            ],
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: SILVER_400, space: 1 } },
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [
              new TextRun({ text: `Frederick Barton  |  Generated ${new Date().toISOString().slice(0, 10)}`, size: 18, font: 'Arial', color: SILVER_400 }),
            ],
            alignment: AlignmentType.RIGHT,
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: SILVER_400, space: 1 } },
          })],
        }),
      },
      children,
    }],
  });
}

// ---------------------------------------------------------------------------
// Build HTML output
// ---------------------------------------------------------------------------
function buildHtml(projectName, version, repo, tag, sections) {
  const sectionOrder = [
    'Overview', 'Key Features', 'Architecture',
    'Requirements', 'Installation & Setup', 'Usage',
    'Configuration', 'Examples', 'Known Limitations',
    'Change Log', 'Roadmap',
  ];

  function mdToHtml(text) {
    if (!text) return '<p>—</p>';
    const lines  = text.split('\n');
    const out    = [];
    let inCode   = false;
    let inUl     = false;
    let inOl     = false;

    const closeList = () => {
      if (inUl) { out.push('</ul>'); inUl = false; }
      if (inOl) { out.push('</ol>'); inOl = false; }
    };

    for (const line of lines) {
      if (line.trim().startsWith('```')) {
        closeList();
        if (inCode) { out.push('</code></pre>'); inCode = false; }
        else        { out.push('<pre><code>');    inCode = true;  }
        continue;
      }
      if (inCode) { out.push(escHtml(line)); continue; }

      if (!line.trim()) { closeList(); out.push('<p></p>'); continue; }

      const numMatch = line.match(/^\s*\d+\.\s+(.+)/);
      if (numMatch) {
        if (!inOl) { closeList(); out.push('<ol>'); inOl = true; }
        out.push(`<li>${inlineFormat(numMatch[1])}</li>`);
        continue;
      }

      const bulMatch = line.match(/^\s*[-*•]\s+(.+)/);
      if (bulMatch) {
        if (!inUl) { closeList(); out.push('<ul>'); inUl = true; }
        out.push(`<li>${inlineFormat(bulMatch[1])}</li>`);
        continue;
      }

      closeList();
      out.push(`<p>${inlineFormat(line.trim())}</p>`);
    }
    closeList();
    return out.join('\n');
  }

  function escHtml(s) {
    return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

  function inlineFormat(s) {
    return s
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
      .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
      .replace(/`([^`]+)`/g, '<code>$1</code>');
  }

  const sectionsHtml = sectionOrder.map(title => {
    const content = sections[title] || '';
    return `
    <section>
      <h1>${escHtml(title)}</h1>
      ${mdToHtml(content)}
    </section>`;
  }).join('\n');

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escHtml(projectName)} ${escHtml(tag)} — User Guide</title>
  <style>
    :root {
      --blue-500:    #3C73BB;
      --asphalt-500: #1E3A5F;
      --asphalt-900: #071326;
      --silver-50:   #F7F9FC;
      --silver-100:  #EEF2F7;
      --silver-400:  #C9D3DE;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Times New Roman', serif;
      font-size: 11pt;
      color: var(--asphalt-900);
      background: #fff;
      max-width: 960px;
      margin: 0 auto;
      padding: 2rem;
    }
    header {
      background: var(--asphalt-500);
      color: var(--silver-50);
      padding: 2rem 2.5rem;
      margin: -2rem -2rem 2rem -2rem;
    }
    header h1 {
      font-family: Arial, sans-serif;
      font-size: 2rem;
      font-weight: bold;
      border: none;
      color: var(--silver-50);
      margin-bottom: 0.25rem;
    }
    header p { font-size: 0.9rem; opacity: 0.85; }
    .meta-table {
      width: 100%;
      border-collapse: collapse;
      margin: 1.5rem 0 2rem 0;
      font-family: Arial, sans-serif;
      font-size: 0.85rem;
    }
    .meta-table td {
      padding: 0.4rem 0.75rem;
      border: 1px solid var(--silver-400);
    }
    .meta-table td:first-child {
      background: var(--silver-100);
      font-weight: bold;
      width: 140px;
    }
    section { margin-bottom: 2.5rem; }
    section h1 {
      font-family: Arial, sans-serif;
      font-size: 1.3rem;
      font-weight: bold;
      color: var(--blue-500);
      border-bottom: 3px solid var(--blue-500);
      padding-bottom: 0.3rem;
      margin-bottom: 1rem;
    }
    p  { margin: 0.5rem 0; line-height: 1.6; }
    ul, ol { margin: 0.5rem 0 0.5rem 1.5rem; }
    li { margin: 0.25rem 0; line-height: 1.5; }
    pre {
      background: var(--silver-100);
      border: 1px solid var(--silver-400);
      padding: 0.75rem 1rem;
      margin: 0.75rem 0;
      overflow-x: auto;
      font-size: 0.85rem;
    }
    code {
      font-family: 'Courier New', monospace;
      font-size: 0.88em;
      background: var(--silver-100);
      padding: 0.1em 0.3em;
      border-radius: 2px;
    }
    pre code { background: none; padding: 0; }
    footer {
      border-top: 1px solid var(--silver-400);
      margin-top: 3rem;
      padding-top: 0.75rem;
      font-family: Arial, sans-serif;
      font-size: 0.8rem;
      color: var(--silver-400);
      text-align: right;
    }
  </style>
</head>
<body>
  <header>
    <h1>${escHtml(projectName)}</h1>
    <p>User Guide &nbsp;·&nbsp; ${escHtml(tag)}</p>
  </header>

  <table class="meta-table">
    <tr><td>Version</td><td>${escHtml(version)}</td></tr>
    <tr><td>Tag</td><td>${escHtml(tag)}</td></tr>
    <tr><td>Repository</td><td>${escHtml(repo)}</td></tr>
    <tr><td>Generated</td><td>${new Date().toISOString().slice(0, 10)}</td></tr>
    <tr><td>Author</td><td>Frederick Barton</td></tr>
    <tr><td>Classification</td><td>Internal Use</td></tr>
  </table>

  ${sectionsHtml}

  <footer>Frederick Barton &nbsp;|&nbsp; Generated ${new Date().toISOString().slice(0, 10)}</footer>
</body>
</html>`;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
(async () => {
  if (!API_KEY) { console.error('ERROR: ANTHROPIC_API_KEY not set'); process.exit(1); }

  const repoRoot = process.cwd();
  const docsDir  = path.join(repoRoot, 'docs');
  if (!fs.existsSync(docsDir)) fs.mkdirSync(docsDir, { recursive: true });

  console.log(`Analyzing ${PROJECT_NAME} v${VERSION}...`);

  // Collect all scripts
  const scripts     = findScripts(repoRoot);
  console.log(`Found ${scripts.length} script(s) to analyze`);

  if (scripts.length === 0) {
    console.log('No scripts found — skipping doc generation');
    return;
  }

  // Read script contents
  const codeContent = readScripts(scripts, repoRoot);

  // Read APP_INFO.md if it exists
  let appInfoContent = '';
  const appInfoPath = path.join(repoRoot, 'APP_INFO.md');
  if (fs.existsSync(appInfoPath)) {
    appInfoContent = fs.readFileSync(appInfoPath, 'utf8');
    console.log('Found APP_INFO.md — including deployment-specific information');
  }

  const totalContent = appInfoContent
    ? `=== DEPLOYMENT INFORMATION (APP_INFO.md) ===\n${appInfoContent}\n\n=== SOURCE CODE ===\n${codeContent}`
    : codeContent;

  console.log(`Sending ${Math.round(totalContent.length / 1024)}KB to Claude API...`);

  // Call Claude API
  const markdown = await generateDocContent(totalContent, PROJECT_NAME, VERSION, REPO);
  console.log('Claude API response received');

  // Parse sections
  const sections = parseMarkdownSections(markdown);
  console.log(`Parsed sections: ${Object.keys(sections).join(', ')}`);

  // Build Word document
  const doc     = buildDocument(PROJECT_NAME, VERSION, REPO, TAG, sections);
  const buffer  = await Packer.toBuffer(doc);
const docPath  = path.join(docsDir, `${PROJECT_NAME}.docx`);
  fs.writeFileSync(docPath, buffer);
  console.log(`Generated: ${docPath}`);

  const html      = buildHtml(PROJECT_NAME, VERSION, REPO, TAG, sections);
  const htmlName  = `${PROJECT_NAME}-${TAG}-UserGuide.html`;
  const htmlPath  = path.join(docsDir, htmlName);
  fs.writeFileSync(htmlPath, html, 'utf8');
  console.log(`Generated: ${htmlPath}`);
})();
