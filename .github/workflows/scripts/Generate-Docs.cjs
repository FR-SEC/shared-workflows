/**
 * =============================================================================
 *  Generate-Docs.js — AI-powered Word document generator for Confluence
 *
 *  Author        : Frederick Barton
 *  Version       : 2.1.0
 *  Last Updated  : 2026-03-30
 *  Environment   : Node.js 20+ / GitHub Actions
 *
 *  Change Log    :
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
  const exts    = ['.ps1', '.py', '.js'];
  const skip    = new Set(['node_modules', '.git', 'docs', '.github', '__pycache__', 'dist', 'build']);

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
  const MAX_BYTES = 50000;
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
  const prompt = `You are a technical writer analyzing a software project called "${projectName}" (version ${version}, repository: ${repo}).

Analyze the following source code and generate comprehensive documentation suitable for a Confluence page. The documentation should be written for end users and administrators, not just developers.

Generate documentation with EXACTLY these sections in this order, using these exact headings:

# Overview
A clear, concise description of what this application/script does and its primary purpose. Write 2-4 paragraphs.

# Key Features
List the main features and capabilities as bullet points.

# Requirements
List system requirements, dependencies, and prerequisites.

# Installation & Setup
Step-by-step installation and configuration instructions.

# Usage
How to use the application/script, including command-line arguments, GUI instructions, or API usage as appropriate.

# Configuration
All configuration options, parameters, and settings with descriptions.

# Examples
Practical, real-world examples of how to use this application/script.

# Known Limitations
Any known limitations, edge cases, or things the application does not support.

# Change Log
A summary of what changed in version ${version} based on any change log information found in the code headers.

Write in a clear, professional style. Use specific details from the code — real function names, actual parameters, true behavior. Do not be vague or generic. If the code is a GUI application describe the interface. If it is a CLI tool describe the commands. If it processes specific file formats or data types, name them explicitly.

Here is the source code to analyze:

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
      max_tokens: 4000,
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
    'Overview', 'Key Features', 'Requirements',
    'Installation & Setup', 'Usage', 'Configuration',
    'Examples', 'Known Limitations', 'Change Log',
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
  const docPath = path.join(docsDir, `${PROJECT_NAME}.docx`);
  fs.writeFileSync(docPath, buffer);

  console.log(`Generated: ${docPath}`);
})();
