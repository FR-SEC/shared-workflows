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
const prompt = `You are an expert technical writer and systems integration engineer. Your audience is Sectra Integration Engineers — experienced Windows/PowerShell practitioners working in medical imaging. Write with the precision and depth of internal engineering documentation, not a product brochure.

Using the project source code and supporting files provided, produce a System Administrator Guide for "${projectName}" (version ${version}, repository: ${repo}).

Tone and style rules:
- Write for practitioners, not beginners. Do not explain PowerShell basics or Windows fundamentals.
- Every statement must be grounded in the actual source material. Do not invent behavior.
- Use specific names: actual parameter names, real file paths, true function behavior.
- Do not use superlatives (powerful, robust, seamless, comprehensive, intelligent).
- Keep prose concise. Prefer tables and lists over paragraphs when enumerating items.
- Use blockquotes (> text) for key design principles, important warnings, and operational notes. These will render as callout boxes.

Generate documentation with EXACTLY these sections in this order. Use ## for subsection headings within each section.

# 1. Introduction
What this tool does and why it exists. 2-3 paragraphs. State the core design principle (e.g. what it does NOT do directly, what it delegates). End with a blockquote summarizing the key design principle.

# 2. Prerequisites
## 2.1 System Requirements
Table: Requirement | Details. Cover OS, PowerShell version, network, disk.

## 2.2 Required Vendor Scripts
Table: Script | Purpose. List every vendor script the tool locates and calls at runtime. Mark which are required vs optional.

## 2.3 Optional Tools
Table: Tool | Purpose | Behavior When Missing. Cover any optional executables.

## 2.4 Credential Requirements
What credentials are needed, how they are stored (DPAPI, Export-Clixml, etc.), where the files are stored.

# 3. Parameters & Command-Line Reference
Table: Parameter | Type | Default | Description. Cover every param() parameter. Include switch parameters.

# 4. First Run & Setup
Step-by-step walkthrough of what happens on first run: vendor script discovery, credential prompts, config persistence. Include the exact command to launch normally and with common startup flags.

# 5. Operations & Workflow
## 5.1 [Primary workflow]
## 5.2 [Secondary workflow]
Document the main operational workflows with step-by-step numbered instructions. Include the menu structure if applicable.

# 6. Configuration Reference
All persisted configuration: what is stored, where, and how to reset it. Table: Setting | Storage Location | Reset Command.

# 7. Examples
5-6 real command examples with realistic context (e.g. "Pre-conversion export for a cloud migration project"). Use actual parameter values. Each example: a short title, the command block, and 1-2 sentences of context. Do not use generic placeholder values.

# 8. Error Guidance & Troubleshooting
Table or subsections covering common failure modes with: Symptom | Likely Cause | Resolution. Be specific — name the actual error text or behavior, not generic categories.

# 9. File Reference
## 9.1 Project Files
Table: File | Purpose

## 9.2 Output Files
Table: File | Location | Description

## 9.3 Persisted State Files
Table: File | Location | Reset Command

# 10. Related Tools
Table: Project | Script | Purpose | Output. List other tools in the Sectra I/O engineering suite that complement this tool (e.g. Mirth-Backup, Cloud Conversion Toolkit, SCH ConfigManager exports). Include a short paragraph describing how outputs from this tool integrate with the others during cloud conversion workflows.

# 11. Version History
Use CHANGELOG.md as the authoritative source. Include all versions. Format each version as ## vX.Y.Z — YYYY-MM-DD with subsections for change categories.

# 12. Document History
Single-entry table: Version | Date | Changes. First entry should be "1.0.0 | [today's date] | Initial User Guide covering [tool name] [version]."

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
      '1. Introduction',
      '2. Prerequisites',
      '3. Parameters & Command-Line Reference',
      '4. First Run & Setup',
      '5. Operations & Workflow',
      '6. Configuration Reference',
      '7. Examples',
      '8. Error Guidance & Troubleshooting',
      '9. File Reference',
      '10. Related Tools',
      '11. Version History',
      '12. Document History',
    ];

  function escHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  function inlineFormat(s) {
    return s
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
      .replace(/`([^`]+)`/g, '<code>$1</code>');
  }

function mdToHtml(text) {
    if (!text) return '<p>—</p>';
    const lines     = text.split('\n');
    const out       = [];
    let inCode      = false;
    let codeLang    = '';
    let codeLines   = [];
    let inUl        = false;
    let inOl        = false;
    let inNestedUl  = false;
    let inTable     = false;
    let tableRows   = [];

    const closeNestedUl = () => {
      if (inNestedUl) { out.push('</ul>'); inNestedUl = false; }
    };
    const closeList = () => {
      closeNestedUl();
      if (inUl) { out.push('</ul>'); inUl = false; }
      if (inOl) { out.push('</ol>'); inOl = false; }
    };

    const flushTable = () => {
      if (!inTable || tableRows.length === 0) return;
      const header = tableRows[0];
      const body   = tableRows.slice(2);
      let html = '<table><thead><tr>';
      for (const cell of header) html += `<th>${inlineFormat(cell.trim())}</th>`;
      html += '</tr></thead><tbody>';
      for (const row of body) {
        if (!row.length) continue;
        html += '<tr>';
        for (const cell of row) html += `<td>${inlineFormat(cell.trim())}</td>`;
        html += '</tr>';
      }
      html += '</tbody></table>';
      out.push(html);
      tableRows = [];
      inTable   = false;
    };

    for (const line of lines) {
      // Code fence
      if (line.trim().startsWith('```')) {
        closeList(); flushTable();
        if (inCode) {
          // suppress empty code blocks
          const hasContent = codeLines.some(l => l.trim().length > 0);
          if (hasContent) {
            out.push(`<pre><code${codeLang ? ` class="language-${escHtml(codeLang)}"` : ''}>`);
            out.push(...codeLines.map(l => escHtml(l)));
            out.push('</code></pre>');
          }
          codeLines = []; inCode = false; codeLang = '';
        } else {
          codeLang = line.trim().slice(3).trim();
          inCode   = true;
        }
        continue;
      }
      if (inCode) { codeLines.push(line); continue; }

      // Table rows
      if (line.trim().startsWith('|')) {
        closeList();
        inTable = true;
        tableRows.push(line.trim().replace(/^\||\|$/g, '').split('|'));
        continue;
      } else if (inTable) {
        flushTable();
      }

      // Blank line
      if (!line.trim()) {
        closeList();
        out.push('<p></p>');
        continue;
      }

      // Headings
      const h4 = line.match(/^####\s+(.+)/);
      if (h4) { closeList(); out.push(`<h4>${inlineFormat(h4[1])}</h4>`); continue; }
      const h3 = line.match(/^###\s+(.+)/);
      if (h3) { closeList(); out.push(`<h3>${inlineFormat(h3[1])}</h3>`); continue; }
      const h2 = line.match(/^##\s+(.+)/);
      if (h2) { closeList(); out.push(`<h2>${inlineFormat(h2[1])}</h2>`); continue; }

      // Blockquote → callout
      const bq = line.match(/^>\s*(.*)/);
      if (bq) {
        closeList();
        const cls = /warning|caution|danger/i.test(bq[1]) ? 'callout callout-warning'
                  : /tip|note|important/i.test(bq[1])     ? 'callout callout-info'
                  : 'callout';
        out.push(`<div class="${cls}"><strong>Note:</strong> ${inlineFormat(bq[1])}</div>`);
        continue;
      }

      // Numbered list
      const numMatch = line.match(/^\s*\d+\.\s+(.+)/);
      if (numMatch) {
        closeNestedUl();
        if (!inOl) { closeList(); out.push('<ol>'); inOl = true; }
        out.push(`<li>${inlineFormat(numMatch[1].trim())}</li>`);
        continue;
      }

      // Bullet list — nest inside ol if one is open
      const bulMatch = line.match(/^\s*[-*•]\s+(.+)/);
      if (bulMatch) {
        if (inOl) {
          if (!inNestedUl) { out.push('<ul>'); inNestedUl = true; }
        } else {
          if (!inUl) { closeList(); out.push('<ul>'); inUl = true; }
        }
        out.push(`<li>${inlineFormat(bulMatch[1].trim())}</li>`);
        continue;
      }

      closeList();
      out.push(`<p>${inlineFormat(line.trim())}</p>`);
    }

    closeList();
    flushTable();
    if (inCode && codeLines.some(l => l.trim().length > 0)) {
      out.push(`<pre><code${codeLang ? ` class="language-${escHtml(codeLang)}"` : ''}>`);
      out.push(...codeLines.map(l => escHtml(l)));
      out.push('</code></pre>');
    }

    return out.join('\n');
  }

  const repoName = repo.split('/').pop() || projectName;

  const sectionsHtml = sectionOrder.map(title => {
    // Match sections whether Claude prefixes with number or not
    const key = Object.keys(sections).find(k =>
      k === title || k.replace(/^\d+\.\s*/, '') === title.replace(/^\d+\.\s*/, '')
    ) || title;
    const content = sections[key] || sections[title] || '';
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
    @import url('https://fonts.googleapis.com/css2?family=Source+Code+Pro:wght@400;600&display=swap');
    :root {
      --blue-500:    #3C73BB;
      --asphalt-500: #1E3A5F;
      --asphalt-900: #071326;
      --silver-50:   #F7F9FC;
      --silver-100:  #EEF2F7;
      --silver-400:  #C9D3DE;
      --warning-bg:  #FDECEA;
      --warning-text:#B71C1C;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Times New Roman', Times, serif;
      font-size: 11pt;
      line-height: 1.4;
      color: var(--asphalt-900);
      background: #fff;
      max-width: 960px;
      margin: 0 auto;
      padding: 40px;
    }
    .cover {
      background: var(--asphalt-500);
      color: var(--silver-50);
      padding: 30px 40px;
      margin: -40px -40px 30px -40px;
    }
    .cover h1 {
      font-family: Arial, Helvetica, sans-serif;
      font-size: 24pt; font-weight: bold;
      margin: 0 0 8px 0; color: var(--silver-50); border: none;
    }
    .cover .subtitle {
      font-family: Arial, Helvetica, sans-serif;
      font-size: 14pt; color: var(--silver-50); opacity: .9;
    }
    .cover .meta {
      font-family: Arial, Helvetica, sans-serif;
      font-size: 10pt; margin-top: 16px; color: var(--silver-50); opacity: .8;
    }
    h1 {
      font-family: Arial, Helvetica, sans-serif;
      font-size: 18pt; font-weight: bold;
      color: var(--blue-500);
      border-bottom: 2px solid var(--blue-500);
      padding-bottom: 6px;
      margin-top: 36px; margin-bottom: 12px;
    }
    h2 {
      font-family: Arial, Helvetica, sans-serif;
      font-size: 14pt; font-weight: bold;
      color: var(--asphalt-500);
      margin-top: 24px; margin-bottom: 9px;
    }
    h3 {
      font-family: Arial, Helvetica, sans-serif;
      font-size: 12pt; font-weight: bold;
      color: var(--asphalt-500);
      margin-top: 18px; margin-bottom: 6px;
    }
    h4 {
      font-family: Arial, Helvetica, sans-serif;
      font-size: 11pt; font-weight: bold; font-style: italic;
      color: var(--asphalt-900);
      margin-top: 12px; margin-bottom: 4px;
    }
    section { margin-bottom: 0; }
    p  { margin: 0 0 8px 0; }
    ul, ol { margin: 0 0 8px 20px; }
    li { margin: 0 0 4px 0; }
    ul.nested { margin-left: 36px; }
    table {
      border-collapse: collapse; width: 100%;
      margin: 12px 0; font-size: 10pt;
    }
    th {
      background: var(--blue-500); color: var(--silver-50);
      font-family: Arial, Helvetica, sans-serif; font-weight: bold;
      text-align: left; padding: 6px 10px;
      border: 1px solid var(--silver-400);
    }
    td {
      padding: 4px 10px;
      border: 1px solid var(--silver-400);
      font-family: 'Times New Roman', Times, serif;
      vertical-align: top;
    }
    tr:nth-child(even) td { background: var(--silver-100); }
    pre, code {
      font-family: 'Source Code Pro', 'Courier New', Courier, monospace;
      font-size: 9pt;
    }
    pre {
      background: var(--silver-100);
      border: 1px solid var(--silver-400);
      border-left: 4px solid var(--blue-500);
      padding: 12px 16px;
      overflow-x: auto; line-height: 1.4;
      margin: 8px 0 12px 0;
    }
    code {
      background: var(--silver-100);
      padding: 1px 4px; border-radius: 3px;
    }
    pre code { background: none; padding: 0; }
    .callout {
      border-left: 4px solid var(--blue-500);
      background: var(--silver-100);
      padding: 12px 16px; margin: 12px 0;
    }
    .callout strong { font-family: Arial, Helvetica, sans-serif; }
    .callout-warning {
      border-left-color: var(--warning-text);
      background: var(--warning-bg);
    }
    .callout-info {
      border-left-color: #1565C0;
      background: #E3F2FD;
    }
    .meta-table {
      width: 100%; border-collapse: collapse;
      margin: 0 0 24px 0;
      font-family: Arial, Helvetica, sans-serif; font-size: 10pt;
    }
    .meta-table td {
      padding: 4px 10px;
      border: 1px solid var(--silver-400);
    }
    .meta-table td:first-child {
      background: var(--silver-100); font-weight: bold; width: 140px;
    }
    .footer {
      margin-top: 48px; padding: 12px 0;
      border-top: 1px solid var(--silver-400);
      font-family: Arial, Helvetica, sans-serif;
      font-size: 9pt; color: #888; text-align: center;
    }
    @media print {
      body { max-width: none; padding: 0; font-size: 10pt; }
      .cover { margin: 0 0 20px 0; }
      pre { white-space: pre-wrap; word-wrap: break-word; }
      h1 { page-break-before: always; }
      h1:first-of-type { page-break-before: avoid; }
      table { page-break-inside: avoid; }
    }
  </style>
</head>
<body>

<div class="cover">
  <h1>${escHtml(projectName)}</h1>
  <div class="subtitle">User Guide &mdash; ${escHtml(tag)}</div>
  <div class="meta">
    Version: ${escHtml(version)} &nbsp;|&nbsp; Repository: ${escHtml(repo)}
    &nbsp;|&nbsp; Generated: ${new Date().toISOString().slice(0, 10)}
  </div>
</div>

${sectionsHtml}

<div class="footer">
  ${escHtml(projectName)} User Guide &nbsp;|&nbsp;
  ${escHtml(tag)} &nbsp;|&nbsp;
  Frederick Barton &nbsp;|&nbsp;
  Generated ${new Date().toISOString().slice(0, 10)}
</div>
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
