/**
 * =============================================================================
 *  Generate-Docs.js — Word document generator from script headers
 *
 *  Author        : Frederick Barton
 *  Version       : 1.0.0
 *  Last Updated  : 2026-03-21
 *  Environment   : Node.js 18+ / GitHub Actions
 *
 *  Change Log    :
 *    1.0.0 - 2026-03-21
 *        - Initial release
 *        - Parses PowerShell, Python, and JavaScript script headers
 *        - Generates formatted Word document in docs/ folder
 *        - Updates existing doc if already present for this script
 *
 * -----------------------------------------------------------------------------
 *  PURPOSE:
 *    Triggered by GitHub Actions on minor/major version tag pushes.
 *    Finds all scripts in the repo, parses their structured headers,
 *    and generates (or updates) a Word document per script in docs/.
 *
 *  USAGE:
 *    node Generate-Docs.js --version 1.2 --repo owner/repo --tag v1.2
 *
 * =============================================================================
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Footer, Header, TabStopType, TabStopPosition,
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
const VERSION = args.version || '1.0';
const REPO    = args.repo    || '';
const TAG     = args.tag     || `v${VERSION}`;

// ---------------------------------------------------------------------------
// Header section labels (order matters — used for display and parsing)
// ---------------------------------------------------------------------------
const PS_SECTIONS = [
  'PURPOSE', 'CASE / PROJECT / TASK', 'ISSUE', 'CAUSE', 'SOLUTION',
  'SCRIPT README (Quick Start)', 'OTHER INFORMATION',
  'GOALS', 'MAJOR LOGIC CHOICES', 'CONFIGURATION OPTIONS',
  'OUTPUTS', 'NOTES',
];
const PY_SECTIONS = ['PURPOSE', 'USAGE', 'DEPENDENCIES', 'NOTES'];
const JS_SECTIONS = [
  'PURPOSE', 'CASE / PROJECT / TASK', 'ISSUE', 'CAUSE', 'SOLUTION',
  'USAGE / ENTRY POINT', 'DEPENDENCIES', 'INPUTS', 'OUTPUTS', 'NOTES',
];

// ---------------------------------------------------------------------------
// Find all scripts in repo (exclude node_modules, .git, docs)
// ---------------------------------------------------------------------------
function findScripts(root) {
  const results = [];
  const exts    = ['.ps1', '.py', '.js'];
  const skip    = new Set(['node_modules', '.git', 'docs', '.github']);

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
// Parse a script header into a structured object
// ---------------------------------------------------------------------------
function parseHeader(filePath) {
  const ext     = path.extname(filePath).toLowerCase();
  const content = fs.readFileSync(filePath, 'utf8');
  const lines   = content.split('\n');

  const info = {
    filePath,
    fileName:    path.basename(filePath),
    ext,
    title:       '',
    description: '',
    author:      '',
    version:     '',
    lastUpdated: '',
    validatedOn: '',
    environment: '',
    changeLog:   [],
    sections:    {},
  };

  // Determine comment style and section list
  let inHeader  = false;
  let inChanges = false;
  let currentSection = null;
  let sectionLines   = [];
  let changeEntry    = null;
  const sectionList  = ext === '.ps1' ? PS_SECTIONS
                     : ext === '.py'  ? PY_SECTIONS
                     : JS_SECTIONS;

  for (let i = 0; i < lines.length; i++) {
    const raw  = lines[i];
    // Strip comment prefix for each language
    let line = raw;
    if (ext === '.ps1') line = raw.replace(/^\s*/, '');
    if (ext === '.py')  line = raw.replace(/^#\s?/, '');
    if (ext === '.js')  line = raw.replace(/^\s*\*\s?/, '').replace(/^\s*\/\*+/, '').replace(/\*+\/\s*$/, '');

    // Detect header start
    if (!inHeader) {
      if (line.includes('='.repeat(20))) { inHeader = true; continue; }
      continue;
    }
    // Detect header end
    if (line.includes('='.repeat(20)) && i > 5) break;

    // Metadata fields
    const metaMatch = line.match(/^\s*([\w\s\/]+?)\s*:\s*(.+)/);
    if (metaMatch && !currentSection) {
      const key = metaMatch[1].trim();
      const val = metaMatch[2].trim();
      if      (key === 'Author')       info.author      = val;
      else if (key === 'Version')      info.version     = val;
      else if (key === 'Last Updated') info.lastUpdated = val;
      else if (key === 'Validated on') info.validatedOn = val;
      else if (key === 'Environment')  info.environment = val;
      // Script title from first line after ===
      continue;
    }

    // Title line (e.g. " ScriptName.ps1 — Description")
    if (!info.title && line.match(/\s*\S+\.(ps1|py|js)\s*[—-]/i)) {
      const parts = line.split(/[—-]/);
      info.fileName   = parts[0].trim();
      info.description = parts.slice(1).join('—').trim();
      continue;
    }

    // Change Log entries
    if (line.trim() === 'Change Log    :' || line.trim() === 'Change Log:') {
      inChanges = true; continue;
    }
    if (inChanges && !currentSection) {
      const versionLine = line.match(/^\s*([\d.]+)\s*-\s*(\d{4}-\d{2}-\d{2})/);
      if (versionLine) {
        if (changeEntry) info.changeLog.push(changeEntry);
        changeEntry = { version: versionLine[1], date: versionLine[2], items: [] };
        continue;
      }
      if (changeEntry && line.trim().startsWith('-')) {
        changeEntry.items.push(line.replace(/^\s*-\s*/, '').trim());
        continue;
      }
      if (changeEntry && line.trim().startsWith('*')) {
        // Sub-bullet — append to last item
        if (changeEntry.items.length) {
          changeEntry.items[changeEntry.items.length - 1] += ' ' + line.replace(/^\s*\*\s*/, '').trim();
        }
        continue;
      }
    }

    // Section headers (e.g. " PURPOSE:", " GOALS:")
    const sectionMatch = sectionList.find(s => line.trim().toUpperCase().startsWith(s.toUpperCase() + ':')
                                            || line.trim().toUpperCase() === s.toUpperCase() + ':');
    if (sectionMatch) {
      if (currentSection) info.sections[currentSection] = sectionLines.join('\n').trim();
      // Save any pending changeLog entry when first section starts
      if (changeEntry) { info.changeLog.push(changeEntry); changeEntry = null; inChanges = false; }
      currentSection = sectionMatch;
      sectionLines   = [];
      continue;
    }

    // Divider lines — skip
    if (line.trim().startsWith('-'.repeat(10))) continue;

    // Accumulate section content
    if (currentSection) sectionLines.push(line.replace(/^\s{1,3}/, ''));
  }

  // Flush last section
  if (currentSection) info.sections[currentSection] = sectionLines.join('\n').trim();
  if (changeEntry)    info.changeLog.push(changeEntry);

  return info;
}

// ---------------------------------------------------------------------------
// Shared styles
// ---------------------------------------------------------------------------
const BLUE       = '2E75B6';
const BLUE_LIGHT = 'D5E8F0';
const GRAY_LIGHT = 'F5F5F5';
const border     = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders    = { top: border, bottom: border, left: border, right: border };
const noBorder   = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders  = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, size: 32, font: 'Arial', color: BLUE })],
    spacing: { before: 360, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 1 } },
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, size: 26, font: 'Arial', color: '404040' })],
    spacing: { before: 240, after: 80 },
  });
}

function bodyText(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, font: 'Arial', ...opts })],
    spacing: { after: 80 },
  });
}

function bulletPara(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'bullets', level },
    children: [new TextRun({ text, size: 22, font: 'Arial' })],
    spacing: { after: 60 },
  });
}

function sectionContent(text) {
  if (!text) return [bodyText('—')];
  const paras = [];
  for (const line of text.split('\n')) {
    const stripped = line.trimEnd();
    if (!stripped) { paras.push(new Paragraph({ spacing: { after: 40 } })); continue; }
    const indent = line.match(/^\s*/)[0].length;
    if (stripped.trimStart().startsWith('* ')) {
      paras.push(bulletPara(stripped.trimStart().slice(2), 1));
    } else if (stripped.trimStart().startsWith('- ') || stripped.trimStart().match(/^\d+\)/)) {
      paras.push(bulletPara(stripped.trimStart().replace(/^[-\d]+[.)]\s*/, ''), indent > 2 ? 1 : 0));
    } else {
      paras.push(bodyText(stripped.trimStart()));
    }
  }
  return paras;
}

// ---------------------------------------------------------------------------
// Change log table
// ---------------------------------------------------------------------------
function changeLogTable(changeLog) {
  if (!changeLog.length) return [];

  const headerRow = new TableRow({
    children: [
      new TableCell({
        borders, width: { size: 1440, type: WidthType.DXA },
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: 'Version', bold: true, color: 'FFFFFF', size: 22, font: 'Arial' })] })],
      }),
      new TableCell({
        borders, width: { size: 1440, type: WidthType.DXA },
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: 'Date', bold: true, color: 'FFFFFF', size: 22, font: 'Arial' })] })],
      }),
      new TableCell({
        borders, width: { size: 6480, type: WidthType.DXA },
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: 'Changes', bold: true, color: 'FFFFFF', size: 22, font: 'Arial' })] })],
      }),
    ],
  });

  const dataRows = changeLog.map((entry, idx) =>
    new TableRow({
      children: [
        new TableCell({
          borders, width: { size: 1440, type: WidthType.DXA },
          shading: { fill: idx % 2 === 0 ? GRAY_LIGHT : 'FFFFFF', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: entry.version, size: 22, font: 'Arial' })] })],
        }),
        new TableCell({
          borders, width: { size: 1440, type: WidthType.DXA },
          shading: { fill: idx % 2 === 0 ? GRAY_LIGHT : 'FFFFFF', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: entry.date, size: 22, font: 'Arial' })] })],
        }),
        new TableCell({
          borders, width: { size: 6480, type: WidthType.DXA },
          shading: { fill: idx % 2 === 0 ? GRAY_LIGHT : 'FFFFFF', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: entry.items.map(item => new Paragraph({
            numbering: { reference: 'bullets', level: 0 },
            children: [new TextRun({ text: item, size: 22, font: 'Arial' })],
            spacing: { after: 40 },
          })),
        }),
      ],
    })
  );

  return [
    heading2('Change Log'),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [1440, 1440, 6480],
      rows: [headerRow, ...dataRows],
    }),
    new Paragraph({ spacing: { after: 200 } }),
  ];
}

// ---------------------------------------------------------------------------
// Cover page
// ---------------------------------------------------------------------------
function coverPage(info, repoName) {
  return [
    new Paragraph({ spacing: { before: 1440, after: 80 } }), // top padding
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: info.fileName, bold: true, size: 56, font: 'Arial', color: BLUE })],
      spacing: { after: 160 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: info.description || 'Script Documentation', size: 32, font: 'Arial', color: '606060' })],
      spacing: { after: 480 },
    }),
    // Metadata table
    new Table({
      width: { size: 6480, type: WidthType.DXA },
      columnWidths: [2160, 4320],
      rows: [
        ['Author',      info.author      || 'Frederick Barton'],
        ['Version',     info.version     || VERSION],
        ['Last Updated', info.lastUpdated || new Date().toISOString().slice(0, 10)],
        ['Repository',  repoName         || ''],
        ['Tag',         TAG],
        [info.ext === '.ps1' ? 'Validated on' : 'Environment',
         info.validatedOn || info.environment || ''],
      ].filter(r => r[1]).map(([label, value], idx) =>
        new TableRow({
          children: [
            new TableCell({
              borders, width: { size: 2160, type: WidthType.DXA },
              shading: { fill: BLUE_LIGHT, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 22, font: 'Arial', color: '404040' })] })],
            }),
            new TableCell({
              borders, width: { size: 4320, type: WidthType.DXA },
              shading: { fill: idx % 2 === 0 ? GRAY_LIGHT : 'FFFFFF', type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: value, size: 22, font: 'Arial' })] })],
            }),
          ],
        })
      ),
    }),
    new Paragraph({ spacing: { after: 200 } }),
  ];
}

// ---------------------------------------------------------------------------
// Build full document for one script
// ---------------------------------------------------------------------------
function buildDocument(info, repoName) {
  const sectionList = info.ext === '.ps1' ? PS_SECTIONS
                    : info.ext === '.py'  ? PY_SECTIONS
                    : JS_SECTIONS;

  // Friendly display labels
  const sectionLabels = {
    'PURPOSE':                   'Purpose',
    'CASE / PROJECT / TASK':     'Case / Project / Task',
    'ISSUE':                     'Issue',
    'CAUSE':                     'Cause',
    'SOLUTION':                  'Solution',
    'SCRIPT README (Quick Start)':'Quick Start Guide',
    'OTHER INFORMATION':         'Other Information',
    'GOALS':                     'Goals',
    'MAJOR LOGIC CHOICES':       'Major Logic Choices',
    'CONFIGURATION OPTIONS':     'Configuration Options',
    'OUTPUTS':                   'Outputs',
    'NOTES':                     'Notes',
    'USAGE':                     'Usage',
    'USAGE / ENTRY POINT':       'Usage / Entry Point',
    'DEPENDENCIES':              'Dependencies',
    'INPUTS':                    'Inputs',
  };

  const children = [
    ...coverPage(info, repoName),
    // Page break before main content
    new Paragraph({ children: [new TextRun({ break: 1 })], pageBreakBefore: true }),
    heading1('Overview'),
    bodyText(info.description || ''),
    new Paragraph({ spacing: { after: 160 } }),
    ...changeLogTable(info.changeLog),
    heading1('Documentation'),
    ...sectionList.flatMap(s => {
      const content = info.sections[s];
      if (!content && s === 'CASE / PROJECT / TASK') return [];
      return [
        heading2(sectionLabels[s] || s),
        ...sectionContent(content || ''),
        new Paragraph({ spacing: { after: 120 } }),
      ];
    }),
  ];

  return new Document({
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [
            { level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
            { level: 1, format: LevelFormat.BULLET, text: '\u25E6', alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
          ],
        },
      ],
    },
    styles: {
      default: { document: { run: { font: 'Arial', size: 22 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 32, bold: true, font: 'Arial', color: BLUE },
          paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 26, bold: true, font: 'Arial', color: '404040' },
          paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 } },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: `${info.fileName}  `, size: 18, font: 'Arial', color: '808080' }),
                new TextRun({ text: `v${info.version || VERSION}`, size: 18, font: 'Arial', color: BLUE }),
              ],
              border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE_LIGHT, space: 1 } },
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: 'Frederick Barton  |  ', size: 18, font: 'Arial', color: '808080' }),
                new TextRun({ text: `Generated ${new Date().toISOString().slice(0, 10)}`, size: 18, font: 'Arial', color: '808080' }),
                new TextRun({ children: [new PageNumber()], size: 18, font: 'Arial', color: '808080' }),
              ],
              alignment: AlignmentType.RIGHT,
              border: { top: { style: BorderStyle.SINGLE, size: 4, color: BLUE_LIGHT, space: 1 } },
            }),
          ],
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
  const repoRoot = process.cwd();
  const docsDir  = path.join(repoRoot, 'docs');
  if (!fs.existsSync(docsDir)) fs.mkdirSync(docsDir, { recursive: true });

  const scripts = findScripts(repoRoot);
  console.log(`Found ${scripts.length} script(s)`);

  for (const scriptPath of scripts) {
    try {
      const info    = parseHeader(scriptPath);
      const docName = path.basename(scriptPath, path.extname(scriptPath)) + '.docx';
      const docPath = path.join(docsDir, docName);

      const doc    = buildDocument(info, REPO);
      const buffer = await Packer.toBuffer(doc);
      fs.writeFileSync(docPath, buffer);

      const action = fs.existsSync(docPath) ? 'Updated' : 'Created';
      console.log(`${action}: ${docPath}`);
    } catch (err) {
      console.error(`Error processing ${scriptPath}: ${err.message}`);
    }
  }
})();
