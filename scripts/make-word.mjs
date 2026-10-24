// Generalized from the Word exporter used for the illustrated legal-review article.
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { resolve, dirname, basename, relative, isAbsolute } from 'node:path';
import { pathToFileURL } from 'node:url';
import MarkdownIt from 'markdown-it';
import {
  Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, AlignmentType,
  Footer, Header, PageNumber, Table, TableRow, TableCell, WidthType,
  BorderStyle, ShadingType, ExternalHyperlink,
} from 'docx';

const ink = '25313D', accent = '087F9C', muted = '647482', line = 'D5DDE2';
const paragraph = (children, options = {}) => new Paragraph({ children,
  spacing: { after: 130, line: 290 }, widowControl: true, ...options });
const border = { style: BorderStyle.SINGLE, size: 4, color: line };

function inline(tokens) {
  const result = [];
  let bold = false, italics = false, strike = false, link = null, linked = [];
  for (const token of tokens ?? []) {
    if (token.type === 'strong_open') { bold = true; continue; }
    if (token.type === 'strong_close') { bold = false; continue; }
    if (token.type === 'em_open') { italics = true; continue; }
    if (token.type === 'em_close') { italics = false; continue; }
    if (token.type === 's_open') { strike = true; continue; }
    if (token.type === 's_close') { strike = false; continue; }
    if (token.type === 'link_open') { link = token.attrGet('href'); linked = []; continue; }
    if (token.type === 'link_close') {
      result.push(new ExternalHyperlink({ link, children: linked })); link = null; continue;
    }
    if (token.type === 'image') throw new Error('Поместите изображение в отдельный абзац Markdown.');
    const run = new TextRun({
      text: token.type === 'softbreak' ? ' ' : token.type === 'hardbreak' ? '' : token.content,
      break: token.type === 'hardbreak' ? 1 : undefined,
      bold, italics, strike,
      font: token.type === 'code_inline' ? 'Courier New' : undefined,
      style: link ? 'Hyperlink' : undefined,
    });
    (link ? linked : result).push(run);
  }
  return result;
}

function imagePath(base, src) {
  const file = resolve(base, decodeURIComponent(src));
  const rel = relative(base, file);
  if (isAbsolute(rel) || rel === '..' || rel.startsWith('../') || rel.startsWith('..\\')) {
    throw new Error(`Изображение должно находиться внутри папки статьи: ${src}`);
  }
  return file;
}

export async function makeWord(input, output, config = {}) {
  const sourcePath = resolve(input), base = dirname(sourcePath);
  const destination = resolve(output);
  const source = await readFile(sourcePath, 'utf8');
  const tokens = new MarkdownIt({ html: false }).parse(source, {});
  const blocks = [], figures = [], lists = [], numbering = [];
  let title = config.title, headings = 0, quoteDepth = 0, firstInItem = false;
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token.type === 'heading_open') {
      const content = tokens[++i];
      const depth = Number(token.tag.slice(1));
      title ??= content.content;
      const level = depth === 1 ? HeadingLevel.TITLE :
        [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3,
          HeadingLevel.HEADING_4, HeadingLevel.HEADING_5][Math.min(depth - 2, 4)];
      blocks.push(paragraph(inline(content.children), { heading: level, keepNext: true,
        spacing: { before: depth === 1 ? 120 : 240, after: 140 } }));
      headings++; continue;
    }
    if (token.type === 'blockquote_open') { quoteDepth++; continue; }
    if (token.type === 'blockquote_close') { quoteDepth--; continue; }
    if (token.type === 'bullet_list_open' || token.type === 'ordered_list_open') {
      const ordered = token.type === 'ordered_list_open';
      const reference = `list-${numbering.length}`;
      if (ordered) numbering.push({ reference, levels: Array.from({ length: 9 }, (_, level) => ({
        level, format: 'decimal', text: `%${level + 1}.`, alignment: AlignmentType.LEFT,
        start: Number(token.attrGet('start') ?? 1),
        style: { paragraph: { indent: { left: 360 * (level + 1), hanging: 260 } } },
      })) });
      lists.push({ ordered, reference }); continue;
    }
    if (token.type === 'bullet_list_close' || token.type === 'ordered_list_close') { lists.pop(); continue; }
    if (token.type === 'list_item_open') { firstInItem = true; continue; }
    if (token.type === 'list_item_close') { firstInItem = false; continue; }
    if (token.type === 'table_open') {
      const rows = [];
      let header = false;
      while (++i < tokens.length && tokens[i].type !== 'table_close') {
        const item = tokens[i];
        if (item.type === 'thead_open') header = true;
        if (item.type === 'thead_close') header = false;
        if (item.type !== 'tr_open') continue;
        const cells = [];
        while (++i < tokens.length && tokens[i].type !== 'tr_close') {
          if (tokens[i].type === 'inline') cells.push(tokens[i]);
        }
        rows.push(new TableRow({ tableHeader: header, cantSplit: true,
          children: cells.map(cell => new TableCell({
            width: { size: 100 / cells.length, type: WidthType.PERCENTAGE },
            margins: { top: 100, bottom: 100, left: 140, right: 140 },
            shading: header ? { type: ShadingType.CLEAR, fill: 'E8F3F7' } : undefined,
            children: [paragraph(inline(cell.children), { spacing: { after: 40, line: 275 } })],
          })),
        }));
      }
      blocks.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows,
        borders: { top: border, bottom: border, left: border, right: border,
          insideHorizontal: border, insideVertical: border } }));
      blocks.push(paragraph([])); continue;
    }
    if (token.type === 'fence' || token.type === 'code_block') {
      blocks.push(paragraph([new TextRun({ text: token.content.trimEnd(), font: 'Courier New', size: 20 })]));
      continue;
    }
    if (token.type === 'hr') {
      blocks.push(paragraph([], { border: { bottom: border } })); continue;
    }
    if (token.type !== 'paragraph_open') continue;
    const content = tokens[++i];
    const children = content.children ?? [];
    if (children.length === 1 && children[0].type === 'image') {
      const image = children[0], src = image.attrGet('src');
      const data = await readFile(imagePath(base, src));
      if (data.length < 24 || !data.subarray(0, 8).equals(Buffer.from([137,80,78,71,13,10,26,10]))) {
        throw new Error(`Нужен PNG: ${src}`);
      }
      const width = data.readUInt32BE(16), height = data.readUInt32BE(20);
      if (!width || !height) throw new Error(`Некорректный размер PNG: ${src}`);
      const limits = { maxWidth: 620, maxHeight: 520, ...config.figures, ...config.images?.[src] };
      if (!(limits.maxWidth > 0 && limits.maxWidth <= 650 && limits.maxHeight > 0 && limits.maxHeight <= 800)) {
        throw new Error(`Размер изображения выходит за границы A4: ${src}`);
      }
      const scale = Math.min(limits.maxWidth / width, limits.maxHeight / height, 1);
      const w = Math.max(1, Math.round(width * scale)), h = Math.max(1, Math.round(height * scale));
      blocks.push(paragraph([new ImageRun({ type: 'png', data,
        transformation: { width: w, height: h },
        altText: { title: image.content, description: image.content, name: basename(src) },
      })], { alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 100, after: 70 } }));
      figures.push({ file: src, pixels: [width, height], layoutPixels: [w, h], alt: image.content });
      continue;
    }
    const options = {};
    if (lists.length && firstInItem) {
      const list = lists.at(-1), level = Math.min(lists.length - 1, 8);
      if (list.ordered) options.numbering = { reference: list.reference, level };
      else options.bullet = { level };
      firstInItem = false;
    }
    if (quoteDepth) Object.assign(options, {
      shading: { type: ShadingType.CLEAR, fill: 'EFF8FB' },
      border: { left: { style: BorderStyle.SINGLE, color: '2FC6F6', size: 14, space: 9 } },
      indent: { left: 150, right: 120 }, keepLines: true,
    });
    if (/^Рисунок\s+\d/.test(content.content)) Object.assign(options, {
      alignment: AlignmentType.CENTER, style: 'Caption', keepLines: true,
    });
    blocks.push(paragraph(inline(children), options));
  }
  const header = config.header ?? title ?? 'База знаний';
  const doc = new Document({ title: title ?? 'Инструкция', creator: config.creator ?? 'База знаний',
    subject: config.subject ?? 'Пользовательская инструкция', numbering: { config: numbering },
    styles: {
      default: { document: { run: { font: 'Arial', size: 23, color: ink } } },
      paragraphStyles: [
        { id: 'Title', name: 'Title', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { font: 'Arial', size: 45, bold: true, color: ink } },
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { font: 'Arial', size: 32, bold: true, color: accent } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { font: 'Arial', size: 26, bold: true, color: ink } },
        { id: 'Caption', name: 'Caption', basedOn: 'Normal',
          run: { italics: true, size: 19, color: muted } },
      ],
    },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 },
      margin: { top: 1080, bottom: 1080, left: 1020, right: 1020, header: 500, footer: 500 } } },
      headers: { default: new Header({ children: [paragraph([new TextRun({ text: header, color: muted, size: 18 })],
        { border: { bottom: border } })] }) },
      footers: { default: new Footer({ children: [paragraph([
        new TextRun({ text: 'Страница ', size: 18, color: muted }),
        new TextRun({ children: [PageNumber.CURRENT], size: 18, color: muted }),
        new TextRun({ text: ' из ', size: 18, color: muted }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: muted }),
      ], { alignment: AlignmentType.RIGHT })] }) }, children: blocks,
    }],
  });
  await mkdir(dirname(destination), { recursive: true });
  await writeFile(destination, await Packer.toBuffer(doc));
  const evidence = { headings, figures, format: 'A4 portrait', font: 'Arial 11.5pt',
    editableText: true, embeddedImages: true, destination: basename(destination) };
  await writeFile(`${destination}.evidence.json`, JSON.stringify(evidence, null, 2) + '\n');
  return evidence;
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const [input, output, configPath, ...extra] = process.argv.slice(2);
  if (!input || !output || extra.length) {
    console.error('Использование: node scripts/make-word.mjs article.md article.docx [word.config.json]');
    process.exitCode = 1;
  } else {
    (async () => {
      const config = configPath ? JSON.parse(await readFile(resolve(configPath), 'utf8')) : {};
      const evidence = await makeWord(input, output, config);
      console.log(`Создан ${output}; встроено изображений: ${evidence.figures.length}.`);
    })().catch(error => { console.error(error.message); process.exitCode = 1; });
  }
}
