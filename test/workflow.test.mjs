import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, readFile, rm, mkdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import { capture } from '../scripts/capture.mjs';
import { makeWord } from '../scripts/make-word.mjs';

const repo = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const noSandbox = process.env.KBASE_TEST_NO_SANDBOX === '1';

test('сценарий → локальные PNG → Word с текстом, таблицей, ссылкой и изображениями', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'kbase-workflow-'));
  try {
    await writeFile(join(dir, 'fixture.js'), `
      window.renderFixture = async scene => {
        document.querySelector('#app').innerHTML = '<section class="card"><h1>Демонстрационный стенд теста</h1><p>' + scene + '</p><button>Открыть</button><div class="details" hidden>Подробности</div></section>';
        document.querySelector('button').onclick = () => document.querySelector('.details').hidden = false;
        fetch('https://example.invalid/blocked-test').catch(() => {});
      };
    `);
    await writeFile(join(dir, 'fixture.css'), '.card { width:420px;padding:24px;font:18px Arial;background:#eef8fb;color:#25313d; }');
    const config = { fixture: { script: 'fixture.js', styles: ['fixture.css'] },
      source: 'test fixture (not a product screenshot)', data: 'synthetic', settleMs: 0,
      shots: [
        { id: '01-start', scene: 'Начало', selector: '.card', caption: 'Начальный экран' },
        { id: '02-result', scene: 'Результат', selector: '.details', caption: 'Подробности результата',
          actions: [{ clickText: 'Открыть' }], viewport: { width: 720 } },
      ],
    };
    await writeFile(join(dir, 'capture.json'), JSON.stringify(config));
    const evidence = await capture(join(dir, 'capture.json'), { noSandbox });
    assert.equal(evidence.screenshots.length, 2);
    assert.ok(evidence.blockedRequests >= 2);
    assert.deepEqual(evidence.errors, []);
    for (const shot of evidence.screenshots) {
      const png = await readFile(join(dir, 'images', shot.file));
      assert.equal(png.subarray(1, 4).toString(), 'PNG');
      assert.ok(png.readUInt32BE(16) > 0);
      assert.equal(shot.viewport.deviceScaleFactor, 2);
    }
    await writeFile(join(dir, 'article.md'), `# Проверяем демонстрационный экран

## Открываем подробности

**Шаг 1.** Нажимаем *Открыть* и проверяем результат.

![Начальный экран](images/01-start.png)

Рисунок 1. Начинаем работу.

> 💡 Данные этого стенда вымышлены.

| Действие | Результат |
|---|---|
| Открыть | Подробности |

1. Выбираем запись.
2. Проверяем результат.

- Один пункт.
- Другой пункт.

Подробнее: [документация](https://example.com/docs).

![Подробности результата](images/02-result.png)
`);
    const docPath = join(dir, 'article.docx');
    const word = await makeWord(join(dir, 'article.md'), docPath);
    assert.equal(word.figures.length, 2);
    const report = JSON.parse(execFileSync('python3', [join(repo, 'scripts/validate-docx.py'), docPath, '--expect-images', '2'], { encoding: 'utf8' }));
    assert.equal(report.valid, true);
    assert.equal(report.tables, 1);
    assert.equal(report.paginationChecked, false);
    const inspect = JSON.parse(execFileSync('python3', ['-c', `
import json,sys,zipfile,xml.etree.ElementTree as E
with zipfile.ZipFile(sys.argv[1]) as z:
 r=E.fromstring(z.read('word/document.xml'))
 ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
 text=' '.join(n.text or '' for n in r.findall('.//w:t',ns))
 rel=E.fromstring(z.read('word/_rels/document.xml.rels'))
 print(json.dumps({'text':text,'bold':len(r.findall('.//w:b',ns)), 'links':[v.get('Target') for v in rel if v.get('TargetMode')=='External']}))
`, docPath], { encoding: 'utf8' }));
    assert.match(inspect.text.replace(/\s+/g, ' '), /Нажимаем Открыть и проверяем результат/);
    assert.ok(inspect.bold >= 1);
    assert.deepEqual(inspect.links, ['https://example.com/docs']);
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('ошибки исходников не превращаются в готовый документ или успешную съёмку', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'kbase-invalid-'));
  try {
    const article = join(dir, 'article.md'), output = join(dir, 'article.docx');
    await writeFile(article, '# Статья\n\n![Снимок](images/missing.png)\n');
    await assert.rejects(makeWord(article, output), /ENOENT/);
    await assert.rejects(readFile(output), /ENOENT/);
    await writeFile(article, '# Статья\n\n![Снимок](../private.png)\n');
    await assert.rejects(makeWord(article, output), /внутри папки статьи/);
    await writeFile(join(dir, 'capture.json'), JSON.stringify({
      source: 'test', data: 'synthetic', fixture: { script: 'unused.js' },
      shots: [{ id: '../escape', selector: '#app', caption: 'Ошибка' }],
    }));
    await assert.rejects(capture(join(dir, 'capture.json')), /уникальный id/);
    await mkdir(join(dir, 'broken'));
    const broken = join(dir, 'broken.docx');
    await writeFile(broken, 'not a DOCX');
    assert.throws(() => execFileSync('python3', [join(repo, 'scripts/validate-docx.py'), broken], { stdio: 'pipe' }));
  } finally { await rm(dir, { recursive: true, force: true }); }
});
