// Reusable adaptation of the article capture harness: real components, synthetic data.
import { readFile, writeFile, mkdir, rm } from 'node:fs/promises';
import { resolve, dirname, join } from 'node:path';
import { pathToFileURL } from 'node:url';
import puppeteer from 'puppeteer';

export async function capture(configPath, { noSandbox = false } = {}) {
  const file = resolve(configPath);
  const base = dirname(file);
  const config = JSON.parse(await readFile(file, 'utf8'));
  if (!config.fixture?.script || !config.source || !config.data || !config.shots?.length) {
    throw new Error('Нужны fixture.script, source, data и непустой shots. См. references/capture.md.');
  }
  const ids = new Set();
  for (const shot of config.shots) {
    if (!/^[a-zA-Z0-9_-]+$/.test(shot.id) || ids.has(shot.id) || !shot.selector || !shot.caption) {
      throw new Error('Каждый кадр требует уникальный id (буквы, цифры, _ или -), selector и caption.');
    }
    ids.add(shot.id);
  }
  const output = resolve(base, config.output ?? 'images');
  const evidencePath = resolve(base, config.evidence ?? 'capture-evidence.json');
  const adapter = config.adapter ? await import(pathToFileURL(resolve(base, config.adapter)).href) : null;
  if (adapter && typeof adapter.prepareShot !== 'function') throw new Error('Адаптер должен экспортировать prepareShot.');
  const script = await readFile(resolve(base, config.fixture.script), 'utf8');
  const styles = await Promise.all((config.fixture.styles ?? []).map(p => readFile(resolve(base, p), 'utf8')));
  await rm(evidencePath, { force: true });
  const browser = await puppeteer.launch({
    headless: true, pipe: true,
    args: noSandbox ? ['--no-sandbox', '--disable-setuid-sandbox'] : [],
  });
  try {
    const page = await browser.newPage();
    await page.setOfflineMode(true);
    page.setDefaultTimeout(10000);
    const errors = [];
    let blockedRequests = 0;
    page.on('pageerror', error => errors.push(error.message));
    await page.setRequestInterception(true);
    page.on('request', request => {
      if (/^(data:|about:|blob:)/.test(request.url())) request.continue();
      else { blockedRequests++; request.abort(); }
    });
    await page.setViewport(config.viewport ?? { width: 1180, height: 1040, deviceScaleFactor: 2 });
    await page.setContent('<!doctype html><html lang="ru"><meta charset="utf-8"><body><div id="app"></div></body></html>');
    for (const content of styles) await page.addStyleTag({ content });
    await page.addScriptTag({ content: script });
    await mkdir(output, { recursive: true });
    const screenshots = [];
    for (const shot of config.shots) {
      await page.setViewport({ ...(config.viewport ?? { width: 1180, height: 1040, deviceScaleFactor: 2 }), ...shot.viewport });
      if (shot.scene !== undefined) {
        await page.evaluate(async ({ hook, scene }) => {
          if (typeof window[hook] !== 'function') throw new Error(`Нет функции сцены ${hook}`);
          await window[hook](scene);
        }, { hook: config.fixture.sceneHook ?? 'renderFixture', scene: shot.scene });
      }
      if (adapter) await adapter.prepareShot({ page, shot });
      for (const action of shot.actions ?? []) {
        if (action.click) await page.locator(action.click).click();
        else if (action.clickText) {
          await page.evaluate(text => {
            const button = [...document.querySelectorAll('button')].find(b => b.textContent.trim() === text);
            if (!button) throw new Error(`Не найдена кнопка: ${text}`);
            button.click();
          }, action.clickText);
        } else throw new Error('Действие поддерживает click или clickText. Для остальных действий нужен adapter.');
      }
      await page.evaluate(() => document.fonts.ready);
      await page.waitForSelector(shot.selector, { visible: true });
      const delay = shot.settleMs ?? config.settleMs ?? 450;
      if (!Number.isFinite(delay) || delay < 0 || delay > 10000) throw new Error('settleMs должен быть от 0 до 10000.');
      await new Promise(done => setTimeout(done, delay));
      const target = await page.$(shot.selector);
      const box = await target.boundingBox();
      if (!box?.width || !box?.height) throw new Error(`Пустая область: ${shot.selector}`);
      await target.screenshot({ path: join(output, `${shot.id}.png`) });
      screenshots.push({ file: `${shot.id}.png`, scene: shot.scene, selector: shot.selector,
        caption: shot.caption, viewport: page.viewport(), bounds: box });
      if (errors.length) throw new Error(`Ошибка интерфейса: ${errors.join('; ')}`);
    }
    const evidence = { offline: true, network: 'blocked', blockedRequests,
      source: config.source, data: config.data, screenshots, errors };
    await mkdir(dirname(evidencePath), { recursive: true });
    await writeFile(evidencePath, JSON.stringify(evidence, null, 2) + '\n');
    return evidence;
  } finally {
    await browser.close();
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const args = process.argv.slice(2);
  const path = args.find(arg => !arg.startsWith('--'));
  if (!path || args.some(arg => arg.startsWith('--') && arg !== '--no-sandbox')) {
    console.error('Использование: node scripts/capture.mjs capture.json [--no-sandbox]');
    process.exitCode = 1;
  } else {
    capture(path, { noSandbox: args.includes('--no-sandbox') })
      .then(result => console.log(`Снято ${result.screenshots.length} кадров; сетевые запросы заблокированы.`))
      .catch(error => { console.error(error.message); process.exitCode = 1; });
  }
}
