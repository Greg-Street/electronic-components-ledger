import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const distRoot = path.join(projectRoot, 'dist');
const distIndex = path.join(distRoot, 'index.html');

const html = await fs.readFile(distIndex, 'utf8');
const scriptMatch = html.match(/<script[^>]+src="([^"]+)"[^>]*><\/script>/i);
const styleMatch = html.match(/<link[^>]+href="([^"]+\.css)"[^>]*>/i);

if (!scriptMatch || !styleMatch) {
  throw new Error('无法从 Vite 构建结果中找到 JS 或 CSS 资源。');
}

const scriptPath = path.join(distRoot, scriptMatch[1].replace(/^\//, '').replaceAll('/', path.sep));
const stylePath = path.join(distRoot, styleMatch[1].replace(/^\//, '').replaceAll('/', path.sep));
const [script, style] = await Promise.all([
  fs.readFile(scriptPath, 'utf8'),
  fs.readFile(stylePath, 'utf8'),
]);

const standalone = html
  .replace(scriptMatch[0], () => `<script type="module">${script.replaceAll('</script>', '<\\/script>')}</script>`)
  .replace(styleMatch[0], () => `<style>${style}</style>`);

await Promise.all([
  fs.writeFile(path.join(projectRoot, '电子元件库存台账.html'), standalone, 'utf8'),
  fs.writeFile(path.join(distRoot, '电子元件库存台账.html'), standalone, 'utf8'),
]);

console.log('已生成单文件应用：电子元件库存台账.html');
