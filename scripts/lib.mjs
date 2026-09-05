import { execSync, spawnSync } from 'node:child_process';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';

export const ROOT = resolve(new URL('..', import.meta.url).pathname);
export const CLASP_JSON = resolve(ROOT, '.clasp.json');
export const DEPLOYMENT_FILE = resolve(ROOT, '.deployment-id');

/** clasp を実行して出力を返す。失敗時は例外。 */
export function clasp(args, { capture = false } = {}) {
  const cmd = `npx clasp ${args}`;
  console.log(`\n$ ${cmd}`);
  if (capture) {
    return execSync(cmd, { cwd: ROOT, encoding: 'utf8', stdio: ['inherit', 'pipe', 'inherit'] });
  }
  const r = spawnSync(cmd, { cwd: ROOT, stdio: 'inherit', shell: true });
  if (r.status !== 0) throw new Error(`clasp が失敗しました: ${cmd}`);
  return '';
}

export function readClaspJson() {
  if (!existsSync(CLASP_JSON)) return null;
  return JSON.parse(readFileSync(CLASP_JSON, 'utf8'));
}

export function readDeploymentId() {
  return existsSync(DEPLOYMENT_FILE) ? readFileSync(DEPLOYMENT_FILE, 'utf8').trim() : '';
}

export function writeDeploymentId(id) {
  writeFileSync(DEPLOYMENT_FILE, id + '\n');
}

/** clasp の出力からデプロイIDを拾う（"Deployed AKfycb... @3" 形式）。 */
export function extractDeploymentId(output) {
  const m = output.match(/AKfycb[\w-]+/);
  return m ? m[0] : '';
}

export function webAppUrl(deploymentId) {
  return `https://script.google.com/macros/s/${deploymentId}/exec`;
}

export function spreadsheetUrl(claspJson) {
  const parent = claspJson && claspJson.parentId;
  const id = Array.isArray(parent) ? parent[0] : parent;
  return id ? `https://docs.google.com/spreadsheets/d/${id}/edit` : '(不明: clasp open-container で開けます)';
}

export function assertLoggedIn() {
  const r = spawnSync('npx clasp show-authorized-user', { cwd: ROOT, shell: true, encoding: 'utf8' });
  if (r.status !== 0 || /not logged in|No credentials/i.test(r.stdout + r.stderr)) {
    throw new Error('clasp にログインしていません。先に `npm run login` を実行してください。');
  }
}
