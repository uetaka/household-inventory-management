/**
 * 初回セットアップ。1回だけ実行する。
 *   1. スプレッドシートと、それに紐付いた Apps Script プロジェクトを作成
 *   2. src/ をプッシュ
 *   3. Web アプリとしてデプロイし、URL を表示
 *
 * 前提: `npm install` と `npm run login` 済み。
 *       https://script.google.com/home/usersettings で Apps Script API を ON にしておくこと。
 */
import { execSync } from 'node:child_process';
import { existsSync, rmSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  ROOT, CLASP_JSON, clasp, readClaspJson, readDeploymentId, writeDeploymentId,
  extractDeploymentId, webAppUrl, spreadsheetUrl, assertLoggedIn,
} from './lib.mjs';

const title = process.argv[2] || '日用品在庫';

assertLoggedIn();

if (existsSync(CLASP_JSON)) {
  console.log('.clasp.json が既にあります。既存プロジェクトに対して push と deploy だけ行います。');
} else {
  clasp(`create-script --type sheets --title "${title}" --rootDir src`);

  // create 直後の pull で src/appsscript.json が既定のものに上書きされ、Code.js が増えるので元に戻す。
  execSync('git checkout -- src/appsscript.json', { cwd: ROOT, stdio: 'inherit' });
  const stray = resolve(ROOT, 'src', 'Code.js');
  if (existsSync(stray)) rmSync(stray);
}

clasp('push -f');

let deploymentId = readDeploymentId();
if (!deploymentId) {
  const out = clasp('create-deployment -d "初回デプロイ"', { capture: true });
  process.stdout.write(out);
  deploymentId = extractDeploymentId(out);
  if (!deploymentId) throw new Error('デプロイIDを取得できませんでした。`npx clasp list-deployments` で確認してください。');
  writeDeploymentId(deploymentId);
}

const cj = readClaspJson();
console.log('\n========================================');
console.log('セットアップ完了');
console.log('スプレッドシート: ' + spreadsheetUrl(cj));
console.log('Apps Script    : https://script.google.com/d/' + cj.scriptId + '/edit');
console.log('Web アプリ      : ' + webAppUrl(deploymentId));
console.log('========================================');
console.log('次にやること:');
console.log('  1. 上の Web アプリ URL をスマホ（または PC）で開き、初回の権限承認を済ませる');
console.log('  2. 画面の「初期設定」で Claude API キーを登録し、シートを作成する');
console.log('  3. .clasp.json と .deployment-id を git にコミットしておく');
