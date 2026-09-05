/**
 * コード変更を反映する。push して既存デプロイを新バージョンに更新する（URL は変わらない）。
 *   npm run deploy            push + 更新
 *   npm run open:app          Web アプリをブラウザで開くだけ
 */
import {
  clasp, readDeploymentId, writeDeploymentId, extractDeploymentId, webAppUrl, assertLoggedIn,
} from './lib.mjs';

const openOnly = process.argv.includes('--open');
let deploymentId = readDeploymentId();

if (openOnly) {
  if (!deploymentId) throw new Error('.deployment-id がありません。先に `npm run bootstrap` を実行してください。');
  clasp(`open-web-app ${deploymentId}`);
  process.exit(0);
}

assertLoggedIn();
clasp('push -f');

const desc = new Date().toISOString().slice(0, 16).replace('T', ' ');
if (deploymentId) {
  clasp(`update-deployment ${deploymentId} -d "${desc}"`);
} else {
  const out = clasp(`create-deployment -d "${desc}"`, { capture: true });
  process.stdout.write(out);
  deploymentId = extractDeploymentId(out);
  if (deploymentId) writeDeploymentId(deploymentId);
}

console.log('\nWeb アプリ: ' + webAppUrl(deploymentId));
