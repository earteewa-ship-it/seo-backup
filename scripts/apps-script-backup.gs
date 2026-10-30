/**
 * Ежедневный бэкап SEO-таблиц в GitHub (earteewa-ship-it/seo-backup).
 *
 * Что делает: для каждой таблицы из TABLES кладёт в репозиторий
 *   sheets/<Название>.xlsx          — полная копия с форматированием
 *   sheets/<Название>/<Лист>.csv    — данные каждого листа (видно изменения построчно)
 * Одна таблица = один коммит. Если за 4,5 минуты не успел (лимит Apps Script — 6 мин),
 * сам продолжает через минуту.
 *
 * Установка — см. README.md, раздел «Как включить ежедневный бэкап таблиц».
 * Токен GitHub хранится в Свойствах скрипта (GITHUB_TOKEN), в коде его нет.
 */

const GH_OWNER  = 'earteewa-ship-it';
const GH_REPO   = 'seo-backup';
const GH_BRANCH = 'main';

const TABLES = [
  { name: 'Выгрузка Miraterri', id: '1NAVmTYSgS1IGkvR9dr37iej4Ax4alRcuzmbvtlS6xWo' },
  { name: 'Выгрузка ECOSTYLE',  id: '1WMaSoPRTUY2JRpukEZbV0jt-50Kt5t2D59Qiw9FuJU0' },
  { name: 'Выгрузка Caprifash', id: '1KJ0_cDN8Vg81CYYU7Qde7vgT1sJ0lCnhoBb_cN0yit4' },
  { name: 'Моноподы',           id: '1tDCXqT5PToiUU6eYvZCc2_QqTQJ10PVAZiDcvBqL5Wo' },
  { name: 'Выгрузка Спорт',     id: '1NWf1Fxspve3rW5YwHmvuU8jJIxBt_BewdYOcy4dDxsk' },
];

const TIME_BUDGET_MS = 4.5 * 60 * 1000;

/** Запустить один раз вручную: ставит ежедневный запуск около 21:00 по часовому поясу скрипта. */
function installDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['dailyBackup', 'continueBackup'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyBackup').timeBased().everyDays(1).atHour(21).nearMinute(30).create();
  Logger.log('Ежедневный бэкап включён (около 21:30).');
}

/** Точка входа по расписанию (можно запустить и вручную для проверки). */
function dailyBackup() {
  PropertiesService.getScriptProperties().setProperty('NEXT_TABLE', '0');
  continueBackup();
}

function continueBackup() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'continueBackup')
    .forEach(t => ScriptApp.deleteTrigger(t));

  const props = PropertiesService.getScriptProperties();
  const started = Date.now();
  let i = Number(props.getProperty('NEXT_TABLE') || 0);
  const errors = [];

  while (i < TABLES.length && Date.now() - started < TIME_BUDGET_MS) {
    try {
      backupTable_(TABLES[i]);
    } catch (e) {
      errors.push(TABLES[i].name + ': ' + e.message);
      Logger.log('ОШИБКА ' + TABLES[i].name + ': ' + e.stack);
    }
    i++;
    props.setProperty('NEXT_TABLE', String(i));
  }

  if (i < TABLES.length) {
    ScriptApp.newTrigger('continueBackup').timeBased().after(60 * 1000).create();
  }
  if (errors.length) {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
      'SEO-бэкап: ошибки', errors.join('\n'));
  }
}

function backupTable_(t) {
  const files = {};
  const token = ScriptApp.getOAuthToken();
  const base = 'https://docs.google.com/spreadsheets/d/' + t.id + '/export?';

  files['sheets/' + safe_(t.name) + '.xlsx'] =
    fetchExport_(base + 'format=xlsx', token).getBytes();

  const ss = SpreadsheetApp.openById(t.id);
  ss.getSheets().forEach(sh => {
    const bytes = fetchExport_(base + 'format=csv&gid=' + sh.getSheetId(), token).getBytes();
    files['sheets/' + safe_(t.name) + '/' + safe_(sh.getName()) + '.csv'] = bytes;
    Utilities.sleep(300);
  });

  commit_(files, 'Бэкап «' + t.name + '» ' +
    Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm'));
}

function fetchExport_(url, token) {
  for (let attempt = 1; attempt <= 5; attempt++) {
    const r = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) return r.getBlob();
    Utilities.sleep(2000 * attempt);   // Google иногда отвечает 429 на частые выгрузки
  }
  throw new Error('не удалось выгрузить ' + url);
}

function safe_(s) {
  return String(s).replace(/[\\/:*?"<>|]/g, '_').trim() || 'sheet';
}

// ---------- GitHub (Git Data API): все файлы таблицы одним коммитом ----------

function gh_(method, path, body) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('Не задан GITHUB_TOKEN в Свойствах скрипта');
  const r = UrlFetchApp.fetch('https://api.github.com/repos/' + GH_OWNER + '/' + GH_REPO + path, {
    method: method,
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
    payload: body ? JSON.stringify(body) : undefined,
    muteHttpExceptions: true,
  });
  const code = r.getResponseCode();
  if (code >= 300) throw new Error('GitHub ' + method + ' ' + path + ' → ' + code + ': ' + r.getContentText().slice(0, 300));
  return JSON.parse(r.getContentText());
}

function commit_(files, message) {
  const ref = gh_('get', '/git/ref/heads/' + GH_BRANCH);
  const parentSha = ref.object.sha;
  const baseTree = gh_('get', '/git/commits/' + parentSha).tree.sha;

  const tree = Object.keys(files).map(path => ({
    path: path,
    mode: '100644',
    type: 'blob',
    sha: gh_('post', '/git/blobs', {
      content: Utilities.base64Encode(files[path]), encoding: 'base64' }).sha,
  }));

  const newTree = gh_('post', '/git/trees', { base_tree: baseTree, tree: tree });
  if (newTree.sha === baseTree) { Logger.log(message + ': изменений нет'); return; }

  const c = gh_('post', '/git/commits', { message: message, tree: newTree.sha, parents: [parentSha] });
  gh_('patch', '/git/refs/heads/' + GH_BRANCH, { sha: c.sha });
  Logger.log(message + ': ' + Object.keys(files).length + ' файлов');
}
