import http from "node:http";
import fs from "node:fs/promises";
import fsSync from "node:fs";
import path from "node:path";
import process from "node:process";
import { spawn } from "node:child_process";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { buildFileIndex, searchFileIndex } from "./file-index.mjs";
import { decideSchedule, markScheduleDue, normalizeScheduleState, setScheduleEnabled } from "./schedule-state.mjs";

const outputDir = "/Users/emily/Documents/jconedrive";
const scriptPath = path.join(outputDir, "download-teams-files.mjs");
const plistPath = "/Users/emily/Library/LaunchAgents/com.emily.jconedrive.teams-sync.plist";
const lockPath = path.join(outputDir, ".teams-sync.lockdir");
const manifestPath = path.join(outputDir, ".teams-sync-manifest.json");
const scheduleStatePath = path.join(outputDir, ".sync-schedule-state.json");
const logPath = path.join(outputDir, ".sync-control.log");
const port = Number(process.env.SYNC_CONTROL_PORT || 8765);
const openCommand = process.env.OPEN_COMMAND || "/usr/bin/open";
const execFileAsync = promisify(execFile);

let child = null;
let lastRun = { state: "idle", startedAt: null, finishedAt: null, exitCode: null, message: "尚未手动运行" };
let logTail = [];

async function syncLockExists() {
  try {
    await fs.access(lockPath);
    return true;
  } catch {
    return false;
  }
}

async function readScheduleState() {
  try {
    return normalizeScheduleState(JSON.parse(await fs.readFile(scheduleStatePath, "utf8")));
  } catch {
    return normalizeScheduleState();
  }
}

async function writeScheduleState(state) {
  const tempPath = `${scheduleStatePath}.tmp`;
  await fs.writeFile(tempPath, `${JSON.stringify(normalizeScheduleState(state), null, 2)}\n`, { mode: 0o600 });
  await fs.rename(tempPath, scheduleStatePath);
}

async function readSchedule() {
  const { stdout } = await execFileAsync("/usr/bin/plutil", ["-convert", "json", "-o", "-", plistPath]);
  const intervals = JSON.parse(stdout).StartCalendarInterval || [];
  return intervals
    .map((item) => `${String(item.Hour).padStart(2, "0")}:${String(item.Minute || 0).padStart(2, "0")}`)
    .sort();
}

async function writeSchedule(times) {
  const normalized = times.map((value) => {
    const match = /^(\d{1,2}):(\d{2})$/.exec(value);
    if (!match) throw new Error(`时间格式无效: ${value}`);
    const hour = Number(match[1]);
    const minute = Number(match[2]);
    if (hour > 23 || minute > 59) throw new Error(`时间超出范围: ${value}`);
    return { Hour: hour, Minute: minute };
  });
  if (normalized.length !== 2) throw new Error("请设置两个同步时间");

  const { stdout } = await execFileAsync("/usr/bin/plutil", ["-convert", "json", "-o", "-", plistPath]);
  const plist = JSON.parse(stdout);
  plist.StartCalendarInterval = normalized;
  const tempJson = `${plistPath}.json.tmp`;
  await fs.writeFile(tempJson, `${JSON.stringify(plist, null, 2)}\n`, { mode: 0o600 });
  try {
    await execFileAsync("/usr/bin/plutil", ["-convert", "xml1", "-o", plistPath, tempJson]);
  } finally {
    await fs.rm(tempJson, { force: true });
  }

  const domain = `gui/${process.getuid()}`;
  await execFileAsync("/bin/launchctl", ["bootout", `${domain}/com.emily.jconedrive.teams-sync`]).catch(() => {});
  await execFileAsync("/bin/launchctl", ["bootstrap", domain, plistPath]);
  return normalized.map(({ Hour, Minute }) => `${String(Hour).padStart(2, "0")}:${String(Minute).padStart(2, "0")}`).sort();
}

function addLog(data) {
  const text = data.toString();
  logTail.push(...text.split(/\r?\n/).filter(Boolean));
  logTail = logTail.slice(-80);
  fsSync.appendFileSync(logPath, text);
}

async function startSync() {
  if (child) throw new Error("已有手动同步任务运行中");
  if (await syncLockExists()) throw new Error("已有同步任务运行中，请等待当前任务完成");

  lastRun = { state: "running", startedAt: new Date().toISOString(), finishedAt: null, exitCode: null, message: "正在启动可见 Chrome" };
  logTail = [];
  child = spawn("/usr/local/bin/node", [scriptPath], {
    cwd: outputDir,
    env: { ...process.env, HEADLESS: "false", TEAMS_PROFILE: path.join(outputDir, ".teams-browser-profile"), CHROME_PATH: "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" },
    stdio: ["ignore", "pipe", "pipe"],
  });
  child.stdout.on("data", addLog);
  child.stderr.on("data", addLog);
  child.on("error", (error) => addLog(`控制服务启动失败: ${error.stack || error}\n`));
  child.on("close", (code, signal) => {
    lastRun = {
      ...lastRun,
      state: code === 0 ? "success" : "error",
      finishedAt: new Date().toISOString(),
      exitCode: code,
      message: code === 0 ? "同步完成" : `同步退出: code=${code ?? "null"}, signal=${signal || "unknown"}`,
    };
    child = null;
  });
}

function sendJson(response, status, value) {
  const body = JSON.stringify(value);
  response.writeHead(status, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
  response.end(body);
}

async function status() {
  const lockExists = await syncLockExists();
  const scheduleState = await readScheduleState();
  return {
    ...lastRun,
    running: Boolean(child) || lockExists,
    message: lockExists && !child ? "已有同步任务运行中" : lastRun.message,
    schedule: await readSchedule(),
    scheduleEnabled: scheduleState.enabled,
    pendingSchedule: scheduleState.pending[0] || null,
    pendingScheduleCount: scheduleState.pending.length,
    log: logTail.slice(-40),
  };
}

async function requestScheduleDue(slotId, scheduledAt) {
  if (!slotId) throw new Error("缺少定时任务时间标识");
  const current = await readScheduleState();
  const result = markScheduleDue(current, slotId, scheduledAt || new Date().toISOString());
  if (result.status === "pending") {
    await writeScheduleState(result.state);
    addLog(`定时同步待确认: ${slotId}\n`);
  }
  return result;
}

async function decidePendingSchedule(slotId, action) {
  const current = await readScheduleState();
  const result = decideSchedule(current, slotId, action);
  await writeScheduleState(result.state);
  if (result.status === "confirmed") {
    try {
      await startSync();
    } catch (error) {
      await writeScheduleState(current);
      throw error;
    }
    addLog(`定时同步已确认: ${slotId}\n`);
  } else {
    addLog(`已忽略本次定时同步: ${slotId}\n`);
  }
  return result;
}

async function loadFileIndex() {
  try {
    const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));
    return buildFileIndex({ ...manifest, directories: await collectSyncedDirectories() });
  } catch {
    return buildFileIndex({ files: {}, directories: await collectSyncedDirectories() });
  }
}

async function collectSyncedDirectories() {
  const collator = new Intl.Collator("zh-CN", { numeric: true, sensitivity: "base" });
  const rootEntries = await fs.readdir(outputDir, { withFileTypes: true });
  const roots = rootEntries
    .filter((entry) => entry.isDirectory() && (/^\d+_/.test(entry.name) || entry.name === "Shared Documents"))
    .sort((left, right) => collator.compare(left.name, right.name));
  const directories = [];

  async function visit(directoryPath, relativePath) {
    directories.push(relativePath);
    const entries = await fs.readdir(directoryPath, { withFileTypes: true });
    const children = entries
      .filter((entry) => entry.isDirectory() && !entry.name.startsWith("."))
      .sort((left, right) => collator.compare(left.name, right.name));
    for (const child of children) {
      await visit(path.join(directoryPath, child.name), path.posix.join(relativePath, child.name));
    }
  }

  for (const root of roots) await visit(path.join(outputDir, root.name), root.name);
  return directories;
}

function pathIsInside(basePath, targetPath) {
  const relative = path.relative(basePath, targetPath);
  return relative && !relative.startsWith(`..${path.sep}`) && relative !== ".." && !path.isAbsolute(relative);
}

async function openLocalPath(relativePath) {
  if (typeof relativePath !== "string" || !relativePath.trim() || path.isAbsolute(relativePath)) {
    throw new Error("本地路径无效");
  }

  const normalized = path.normalize(relativePath.trim());
  const rootName = normalized.split(path.sep)[0];
  if (!/^\d+_/.test(rootName) && rootName !== "Shared Documents") {
    throw new Error("只能打开已同步目录中的文件或文件夹");
  }

  const realOutputDir = await fs.realpath(outputDir);
  const candidatePath = path.resolve(outputDir, normalized);
  if (!pathIsInside(realOutputDir, candidatePath)) throw new Error("本地路径超出同步目录");

  const realTargetPath = await fs.realpath(candidatePath);
  if (!pathIsInside(realOutputDir, realTargetPath)) throw new Error("本地路径超出同步目录");
  const targetStat = await fs.stat(realTargetPath);
  const type = targetStat.isDirectory() ? "folder" : targetStat.isFile() ? "file" : null;
  if (!type) throw new Error("只支持打开文件或文件夹");

  await execFileAsync(openCommand, [realTargetPath]);
  return { ok: true, type };
}

function publicTree(node) {
  return {
    type: node.type,
    name: node.name,
    path: node.path,
    localName: node.localName,
    openPath: node.type === "file" ? localFilePath(node.path, node.localName) : node.path,
    children: node.children?.map(publicTree),
  };
}

function localFilePath(relativePath, localName) {
  const directory = path.posix.dirname(relativePath);
  return directory === "." ? (localName || path.posix.basename(relativePath)) : path.posix.join(directory, localName || path.posix.basename(relativePath));
}

async function fileIndex(query = "") {
  const index = await loadFileIndex();
  const trimmedQuery = query.trim();
  const results = trimmedQuery
    ? searchFileIndex(index, trimmedQuery).map(({ metadata, ...file }) => ({ ...file, openPath: localFilePath(file.path, file.localName) }))
    : [];
  const manifestStat = await fs.stat(manifestPath).catch(() => null);
  return {
    query,
    count: trimmedQuery ? results.length : index.files.length,
    stats: index.stats,
    updatedAt: manifestStat?.mtime?.toISOString() || null,
    tree: trimmedQuery ? [] : index.tree.map(publicTree),
    results,
  };
}

const html = `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Teams 同步控制台</title>
  <style>
    :root { color-scheme: light; --ink:#1d2925; --muted:#69746f; --line:#d7dfda; --paper:#f7faf8; --green:#126b51; --green-soft:#e4f2ec; --amber:#a25d14; --red:#a63838; }
    * { box-sizing:border-box; } body { margin:0; min-height:100vh; background:var(--paper); color:var(--ink); font:15px/1.5 -apple-system,BlinkMacSystemFont,"SF Pro Text","PingFang SC",sans-serif; }
    main { width:min(900px, calc(100% - 32px)); margin:40px auto; } header { display:flex; justify-content:space-between; align-items:flex-end; gap:20px; margin-bottom:24px; }
    h1 { margin:0; font-size:28px; letter-spacing:0; } .sub { margin:5px 0 0; color:var(--muted); }
    .grid { display:grid; grid-template-columns:1fr 1fr; gap:16px; } section { background:#fff; border:1px solid var(--line); border-radius:8px; padding:20px; box-shadow:0 4px 14px rgba(25,50,40,.04); }
    .wide { grid-column:1 / -1; } h2 { margin:0 0 16px; font-size:17px; } .status { display:flex; align-items:center; gap:10px; font-weight:600; }
    .dot { width:10px; height:10px; border-radius:50%; background:#929d98; } .dot.running { background:#c4781c; box-shadow:0 0 0 4px #f8ead5; } .dot.success { background:var(--green); } .dot.error { background:var(--red); }
    button { border:0; border-radius:6px; padding:10px 15px; background:var(--green); color:#fff; font:inherit; font-weight:600; cursor:pointer; } button:disabled { cursor:wait; opacity:.55; }
    .secondary { background:#edf3ef; color:var(--green); } .row { display:flex; justify-content:space-between; align-items:center; gap:12px; } .meta { color:var(--muted); font-size:13px; margin-top:12px; }
    .schedule-header { display:flex; justify-content:space-between; align-items:center; gap:12px; margin-bottom:14px; } .schedule-header h2 { margin:0; }
    .times { display:flex; gap:12px; align-items:end; flex-wrap:wrap; } label { display:grid; gap:5px; color:var(--muted); font-size:13px; } input { border:1px solid var(--line); border-radius:6px; padding:9px 10px; font:inherit; color:var(--ink); background:#fff; }
    pre { margin:0; height:260px; overflow:auto; white-space:pre-wrap; word-break:break-word; background:#18231f; color:#d8e8df; border-radius:6px; padding:14px; font:12px/1.55 ui-monospace,SFMono-Regular,Menlo,monospace; }
    .file-index { min-height:260px; } .file-index h2 { margin:0; } .file-toolbar { display:grid; gap:12px; } .file-heading { display:flex; align-items:baseline; gap:12px; flex-wrap:wrap; } .file-heading .meta { margin-top:0; } .file-search-controls { display:flex; align-items:center; gap:12px; } .file-search-controls button { flex:0 0 auto; } .file-search-controls input { flex:1 1 auto; min-width:0; } .file-tree, .file-results { margin-top:16px; border-top:1px solid var(--line); padding-top:8px; }
    .file-tree details { margin:0; } .file-tree summary { cursor:pointer; list-style:none; padding:7px 8px; border-bottom:1px solid #edf1ee; font-weight:600; } .file-tree summary::-webkit-details-marker { display:none; } .file-tree summary::before { content:"▸"; display:inline-block; width:18px; color:var(--muted); } .file-tree details[open] > summary::before { content:"▾"; } .tree-children { margin-left:20px; border-left:1px solid var(--line); } .file-node { display:flex; gap:8px; align-items:center; padding:6px 8px 6px 12px; border-bottom:1px solid #f0f3f1; } .open-item { display:inline; padding:0; border:0; border-radius:2px; background:transparent; color:inherit; font:inherit; font-weight:inherit; text-align:left; overflow-wrap:anywhere; } .open-item:hover { color:var(--green); text-decoration:underline; }
    .file-result { display:grid; grid-template-columns:minmax(160px, .8fr) minmax(240px, 2fr) auto; gap:12px; align-items:center; padding:9px 8px; border-bottom:1px solid #edf1ee; } .file-result .name { font-weight:600; overflow-wrap:anywhere; } .file-result .path { color:var(--muted); font-size:13px; overflow-wrap:anywhere; } .copy-path { padding:6px 9px; font-size:13px; background:#edf3ef; color:var(--green); }
    .notice { margin-top:16px; padding:11px 13px; border-left:3px solid #d48a2c; background:#fff7e9; color:#754515; font-size:13px; }
    .modal-backdrop { position:fixed; inset:0; display:grid; place-items:center; padding:20px; background:rgba(19,31,26,.38); z-index:10; } .modal-backdrop[hidden] { display:none; } .modal { width:min(420px,100%); background:#fff; border:1px solid var(--line); border-radius:8px; padding:24px; box-shadow:0 14px 40px rgba(18,40,30,.2); } .modal h2 { margin:0 0 8px; } .modal p { margin:0; color:var(--muted); } .modal-actions { display:flex; justify-content:flex-end; gap:10px; margin-top:22px; }
    @media (max-width:700px) { main { margin:24px auto; } header { display:block; } .grid { grid-template-columns:1fr; } .wide { grid-column:auto; } }
  </style>
</head>
<body>
  <main>
    <header><div><h1>Teams 同步控制台</h1><p class="sub">本机控制，共享文件按原目录递归同步</p></div><button id="refresh" class="secondary">刷新状态</button></header>
    <div class="grid">
      <section><h2>当前状态</h2><div class="row"><div class="status"><span id="dot" class="dot"></span><span id="state">读取中</span></div><button id="start">开始同步</button></div><div id="meta" class="meta"></div><div class="notice">首次点击可能会打开 Chrome 登录窗口。登录完成后，后续任务会复用该浏览器会话。</div></section>
      <section><div class="schedule-header"><h2>自动同步时间</h2><button id="service-toggle" class="secondary">暂停服务</button></div><div class="times"><label>时间 1<input id="time1" type="time"></label><label>时间 2<input id="time2" type="time"></label><button id="save">保存时间</button></div><div id="schedule" class="meta"></div></section>
      <section class="wide"><h2>运行日志</h2><pre id="log">暂无日志</pre></section>
      <section class="wide file-index"><div class="file-toolbar"><div class="file-heading"><h2>SharePoint 文件索引</h2><span id="file-meta" class="meta">更新于 -</span></div><span id="file-stats" class="meta">统计读取中</span><div class="file-search-controls"><button id="file-toggle" class="secondary">全部收起</button><input id="file-search" type="search" placeholder="按文件名模糊查询"></div></div><div id="file-search-meta" class="meta" hidden></div><div id="file-results" class="file-results"></div><div id="file-tree" class="file-tree">读取中</div></section>
    </div>
  </main>
  <div id="schedule-modal" class="modal-backdrop" hidden><div class="modal" role="dialog" aria-modal="true" aria-labelledby="schedule-modal-title"><h2 id="schedule-modal-title">定时同步确认</h2><p id="schedule-modal-text"></p><div class="modal-actions"><button id="skip-schedule" class="secondary">忽略本次</button><button id="confirm-schedule">继续同步</button></div></div></div>
  <script>
    const $ = (id) => document.getElementById(id);
    const fmt = (v) => v ? new Date(v).toLocaleString() : "-";
    let allExpanded = true;
    let scheduleEnabled = true;
    let pendingSlotId = null;
    async function load() {
      const r = await fetch('/api/status'); const s = await r.json();
      $('dot').className = 'dot ' + (s.running ? 'running' : s.state);
      $('state').textContent = s.running ? '同步进行中' : (s.message || s.state);
      $('meta').textContent = '开始：' + fmt(s.startedAt) + '　结束：' + fmt(s.finishedAt);
      scheduleEnabled = s.scheduleEnabled !== false;
      $('service-toggle').textContent = scheduleEnabled ? '暂停服务' : '启动服务';
      $('schedule').textContent = '当前计划：' + (s.schedule || []).join('、') + '　服务：' + (scheduleEnabled ? '已开启' : '已暂停') + (s.pendingScheduleCount ? '　待确认 ' + s.pendingScheduleCount + ' 次' : '');
      if (document.activeElement?.type !== 'time') { $('time1').value = s.schedule?.[0] || '08:00'; $('time2').value = s.schedule?.[1] || '12:00'; }
      $('start').disabled = s.running;
      $('log').textContent = (s.log || []).join('\\n') || '暂无日志'; $('log').scrollTop = $('log').scrollHeight;
      renderSchedulePrompt(s.pendingSchedule);
      await loadFiles();
    }
    function renderSchedulePrompt(schedule) {
      if (!schedule) { $('schedule-modal').hidden = true; pendingSlotId = null; return; }
      if (pendingSlotId === schedule.slotId && !$('schedule-modal').hidden) return;
      pendingSlotId = schedule.slotId;
      $('schedule-modal-text').textContent = '计划时间 ' + fmt(schedule.scheduledAt) + '，现在开始同步共享文件吗？';
      $('schedule-modal').hidden = false;
    }
    async function decideSchedule(action) {
      if (!pendingSlotId) return;
      $('confirm-schedule').disabled = true; $('skip-schedule').disabled = true;
      const response = await fetch('/api/schedule-decision', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({ slotId:pendingSlotId, action }) });
      const data = await response.json().catch(() => ({}));
      $('confirm-schedule').disabled = false; $('skip-schedule').disabled = false;
      if (!response.ok) alert(data.error || '定时同步处理失败');
      await load();
    }
    const escapeHtml = (value) => String(value ?? '').replace(/[&<>"']/g, (char) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
    const fileCount = (node) => node.type === 'file' ? 1 : (node.children || []).reduce((total, child) => total + fileCount(child), 0);
    function renderTree(nodes) {
      return (nodes || []).map((node) => node.type === 'folder'
        ? '<details open><summary>📁 <button type="button" class="open-item folder-name" data-open-path="' + escapeHtml(node.openPath) + '" title="在 Finder 中打开">' + escapeHtml(node.name) + '</button> <span class="meta">(' + fileCount(node) + ')</span></summary><div class="tree-children">' + renderTree(node.children) + '</div></details>'
        : '<div class="file-node"><span>📄</span><button type="button" class="open-item name" data-open-path="' + escapeHtml(node.openPath) + '" title="打开本地文件">' + escapeHtml(node.name) + '</button></div>').join('');
    }
    function renderResults(results) {
      return (results || []).map((file) => '<div class="file-result"><button type="button" class="open-item name" data-open-path="' + escapeHtml(file.openPath) + '" title="打开本地文件">📄 ' + escapeHtml(file.name) + '</button><span class="path">' + escapeHtml(file.directory || '(根目录)') + '</span><button class="copy-path" data-copy-path="' + escapeHtml(file.path) + '">复制路径</button></div>').join('');
    }
    async function handleOpenItem(event) {
      const button = event.target.closest('[data-open-path]');
      if (!button) return false;
      event.preventDefault(); event.stopPropagation(); button.disabled = true;
      const response = await fetch('/api/open', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({path:button.dataset.openPath}) });
      const data = await response.json().catch(() => ({}));
      button.disabled = false;
      if (!response.ok) alert(data.error || '无法打开本地文件');
      return true;
    }
    function syncFolderButtonState() {
      const folders = [...document.querySelectorAll('#file-tree details')];
      allExpanded = folders.length === 0 || folders.every((folder) => folder.open);
      $('file-toggle').textContent = allExpanded ? '全部收起' : '全部展开';
    }
    function bindFolderState() {
      document.querySelectorAll('#file-tree details').forEach((folder) => folder.addEventListener('toggle', syncFolderButtonState));
    }
    async function loadFiles() {
      const desiredExpanded = allExpanded;
      const query = $('file-search').value.trim();
      const response = await fetch('/api/files?query=' + encodeURIComponent(query));
      const data = await response.json();
      $('file-meta').textContent = '更新于 ' + (data.updatedAt ? fmt(data.updatedAt) : '-');
      $('file-search-meta').textContent = query ? '匹配 ' + data.count + ' 个文件' : '';
      $('file-search-meta').hidden = !query;
      const extensionStats = Object.entries(data.stats?.extensions || {}).map(([extension, count]) => extension === '(none)' ? '无扩展名 ' + count : extension.toUpperCase() + ' ' + count).join(' · ');
      $('file-stats').textContent = '文件夹 ' + (data.stats?.folderCount || 0) + ' · 总文件数 ' + (data.stats?.fileCount || 0) + (extensionStats ? ' · ' + extensionStats : '');
      $('file-results').innerHTML = query ? (renderResults(data.results) || '<div class="meta">没有匹配文件</div>') : '';
      $('file-tree').innerHTML = query ? '' : (renderTree(data.tree) || '<div class="meta">暂无同步文件</div>');
      $('file-toggle').disabled = Boolean(query);
      bindFolderState();
      if (!query) setAllFolders(desiredExpanded);
      else $('file-toggle').textContent = desiredExpanded ? '全部收起' : '全部展开';
    }
    function setAllFolders(expanded) {
      allExpanded = expanded;
      document.querySelectorAll('#file-tree details').forEach((folder) => { folder.open = expanded; });
      $('file-toggle').textContent = expanded ? '全部收起' : '全部展开';
    }
    $('start').onclick = async () => {
      $('start').disabled = true;
      const response = await fetch('/api/sync', { method:'POST' });
      const data = await response.json().catch(() => ({}));
      if (!response.ok) alert(data.error || '同步未启动');
      await load();
    };
    $('service-toggle').onclick = async () => { const response = await fetch('/api/schedule-service', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({enabled:!scheduleEnabled}) }); const data = await response.json().catch(() => ({})); if (!response.ok) alert(data.error || '服务状态修改失败'); await load(); };
    $('save').onclick = async () => { const r = await fetch('/api/schedule', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({times:[$('time1').value,$('time2').value]}) }); const data = await r.json(); if (!r.ok) alert(data.error || '保存失败'); await load(); };
    $('refresh').onclick = load; load(); setInterval(load, 3000);
    $('confirm-schedule').onclick = () => decideSchedule('confirm'); $('skip-schedule').onclick = () => decideSchedule('skip');
    $('file-toggle').onclick = () => setAllFolders(!allExpanded);
    let searchTimer; $('file-search').oninput = () => { clearTimeout(searchTimer); searchTimer = setTimeout(loadFiles, 180); };
    $('file-tree').onclick = handleOpenItem;
    $('file-results').onclick = async (event) => { if (await handleOpenItem(event)) return; const button = event.target.closest('[data-copy-path]'); if (!button) return; await navigator.clipboard.writeText(button.dataset.copyPath); button.textContent = '已复制'; setTimeout(() => { button.textContent = '复制路径'; }, 1200); };
  </script>
</body>
</html>`;

const server = http.createServer(async (request, response) => {
  try {
    if (request.method === "GET" && request.url === "/") {
      response.writeHead(200, { "Content-Type": "text/html; charset=utf-8" }); response.end(html); return;
    }
    if (request.method === "GET" && request.url === "/api/status") { sendJson(response, 200, await status()); return; }
    if (request.method === "GET" && request.url.startsWith("/api/files")) {
      const query = new URL(request.url, "http://127.0.0.1").searchParams.get("query") || "";
      sendJson(response, 200, await fileIndex(query)); return;
    }
    if (request.method === "POST" && request.url === "/api/open") {
      if (!String(request.headers["content-type"] || "").startsWith("application/json")) throw new Error("请求格式无效");
      let body = ""; for await (const chunk of request) body += chunk;
      const { path: relativePath } = JSON.parse(body);
      sendJson(response, 200, await openLocalPath(relativePath)); return;
    }
    if (request.method === "POST" && request.url === "/api/sync") { await startSync(); sendJson(response, 202, { ok: true }); return; }
    if (request.method === "POST" && request.url === "/api/schedule-due") {
      let body = ""; for await (const chunk of request) body += chunk;
      const { slotId, scheduledAt } = JSON.parse(body);
      const result = await requestScheduleDue(slotId, scheduledAt);
      sendJson(response, 200, { ok: true, ...result, pendingSchedule: result.state.pending[0] || null }); return;
    }
    if (request.method === "POST" && request.url === "/api/schedule-decision") {
      let body = ""; for await (const chunk of request) body += chunk;
      const { slotId, action } = JSON.parse(body);
      const result = await decidePendingSchedule(slotId, action);
      sendJson(response, 202, { ok: true, status: result.status, pendingSchedule: result.state.pending[0] || null }); return;
    }
    if (request.method === "POST" && request.url === "/api/schedule-service") {
      let body = ""; for await (const chunk of request) body += chunk;
      const { enabled } = JSON.parse(body);
      if (typeof enabled !== "boolean") throw new Error("服务状态必须是布尔值");
      const state = setScheduleEnabled(await readScheduleState(), enabled);
      await writeScheduleState(state);
      addLog(`自动同步服务已${enabled ? "启动" : "暂停"}\n`);
      sendJson(response, 200, { ok: true, enabled: state.enabled }); return;
    }
    if (request.method === "POST" && request.url === "/api/schedule") {
      let body = ""; for await (const chunk of request) body += chunk;
      const { times } = JSON.parse(body); const schedule = await writeSchedule(times); sendJson(response, 200, { ok: true, schedule }); return;
    }
    sendJson(response, 404, { error: "Not found" });
  } catch (error) { sendJson(response, 400, { error: error.message || String(error) }); }
});

server.listen(port, "127.0.0.1", () => console.log(`同步控制台已启动: http://127.0.0.1:${port}`));
