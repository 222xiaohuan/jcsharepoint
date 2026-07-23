import fs from "node:fs/promises";
import fsSync from "node:fs";
import path from "node:path";
import process from "node:process";
import { chromium } from "playwright";

const startUrl = process.argv[2] || process.env.TEAMS_URL || "https://teams.microsoft.com.mcas.ms/l/channel/19%3A59Sc_OBGl3IvUTP6zbmezHAoLqwrHmiKjvOzgWEP7m41%40thread.tacv2/Project%20Nexus?groupId=b8d50ba5-7e87-4a1a-b95d-aecccedb6a85&tenantId=ec216e11-92ba-4e65-83db-e661714f5916&ngc=true&subEntityId=FileBrowserTabApp";
const rawOutputDir = process.argv[3] || process.env.OUTPUT_DIR || "/Users/emily/Documents/jconedrive";
const outputDir = path.isAbsolute(rawOutputDir) ? rawOutputDir : path.resolve(rawOutputDir);
const rawProfileDir = process.env.TEAMS_PROFILE || path.join(outputDir, ".teams-browser-profile");
const profileDir = path.isAbsolute(rawProfileDir) ? rawProfileDir : path.resolve(rawProfileDir);
const manifestPath = path.join(outputDir, ".teams-sync-manifest.json");
const lockDir = path.join(outputDir, ".teams-sync.lockdir");
const chromePath = process.env.CHROME_PATH || "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function safeName(value) {
  return value
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/[ .]+$/g, "")
    .trim() || "unnamed";
}

async function loadManifest() {
  try {
    const data = JSON.parse(fsSync.readFileSync(manifestPath, "utf8"));
    return data?.files && typeof data.files === "object" ? data : { version: 1, files: {} };
  } catch {
    return { version: 1, files: {} };
  }
}

async function saveManifest(manifest) {
  const tempPath = `${manifestPath}.tmp`;
  await fs.writeFile(tempPath, `${JSON.stringify(manifest, null, 2)}\n`, { mode: 0o600 });
  await fs.rename(tempPath, manifestPath);
}

async function clearStaleBrowserLocks() {
  for (const name of ["SingletonLock", "SingletonCookie", "SingletonSocket"]) {
    const lock = path.join(profileDir, name);
    fsSync.rmSync(lock, { force: true });
  }
}

async function pageLooksLikeLogin(page) {
  return /login\.microsoftonline\.com|login\.live\.com/i.test(page.url());
}

function pageLooksLikeLauncher(page) {
  return /\/dl\/launcher\/launcher\.html/i.test(page.url());
}

async function waitForApp(page) {
  const deadline = Date.now() + (process.env.HEADLESS === "false" ? 15 * 60_000 : 30_000);
  for (;;) {
    if (!(await pageLooksLikeLogin(page)) && !pageLooksLikeLauncher(page)) return;
    if (Date.now() >= deadline) {
      throw new Error("浏览器登录状态已失效。请先手动运行一次：HEADLESS=false node download-teams-files.mjs");
    }
    await sleep(1000);
  }
}

async function findDocumentScope(page) {
  for (const frame of page.frames()) {
    const rowCount = await frame.locator('[role="row"], tr').count().catch(() => 0);
    if (rowCount >= 1) return frame;
  }
  return null;
}

async function waitForDocumentList(page) {
  await page.waitForLoadState("domcontentloaded", { timeout: 30000 }).catch(() => {});
  const deadline = Date.now() + 180_000;
  for (;;) {
    if (!(await pageLooksLikeLogin(page))) {
      const scope = await findDocumentScope(page);
      if (scope) return scope;
    }
    if (Date.now() >= deadline) {
      const pageHint = await page.locator("body").innerText({ timeoutMs: 3000 }).catch(() => "");
      throw new Error(`等待 Teams 文档列表超时，请检查登录状态和共享页面权限。URL=${page.url()} TITLE=${await page.title()} PAGE=${pageHint.replace(/\s+/g, " ").slice(0, 300)}`);
    }
    await sleep(1000);
  }
}

async function useTeamsWebApp(page) {
  for (const scope of page.frames()) {
    const candidates = [
      scope.getByText("改用 Web 应用", { exact: true }),
      scope.getByText("Use the web app instead", { exact: true }),
    ];
    for (const candidate of candidates) {
      const count = await candidate.count().catch(() => 0);
      if (count !== 1 || !(await candidate.isVisible().catch(() => false))) continue;
      await candidate.click();
      await sleep(1500);
      return true;
    }
  }
  return false;
}

async function openSharedTab(page) {
  for (const scope of page.frames()) {
    const candidates = [
      scope.locator('[role="tab"]').filter({ hasText: /已共享|Shared/i }),
      scope.locator('button, a').filter({ hasText: /已共享|Shared/i }),
    ];

    for (const candidate of candidates) {
      const count = await candidate.count().catch(() => 0);
      if (count === 1) {
        await candidate.click();
        await sleep(1200);
        return true;
      }
    }
  }

  // Teams sometimes exposes the channel selector as a dropdown instead of tabs.
  for (const scope of page.frames()) {
    const triggerCandidates = [
      scope.locator('button, [role="tab"], [aria-haspopup="menu"]').filter({ hasText: /^对话$/ }),
      scope.getByText("对话", { exact: true }),
    ];
    for (const candidate of triggerCandidates) {
      const visible = [];
      for (const item of await candidate.all()) {
        if (await item.isVisible()) visible.push(item);
      }
      if (!visible.length) continue;
      await visible[visible.length - 1].click();
      await sleep(500);
      const options = scope.getByText("已共享", { exact: true });
      const visibleOptions = [];
      for (const option of await options.all()) {
        if (await option.isVisible()) visibleOptions.push(option);
      }
      if (visibleOptions.length) {
        await visibleOptions[visibleOptions.length - 1].click();
        await sleep(1200);
        return true;
      }
    }
  }
  return false;
}

async function readItems(scope) {
  return scope.evaluate(() => {
    const rows = Array.from(document.querySelectorAll('[role="row"], tr'));
    const seen = new Set();
    const items = [];

    for (const row of rows) {
      const text = (row.innerText || "").trim();
      if (!text || seen.has(text) || /^(名称|Name)\s+/i.test(text)) continue;

      const lines = text.split(/\n+/).map((line) => line.trim()).filter(Boolean);
      const name = lines[0];
      if (!name || /^(名称|Name|修��时间|Modified|修改者|Modified by)$/i.test(name)) continue;

      const hints = Array.from(row.querySelectorAll("[aria-label], [title], [data-icon-name]"))
        .flatMap((el) => [el.getAttribute("aria-label"), el.getAttribute("title"), el.getAttribute("data-icon-name")])
        .filter(Boolean)
        .join(" ");
      const rowClass = typeof row.className === "string" ? row.className : "";
      const markup = (row.innerHTML || "").slice(0, 5000);
      const isFolder = /folder|文件夹|文件夹图标|directory/i.test(`${hints} ${rowClass} ${markup}`);
      const cells = Array.from(row.querySelectorAll('[role="gridcell"], td'));
      const modifiedCell = cells[1];
      const modified = modifiedCell
        ? `${modifiedCell.innerText || ""} ${Array.from(modifiedCell.querySelectorAll("[aria-label], [title]")).map((el) => `${el.getAttribute("aria-label") || ""} ${el.getAttribute("title") || ""}`).join(" ")}`.trim()
        : lines.slice(1, 3).join(" | ");
      const signature = JSON.stringify({ name, modified, hints });

      seen.add(text);
      items.push({ name, isFolder, text, signature });
    }
    return items;
  });
}

async function rowFor(scope, item) {
  const rows = scope.locator('[role="row"], tr').filter({ hasText: item.name });
  const visible = [];
  const exact = [];
  const normalize = (value) => value.replace(/\s+/g, " ").trim();
  for (const row of await rows.all()) {
    if (!(await row.isVisible())) continue;
    visible.push(row);
    if (normalize(await row.innerText()) === normalize(item.text)) exact.push(row);
  }
  if (exact.length === 1) return exact[0];
  if (visible.length === 1) return visible[0];
  return exact[0] || visible[0] || null;
}

async function visibleDownloadControl(scope) {
  const candidates = scope.locator(
    '[aria-label*="下载" i], [title*="下载" i], [data-icon-name*="download" i], [aria-label*="download" i], [title*="download" i]'
  );
  const visible = [];
  for (const candidate of await candidates.all()) {
    if (await candidate.isVisible()) visible.push(candidate);
  }
  return visible.length === 1 ? visible[0] : null;
}

async function visibleDownloadMenuItem(page) {
  for (const menuScope of page.frames()) {
    const candidates = [
      menuScope.getByRole("menuitem", { name: /下载|download/i }),
      menuScope.getByText("下载", { exact: true }),
      menuScope.getByText("Download", { exact: true }),
    ];
    for (const candidate of candidates) {
      const visible = [];
      for (const item of await candidate.all()) {
        if (await item.isVisible()) visible.push(item);
      }
      if (visible.length === 1) return visible[0];
    }
  }
  return null;
}

async function downloadFile(page, scope, row, targetPath) {
  await row.click();
  let control = await visibleDownloadControl(scope);

  if (!control) {
    const more = row.getByRole("button", { name: /更多|more options|more/i });
    if (await more.count() === 1) {
      await more.click();
      control = await visibleDownloadMenuItem(page);
    }
  }

  if (!control) {
    throw new Error(`找不到下载控件，文件: ${await row.innerText()}`);
  }

  const downloadPromise = page.waitForEvent("download", { timeout: 60000 });
  await control.click();
  let download;
  try {
    download = await downloadPromise;
  } catch (error) {
    console.warn(`下载事件超时，跳过: ${(await row.innerText()).trim()}`);
    return null;
  }
  const suggestedName = safeName(download.suggestedFilename());
  const tempPath = path.join(outputDir, ".downloads", `${process.pid}-${Date.now()}-${suggestedName}.part`);
  await fs.mkdir(path.dirname(tempPath), { recursive: true });
  await download.saveAs(tempPath);

  await fs.mkdir(path.dirname(targetPath), { recursive: true });
  await fs.rename(tempPath, targetPath);
  console.log(`已同步: ${path.relative(outputDir, targetPath)}`);
  return { localName: path.basename(targetPath), generatedArchive: false };
}

async function openFolder(page, row, folderPath) {
  await row.dblclick();
  await sleep(1200);
  await waitForDocumentList(page);
  await fs.mkdir(folderPath, { recursive: true });
}

async function visit(page, relativeDir, visited, manifest) {
  const currentKey = `${page.url()}|${relativeDir}`;
  if (visited.has(currentKey)) return;
  visited.add(currentKey);

  const targetDir = path.join(outputDir, relativeDir);
  await fs.mkdir(targetDir, { recursive: true });
  let scope = await waitForDocumentList(page);
  const items = await readItems(scope);
  if (!items.length) {
    console.log(`空目录，跳过: ${relativeDir || "."}`);
    return;
  }

  console.log(`扫描 ${relativeDir || "."}: ${items.length} 项`);
  for (const item of items) {
    const row = await rowFor(scope, item);
    if (!row) {
      console.warn(`跳过无法唯一定位的项目: ${item.name}`);
      continue;
    }

    if (item.isFolder) {
      const childDir = path.join(relativeDir, safeName(item.name));
      await openFolder(page, row, path.join(outputDir, childDir));
      await visit(page, childDir, visited, manifest);
      await page.goBack({ waitUntil: "commit", timeout: 10000 }).catch(() => {});
      scope = await waitForDocumentList(page);
      continue;
    }

    const key = path.join(relativeDir, item.name);
    const previous = manifest.files[key];
    const localName = previous?.localName || safeName(item.name);
    const expected = path.join(targetDir, localName);
    let exists = true;
    try { await fs.access(expected); } catch { exists = false; }

    if (previous?.generatedArchive && previous.signature === item.signature) {
      console.log(`未变化，跳过打包项: ${path.relative(outputDir, targetDir)}`);
      continue;
    }

    if (exists && previous?.signature === item.signature) {
      console.log(`未变化，跳过: ${path.relative(outputDir, expected)}`);
      continue;
    }

    if (exists && !previous) {
      manifest.files[key] = { signature: item.signature, localName };
      await saveManifest(manifest);
      console.log(`建立同步基线，跳过: ${path.relative(outputDir, expected)}`);
      continue;
    }

    const result = await downloadFile(page, scope, row, expected);
    if (result) {
      manifest.files[key] = {
        signature: item.signature,
        localName: result.localName || null,
        generatedArchive: result.generatedArchive,
      };
      await saveManifest(manifest);
    }
  }
}

async function removeOneDriveArchives() {
  const entries = await fs.readdir(outputDir, { withFileTypes: true });
  for (const entry of entries) {
    const entryPath = path.join(outputDir, entry.name);
    if (entry.isDirectory() && !entry.name.startsWith(".")) {
      await removeOneDriveArchivesFrom(entryPath);
    } else if (entry.isFile() && /onedrive.*\.zip$/i.test(entry.name)) {
      await fs.rm(entryPath, { force: true });
      console.log(`已清理本地 OneDrive 压缩包: ${path.relative(outputDir, entryPath)}`);
    }
  }
}

async function removeOneDriveArchivesFrom(directory) {
  const entries = await fs.readdir(directory, { withFileTypes: true });
  for (const entry of entries) {
    const entryPath = path.join(directory, entry.name);
    if (entry.isDirectory() && !entry.name.startsWith(".")) {
      await removeOneDriveArchivesFrom(entryPath);
    } else if (entry.isFile() && /onedrive.*\.zip$/i.test(entry.name)) {
      await fs.rm(entryPath, { force: true });
      console.log(`已清理本地 OneDrive 压缩包: ${path.relative(outputDir, entryPath)}`);
    }
  }
}

const manifest = await loadManifest();
let context;
let lockAcquired = false;
try {
  await fs.mkdir(outputDir, { recursive: true });
  try {
    await fs.mkdir(lockDir);
    lockAcquired = true;
  } catch {
    console.log("已有 Teams 同步任务运行，本次跳过。");
  }

  if (lockAcquired) {
    await clearStaleBrowserLocks();
    context = await chromium.launchPersistentContext(profileDir, {
      headless: process.env.HEADLESS !== "false",
      executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH || chromePath,
      acceptDownloads: true,
      downloadsPath: path.join(outputDir, ".downloads"),
      viewport: { width: 1440, height: 1000 },
    });
    const pages = context.pages();
    const page = pages[0] || await context.newPage();
    await page.goto(startUrl, { waitUntil: "domcontentloaded" });
    await useTeamsWebApp(page);

    if (await pageLooksLikeLogin(page) || pageLooksLikeLauncher(page)) {
      console.log("请在打开的浏览器中完成 Microsoft 登录和 MFA；脚本会自动继续。\n");
    }
    await waitForApp(page);
    await openSharedTab(page);
    await waitForApp(page);
    console.log("已进入文档列表，开始递归下载。按 Ctrl+C 可安全停止，已完成的文件会保留。\n");
    await visit(page, "", new Set(), manifest);
    await removeOneDriveArchives();
    console.log("\n下载完成。");
  }
} finally {
  if (context) await context.close();
  if (lockAcquired) await fs.rm(lockDir, { recursive: true, force: true });
}
