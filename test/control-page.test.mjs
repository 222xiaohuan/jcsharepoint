import test, { after, before } from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";

const port = 18000 + (process.pid % 1000);
let server;
let pageHtml;

before(async () => {
  server = spawn(process.execPath, ["sync-control-server.mjs"], {
    cwd: new URL("..", import.meta.url),
    env: { ...process.env, SYNC_CONTROL_PORT: String(port), OPEN_COMMAND: "/usr/bin/true" },
    stdio: "ignore",
  });

  const deadline = Date.now() + 5000;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(`http://127.0.0.1:${port}/`);
      if (response.ok) {
        pageHtml = await response.text();
        return;
      }
    } catch {}
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  throw new Error("测试控制服务启动超时");
});

after(() => server?.kill());

test("文件夹展开按钮位于文件名搜索框前面的同一工具栏", () => {
  const toolbar = pageHtml.match(/<div class="file-search-controls">([\s\S]*?)<\/div>/)?.[1] || "";

  assert.ok(toolbar.includes('id="file-toggle"'));
  assert.ok(toolbar.includes('id="file-search"'));
  assert.ok(toolbar.indexOf('id="file-toggle"') < toolbar.indexOf('id="file-search"'));
});

test("文件统计使用总文件数文案", () => {
  assert.ok(pageHtml.includes("总文件数"));
  assert.ok(!pageHtml.includes("' · 文件 ' + (data.stats?.fileCount"));
});

test("更新时间显示在文件索引标题后且不重复显示文件总数", () => {
  const heading = pageHtml.match(/<div class="file-heading">([\s\S]*?)<\/div>/)?.[1] || "";

  assert.ok(heading.includes("SharePoint 文件索引"));
  assert.ok(heading.includes('id="file-meta"'));
  assert.ok(heading.indexOf("SharePoint 文件索引") < heading.indexOf('id="file-meta"'));
  assert.ok(pageHtml.includes("'更新于 ' + (data.updatedAt"));
  assert.ok(!pageHtml.includes("'共 ' + data.count + ' 个文件"));
});

test("本地打开接口仅允许同步目录内真实存在的目标", async () => {
  const allowed = await fetch(`http://127.0.0.1:${port}/api/open`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: "00_PMO" }),
  });
  assert.equal(allowed.status, 200);
  assert.deepEqual(await allowed.json(), { ok: true, type: "folder" });

  const denied = await fetch(`http://127.0.0.1:${port}/api/open`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: "../../Applications" }),
  });
  assert.equal(denied.status, 400);
});

test("文件索引接口为文件夹和文件返回可打开的本地相对路径", async () => {
  const response = await fetch(`http://127.0.0.1:${port}/api/files`);
  const data = await response.json();
  const firstFolder = data.tree[0];
  const queue = [...data.tree];
  let firstFile;
  while (queue.length && !firstFile) {
    const node = queue.shift();
    if (node.type === "file") firstFile = node;
    else queue.push(...(node.children || []));
  }

  assert.equal(firstFolder.openPath, firstFolder.path);
  assert.ok(firstFile.openPath.endsWith(firstFile.localName));

  const openedFile = await fetch(`http://127.0.0.1:${port}/api/open`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: firstFile.openPath }),
  });
  assert.deepEqual(await openedFile.json(), { ok: true, type: "file" });

  const searchResponse = await fetch(`http://127.0.0.1:${port}/api/files?query=project%20charter`);
  const searchData = await searchResponse.json();
  assert.ok(searchData.results[0].openPath.endsWith(searchData.results[0].localName));
});

test("文件夹名称、文件名称和搜索结果都可调用本地打开接口", () => {
  assert.ok(pageHtml.includes('class="open-item folder-name"'));
  assert.ok(pageHtml.includes('class="open-item name"'));
  assert.ok(pageHtml.includes('data-open-path="'));
  assert.ok(pageHtml.includes("fetch('/api/open'"));
  assert.ok(pageHtml.includes("$('file-tree').onclick"));
  assert.ok(pageHtml.includes("$('file-results').onclick"));
});
