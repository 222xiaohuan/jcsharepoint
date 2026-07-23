import fs from "node:fs/promises";
import fsSync from "node:fs";
import path from "node:path";
import process from "node:process";
import { Readable } from "node:stream";
import { pipeline } from "node:stream/promises";
import { PublicClientApplication } from "@azure/msal-node";

const outputDir = path.resolve(process.env.OUTPUT_DIR || "/Users/emily/Documents/jconedrive");
const tenantId = process.env.TENANT_ID || "ec216e11-92ba-4e65-83db-e661714f5916";
const hostname = process.env.SHAREPOINT_HOST || "hkjcazure.sharepoint.com";
const sitePath = process.env.SHAREPOINT_SITE || "/sites/RTAReplacementworkingteam-ProjectRTARevamp";
const driveName = process.env.SHAREPOINT_DRIVE || "Documents";
const clientId = process.env.GRAPH_CLIENT_ID || "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const graphBase = "https://graph.microsoft.com/v1.0";
const tokenCachePath = path.join(outputDir, ".graph-msal-cache.json");
const manifestPath = path.join(outputDir, ".sharepoint-sync-manifest.json");
const lockDir = path.join(outputDir, ".sharepoint-sync.lockdir");

function safeName(value) {
  return value.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").replace(/[ .]+$/g, "").trim() || "unnamed";
}

function cachePlugin() {
  return {
    async beforeCacheAccess(context) {
      try { context.tokenCache.deserialize(fsSync.readFileSync(tokenCachePath, "utf8")); } catch {}
    },
    async afterCacheAccess(context) {
      if (context.cacheHasChanged) fsSync.writeFileSync(tokenCachePath, context.tokenCache.serialize(), { mode: 0o600 });
    },
  };
}

async function getAccessToken() {
  const app = new PublicClientApplication({
    auth: { clientId, authority: `https://login.microsoftonline.com/${tenantId}` },
    cache: { cachePlugin: cachePlugin() },
  });
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length) {
    try {
      const result = await app.acquireTokenSilent({
        account: accounts[0],
        scopes: ["Sites.Read.All", "Files.Read.All"],
      });
      return result.accessToken;
    } catch {}
  }

  const result = await app.acquireTokenByDeviceCode({
    scopes: ["Sites.Read.All", "Files.Read.All", "offline_access"],
    deviceCodeCallback: (response) => {
      console.log(`\n首次授权需要登录：${response.message}\n`);
    },
  });
  return result.accessToken;
}

async function graph(token, endpoint, options = {}) {
  const response = await fetch(`${graphBase}${endpoint}`, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, ...(options.headers || {}) },
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph ${response.status} ${endpoint}: ${body.slice(0, 500)}`);
  }
  return response;
}

async function graphJson(token, endpoint) {
  return (await graph(token, endpoint)).json();
}

async function listChildren(token, driveId, itemId) {
  let endpoint = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children?$select=id,name,size,lastModifiedDateTime,eTag,file,folder,parentReference&$top=200`;
  const items = [];
  while (endpoint) {
    const page = await graphJson(token, endpoint);
    items.push(...(page.value || []));
    endpoint = page["@odata.nextLink"]?.replace(graphBase, "") || null;
  }
  return items;
}

async function downloadItem(token, driveId, itemId, targetPath) {
  const response = await graph(token, `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`, { redirect: "follow" });
  const tempPath = `${targetPath}.sync-${process.pid}.tmp`;
  await fs.mkdir(path.dirname(targetPath), { recursive: true });
  await pipeline(Readable.fromWeb(response.body), fsSync.createWriteStream(tempPath));
  await fs.rename(tempPath, targetPath);
}

async function loadManifest() {
  try { return JSON.parse(await fs.readFile(manifestPath, "utf8")); } catch { return { version: 1, items: {} }; }
}

async function saveManifest(manifest) {
  const tempPath = `${manifestPath}.tmp`;
  await fs.writeFile(tempPath, `${JSON.stringify(manifest, null, 2)}\n`, { mode: 0o600 });
  await fs.rename(tempPath, manifestPath);
}

async function syncFolder(token, driveId, itemId, relativeDir, manifest) {
  const children = await listChildren(token, driveId, itemId);
  for (const item of children) {
    const name = safeName(item.name);
    const relativePath = path.join(relativeDir, name);
    const target = path.join(outputDir, relativePath);

    if (item.folder) {
      await fs.mkdir(target, { recursive: true });
      await syncFolder(token, driveId, item.id, relativePath, manifest);
      continue;
    }
    if (!item.file) continue;

    const previous = manifest.items[item.id];
    let localExists = true;
    try { await fs.access(target); } catch { localExists = false; }

    if (localExists && previous?.eTag === item.eTag && previous.path === relativePath) {
      console.log(`未变化，跳过: ${relativePath}`);
      continue;
    }

    if (localExists && !previous) {
      const localStat = await fs.stat(target);
      const remoteTime = Date.parse(item.lastModifiedDateTime || "");
      if (Number.isFinite(remoteTime) && localStat.size === item.size && localStat.mtimeMs >= remoteTime - 120000) {
        manifest.items[item.id] = { eTag: item.eTag, path: relativePath, size: item.size, lastModifiedDateTime: item.lastModifiedDateTime };
        await saveManifest(manifest);
        console.log(`建立基线，跳过: ${relativePath}`);
        continue;
      }
    }

    await downloadItem(token, driveId, item.id, target);
    manifest.items[item.id] = { eTag: item.eTag, path: relativePath, size: item.size, lastModifiedDateTime: item.lastModifiedDateTime };
    await saveManifest(manifest);
    console.log(`已同步: ${relativePath}`);
  }
}

async function main() {
  try {
    await fs.mkdir(outputDir, { recursive: true });
    try {
      await fs.mkdir(lockDir);
    } catch {
      console.log("已有 Graph 同步任务运行，本次跳过。");
      return;
    }

    try {
      const token = await getAccessToken();
      const site = await graphJson(token, `/sites/${hostname}:${sitePath}?$select=id,name,webUrl`);
      const drives = await graphJson(token, `/sites/${encodeURIComponent(site.id)}/drives?$select=id,name,driveType,webUrl`);
      const drive = (drives.value || []).find((candidate) => candidate.name === driveName)
        || (drives.value || []).find((candidate) => candidate.driveType === "documentLibrary");
      if (!drive) throw new Error(`找不到 SharePoint 文档库: ${driveName}`);

      const manifest = await loadManifest();
      const root = await graphJson(token, `/drives/${encodeURIComponent(drive.id)}/root?$select=id,name`);
      console.log(`开始同步: ${site.name} / ${drive.name}`);
      await syncFolder(token, drive.id, root.id, "", manifest);
      console.log("同步完成。");
    } finally {
      await fs.rm(lockDir, { recursive: true, force: true });
    }
  } catch (error) {
    console.error(error.stack || error.message || error);
    process.exitCode = 1;
  }
}

await main();
