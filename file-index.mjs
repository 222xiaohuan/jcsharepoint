function findChild(children, name, type) {
  return children.find((child) => child.name === name && child.type === type);
}

function ensureDirectory(tree, parts) {
  let children = tree;
  const currentPath = [];
  for (const name of parts) {
    currentPath.push(name);
    let node = findChild(children, name, "folder");
    if (!node) {
      node = { type: "folder", name, path: currentPath.join("/"), children: [] };
      children.push(node);
    }
    children = node.children;
  }
}

function normalizeSearchText(value) {
  return Array.from(String(value).toLocaleLowerCase())
    .filter((character) => /[\p{L}\p{N}]/u.test(character))
    .join("");
}

function fuzzyMatch(value, query) {
  const text = normalizeSearchText(value);
  let cursor = 0;
  for (const character of normalizeSearchText(query)) {
    cursor = text.indexOf(character, cursor);
    if (cursor === -1) return false;
    cursor += 1;
  }
  return true;
}

export function buildFileIndex(manifest) {
  const tree = [];
  const files = [];

  for (const directory of manifest?.directories || []) {
    const parts = directory.split("/").filter(Boolean);
    if (parts.length) ensureDirectory(tree, parts);
  }

  for (const [relativePath, metadata] of Object.entries(manifest?.files || {})) {
    const parts = relativePath.split("/").filter(Boolean);
    if (!parts.length) continue;

    let children = tree;
    let currentPath = [];
    for (let index = 0; index < parts.length; index += 1) {
      const name = parts[index];
      const isFile = index === parts.length - 1;
      currentPath = [...currentPath, name];
      const type = isFile ? "file" : "folder";
      let node = findChild(children, name, type);
      if (!node) {
        node = { type, name, path: currentPath.join("/") };
        if (!isFile) node.children = [];
        children.push(node);
      }
      if (isFile) {
        node.localName = metadata?.localName || name;
        node.metadata = metadata || {};
        files.push({
          name,
          path: relativePath,
          directory: parts.slice(0, -1).join("/"),
          localName: node.localName,
          metadata: node.metadata,
        });
      } else {
        children = node.children;
      }
    }
  }

  const stats = { folderCount: 0, fileCount: files.length, extensions: {} };
  const collectStats = (nodes) => {
    for (const node of nodes) {
      if (node.type === "folder") {
        stats.folderCount += 1;
        collectStats(node.children);
        continue;
      }
      const extension = node.name.includes(".") ? node.name.split(".").pop().toLocaleLowerCase() : "(none)";
      stats.extensions[extension] = (stats.extensions[extension] || 0) + 1;
    }
  };
  collectStats(tree);

  return { tree, files, stats };
}

export function searchFileIndex(index, query) {
  if (!String(query || "").trim()) return [...(index?.files || [])];
  const files = index?.files || [];
  const byName = files.filter((file) => fuzzyMatch(file.name, query));
  return byName.length ? byName : files.filter((file) => fuzzyMatch(file.path, query));
}
