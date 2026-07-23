import test from "node:test";
import assert from "node:assert/strict";
import { buildFileIndex, searchFileIndex } from "../file-index.mjs";

test("文件索引按同步清单顺序生成嵌套目录树", () => {
  const index = buildFileIndex({
    files: {
      "00_PMO/10_Project Meetings/weekly.docx": { localName: "weekly.docx" },
      "00_PMO/10_Project Meetings/archive/old.docx": { localName: "old.docx" },
      "10_Project_Initiation & Discovery/brief.xlsx": { localName: "brief.xlsx" },
    },
  });

  assert.deepEqual(index.tree.map((node) => node.name), ["00_PMO", "10_Project_Initiation & Discovery"]);
  assert.deepEqual(index.tree[0].children.map((node) => node.name), ["10_Project Meetings"]);
  assert.deepEqual(index.tree[0].children[0].children.map((node) => node.name), ["weekly.docx", "archive"]);
  assert.equal(index.tree[0].children[0].children[1].children[0].path, "00_PMO/10_Project Meetings/archive/old.docx");
});

test("文件名搜索忽略大小写和空格并返回完整目录", () => {
  const index = buildFileIndex({
    files: {
      "00_PMO/00_Project Governance/PM-1.01 Project Charter.pptx": { localName: "PM-1.01 Project Charter.pptx" },
      "10_Project_Initiation & Discovery/brief.xlsx": { localName: "brief.xlsx" },
    },
  });

  assert.deepEqual(searchFileIndex(index, "project charter").map((file) => file.path), [
    "00_PMO/00_Project Governance/PM-1.01 Project Charter.pptx",
  ]);
});

test("文件名有匹配时不返回仅路径匹配的文件", () => {
  const index = buildFileIndex({
    files: {
      "00_PMO/Project Charter Documents/PM-1.01 Project Charter.pptx": { localName: "PM-1.01 Project Charter.pptx" },
      "00_PMO/Project Charter Documents/PM-8.01 Hypercare Plan.xlsx": { localName: "PM-8.01 Hypercare Plan.xlsx" },
    },
  });

  assert.deepEqual(searchFileIndex(index, "project charter").map((file) => file.name), ["PM-1.01 Project Charter.pptx"]);
});

test("文件索引保留空目录并统计文件夹、文件和扩展名", () => {
  const index = buildFileIndex({
    directories: ["00_PMO", "00_PMO/Empty Folder", "10_Project_Initiation & Discovery", "Shared Documents"],
    files: {
      "00_PMO/weekly.docx": { localName: "weekly.docx" },
      "10_Project_Initiation & Discovery/brief.xlsx": { localName: "brief.xlsx" },
    },
  });

  assert.deepEqual(index.tree.map((node) => node.name), ["00_PMO", "10_Project_Initiation & Discovery", "Shared Documents"]);
  assert.deepEqual(index.tree[0].children.map((node) => node.name), ["Empty Folder", "weekly.docx"]);
  assert.deepEqual(index.stats, {
    folderCount: 4,
    fileCount: 2,
    extensions: { docx: 1, xlsx: 1 },
  });
});
