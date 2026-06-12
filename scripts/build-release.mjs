import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const targets = [
  {
    name: "linux-x64",
    pkgTarget: "node20-linux-x64",
    output: "dist/md-to-docx-linux-x64",
  },
  {
    name: "linux-arm64",
    pkgTarget: "node20-linux-arm64",
    output: "dist/md-to-docx-linux-arm64",
  },
  {
    name: "macos-x64",
    pkgTarget: "node20-macos-x64",
    output: "dist/md-to-docx-macos-x64",
  },
  {
    name: "macos-arm64",
    pkgTarget: "node20-macos-arm64",
    output: "dist/md-to-docx-macos-arm64",
  },
  {
    name: "win-x64",
    pkgTarget: "node20-win-x64",
    output: "dist/md-to-docx-win-x64.exe",
  },
];

function run(cmd, args = [], opts = {}) {
  const result = spawnSync(cmd, args, {
    stdio: "inherit",
    shell: process.platform === "win32",
    ...opts,
  });
  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

fs.mkdirSync("dist", { recursive: true });

const requestedTargets = process.argv.slice(2);
const selectedTargets = requestedTargets.length > 0
  ? targets.filter((target) => requestedTargets.includes(target.name))
  : targets;

const unknownTargets = requestedTargets.filter((name) => !targets.some((target) => target.name === name));
if (unknownTargets.length > 0) {
  console.error(`Unknown release target(s): ${unknownTargets.join(", ")}`);
  console.error(`Known targets: ${targets.map((target) => target.name).join(", ")}`);
  process.exit(1);
}

for (const target of selectedTargets) {
  console.log(`Building ${target.name} (${target.pkgTarget})...`);
  run("pnpm", [
    "exec",
    "pkg",
    "build/md-to-docx.cjs",
    "--targets",
    target.pkgTarget,
    "--no-bytecode",
    "--public",
    "--output",
    target.output,
  ]);

  if (process.platform !== "win32" && path.extname(target.output) !== ".exe") {
    fs.chmodSync(target.output, 0o755);
  }
}

console.log("Release binaries written to dist/.");
