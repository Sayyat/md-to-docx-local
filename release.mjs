import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const run = (cmd, args = [], opts = {}) => {
  const result = spawnSync(cmd, args, {
    stdio: "inherit",
    shell: process.platform === "win32",
    ...opts,
  });
  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
  return result;
};

const out = (cmd, args = [], opts = {}) => {
  const result = spawnSync(cmd, args, {
    encoding: "utf8",
    shell: process.platform === "win32",
    ...opts,
  });
  if (result.status !== 0) {
    console.error(`Command failed: ${cmd} ${args.join(" ")}`);
    if (result.stderr) {
      console.error(result.stderr);
    }
    process.exit(result.status ?? 1);
  }
  return result.stdout.toString().trim();
};

const versionType = process.argv[2] || "patch";
let notes = null;
for (let index = 3; index < process.argv.length; index += 1) {
  if (process.argv[index] === "--notes") {
    notes = process.argv.slice(index + 1).join(" ");
    break;
  }
}

if (spawnSync("git", ["rev-parse", "--is-inside-work-tree"], { stdio: "ignore" }).status !== 0) {
  console.error("This project is not inside a git repository yet.");
  console.error("Initialize it and push it to GitHub before running `pnpm release`.");
  process.exit(1);
}

const pkgJsonPath = path.resolve("package.json");
const pkg = JSON.parse(fs.readFileSync(pkgJsonPath, "utf8"));

console.log(`Releasing ${pkg.name} from v${pkg.version} using ${versionType}...`);

const status = out("git", ["status", "--porcelain"]);
if (status) {
  console.error("Git working directory is not clean. Commit or stash your changes first.");
  console.error(status);
  process.exit(1);
}

run("pnpm", ["version", versionType, "--git-tag-version=false"]);

if (fs.existsSync("package-lock.json")) {
  run("npm", ["install", "--package-lock-only", "--ignore-scripts"]);
}

const nextPkg = JSON.parse(fs.readFileSync(pkgJsonPath, "utf8"));
const tag = `v${nextPkg.version}`;

if (spawnSync("git", ["rev-parse", "-q", "--verify", `refs/tags/${tag}`], { stdio: "ignore" }).status === 0) {
  console.error(`Git tag already exists: ${tag}`);
  process.exit(1);
}

console.log(`Committing version bump for ${tag}...`);
run("git", ["add", "package.json"]);
for (const lockfile of ["pnpm-lock.yaml", "package-lock.json"]) {
  if (fs.existsSync(lockfile)) {
    run("git", ["add", lockfile]);
  }
}
run("git", ["commit", `-mchore(release): ${tag}`]);

console.log(`Creating annotated tag ${tag}...`);
run("git", ["tag", "-a", tag, "-m", notes || `Release ${tag}`]);

console.log("Pushing commit and tag...");
run("git", ["push"]);
run("git", ["push", "origin", tag]);

console.log(`Release tag ${tag} pushed. GitHub Actions will build binaries and publish the release assets.`);
