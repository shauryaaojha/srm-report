#!/usr/bin/env node
// Installs the srm-report skill for AI coding agents, and sends a user's improvements back as a pull request.
const { execFileSync } = require("child_process");
const crypto = require("crypto");
const fs = require("fs");
const os = require("os");
const path = require("path");

const OWNER = "shauryaaojha";
const REPO = "srm-report";
const PKG_ROOT = path.join(__dirname, "..");
const SKILL_PATHS = ["SKILL.md", "references", "scripts", "assets"];
// Global skill folders from each agent's own docs. `home` is what marks the agent as present on this machine.
const TARGETS = [
  { id: "agents", dir: [".agents", "skills"], reads: "Codex, Cursor, Gemini CLI, GitHub Copilot, Windsurf, OpenCode, Roo Code" },
  { id: "claude", dir: [".claude", "skills"], reads: "Claude Code", home: ".claude" },
  { id: "antigravity", dir: [".gemini", "config", "skills"], reads: "Antigravity", home: ".gemini" },
  { id: "cline", dir: [".cline", "skills"], reads: "Cline", home: ".cline" },
  { id: "kiro", dir: [".kiro", "skills"], reads: "Kiro", home: ".kiro" },
];
const SHARED = ["codex", "cursor", "gemini", "copilot", "windsurf", "opencode", "roo"]; // all read ~/.agents/skills
const targetDir = (t) => path.join(os.homedir(), ...t.dir, REPO);
// Reports and templates must never travel. Only text that improves the skill goes in a pull request.
const BLOCKED_FILE = /\.(docx?|dotx|pdf|pptx?|xlsx?|png|jpe?g|gif|zip|tgz)$/i;
const SKIP = /(^|[\\/])(__pycache__|\.git)([\\/]|$)|\.pyc$/;
const PERSONAL = [
  [/\bRA\d{10,}\b/, "SRM registration number"],
  [/[\w.+-]+@[\w-]+\.[\w.]{2,}/, "email address"],
  [/(?<![\d.])[6-9]\d{9}(?![\d.])/, "phone number"],
];

const HELP = `srm-report: SRM report skill for AI coding agents

  npx srm-report install
      Install for every agent found on this machine. ~/.agents/skills is always included.

  npx srm-report install --agent NAMES
      Install for the named agents only, comma separated, or "all":
      ${[...SHARED, ...TARGETS.slice(1).map((t) => t.id)].join(", ")}

  npx srm-report install --dir PATH
      Install to one folder, for any other agent. Then tell it to follow PATH/SKILL.md.

      Add --force to replace an existing copy.

  npx srm-report contribute [--dir PATH]
      Dry run. Compare your copy with the version you installed and print the diff. Sends nothing.

  npx srm-report contribute --yes --title "What improved" [--body "Why"] [--dir PATH]
      Push those changes and open a pull request. Needs the GitHub CLI, signed in.
      Run this only after the person whose machine this is has read the diff and agreed.
`;

function parse(argv) {
  const out = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--yes" || a === "--force" || a === "--help") out[a.slice(2)] = true;
    else if (a.startsWith("--")) {
      if (i + 1 >= argv.length) die(`${a} needs a value`);
      out[a.slice(2)] = argv[++i];
    } else out._.push(a);
  }
  return out;
}

function die(msg, code = 1) {
  console.error(`srm-report: ${msg}`);
  process.exit(code);
}

function run(cmd, args, cwd) {
  return execFileSync(cmd, args, { cwd, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] }).trim();
}

function tryRun(cmd, args, cwd) {
  try {
    return run(cmd, args, cwd);
  } catch {
    return null;
  }
}

function copySkill(from, to) {
  for (const item of SKILL_PATHS) {
    const src = path.join(from, item);
    if (fs.existsSync(src)) fs.cpSync(src, path.join(to, item), { recursive: true, filter: (s) => !SKIP.test(s) });
  }
}

function chooseTargets(opts) {
  if (opts.dir) return [{ path: path.resolve(opts.dir), reads: "the agent you point at it" }];
  let picked;
  if (opts.agent) {
    const names = opts.agent.toLowerCase().split(",").map((s) => s.trim()).filter(Boolean);
    picked = names.includes("all") ? TARGETS : TARGETS.filter((t) => names.includes(t.id) || (t.id === "agents" && names.some((n) => SHARED.includes(n))));
    const known = ["all", ...SHARED, ...TARGETS.map((t) => t.id)];
    const bad = names.filter((n) => !known.includes(n));
    if (bad.length || !picked.length) die(`unknown agent "${bad.join(", ") || opts.agent}". Known: ${known.join(", ")}`);
  } else {
    picked = TARGETS.filter((t) => !t.home || fs.existsSync(path.join(os.homedir(), t.home)));
  }
  return picked.map((t) => ({ path: targetDir(t), reads: t.reads }));
}

function skillHash(dir) {
  const h = crypto.createHash("sha256");
  const files = SKILL_PATHS.flatMap((p) => {
    const full = path.join(dir, p);
    if (!fs.existsSync(full)) return [];
    return fs.statSync(full).isDirectory() ? listFiles(full).map((f) => path.join(p, f)) : [p];
  }).sort();
  for (const f of files) h.update(f.replace(/\\/g, "/")).update(fs.readFileSync(path.join(dir, f)));
  return h.digest("hex");
}

function install(opts) {
  const { version } = JSON.parse(fs.readFileSync(path.join(PKG_ROOT, "package.json"), "utf8"));
  let done = 0;
  for (const t of chooseTargets(opts)) {
    if (fs.existsSync(path.join(t.path, "SKILL.md")) && !opts.force) {
      console.log(`Skipped ${t.path}: already there, and it may hold your improvements. Send them with "npx srm-report contribute", or pass --force to replace it.`);
      continue;
    }
    for (const p of SKILL_PATHS) fs.rmSync(path.join(t.path, p), { recursive: true, force: true });
    fs.mkdirSync(t.path, { recursive: true });
    copySkill(PKG_ROOT, t.path);
    fs.writeFileSync(path.join(t.path, "VERSION"), version + "\n");
    fs.writeFileSync(path.join(t.path, "INSTALL_HASH"), skillHash(t.path) + "\n");
    console.log(`Installed ${version} to ${t.path} (for ${t.reads})`);
    done++;
  }
  if (done) {
    console.log("Restart your agent so it finds the skill.");
    console.log("If you improve it while making a report, send it back: npx srm-report contribute");
  }
}

// With several copies installed, the one to send is the one that changed since install.
function findCopy(opts) {
  if (opts.dir) {
    const d = path.resolve(opts.dir);
    if (!fs.existsSync(path.join(d, "SKILL.md"))) die(`no skill found at ${d}`);
    return d;
  }
  const copies = TARGETS.map(targetDir).filter((d) => fs.existsSync(path.join(d, "SKILL.md")));
  if (!copies.length) die('no installed skill found. Run "npx srm-report install", or pass --dir PATH.');
  const stored = (d) => {
    const f = path.join(d, "INSTALL_HASH");
    return fs.existsSync(f) ? fs.readFileSync(f, "utf8").trim() : null;
  };
  const changed = copies.filter((d) => skillHash(d) !== stored(d));
  if (changed.length > 1 && new Set(changed.map(skillHash)).size > 1) {
    die(`more than one copy has changed. Pick one with --dir, after merging what you want to keep:\n  ${changed.join("\n  ")}`);
  }
  return changed[0] || copies[0];
}

function listFiles(dir, base = dir, out = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (SKIP.test(p)) continue;
    if (e.isDirectory()) listFiles(p, base, out);
    else out.push(path.relative(base, p));
  }
  return out;
}

function scanAdded(diff) {
  const hits = [];
  for (const line of diff.split("\n")) {
    if (!line.startsWith("+") || line.startsWith("+++")) continue;
    for (const [rx, label] of PERSONAL) if (rx.test(line)) hits.push(`${label}: ${line.slice(1).trim().slice(0, 100)}`);
  }
  return hits;
}

function contribute(opts) {
  const dir = findCopy(opts);
  console.log(`Using the copy at ${dir}`);
  if (tryRun("gh", ["auth", "status"]) === null) die('the GitHub CLI is missing or signed out. Install gh and run "gh auth login".');
  if (opts.yes && !opts.title) die('--yes needs --title "what improved"');

  const blocked = SKILL_PATHS.flatMap((p) => {
    const full = path.join(dir, p);
    if (!fs.existsSync(full)) return [];
    return fs.statSync(full).isDirectory() ? listFiles(full).map((f) => path.join(p, f)) : [p];
  }).filter((f) => BLOCKED_FILE.test(f));
  if (blocked.length) die(`these files can't go in a pull request. Reports and templates stay on your machine. Move them out of the skill folder:\n  ${blocked.join("\n  ")}`, 2);

  const login = run("gh", ["api", "user", "--jq", ".login"]);
  const own = login === OWNER;
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "srm-report-"));
  const work = path.join(tmp, REPO);
  try {
    if (!own) run("gh", ["repo", "fork", `${OWNER}/${REPO}`, "--clone=false"]);
    run("gh", ["repo", "clone", `${login}/${REPO}`, work]);
    if (!own) tryRun("git", ["fetch", "upstream", "--tags"], work);

    // Start from the version that was installed, so changes made upstream since then don't look like deletions.
    const vf = path.join(dir, "VERSION");
    const version = fs.existsSync(vf) ? fs.readFileSync(vf, "utf8").trim() : null;
    const tag = version ? `v${version}` : null;
    const hasTag = tag && tryRun("git", ["rev-parse", "--verify", "--quiet", `refs/tags/${tag}`], work) !== null;
    const branch = `improve/${(opts.title || "changes").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 40) || "changes"}-${Date.now().toString(36)}`;
    run("git", hasTag ? ["checkout", "-b", branch, tag] : ["checkout", "-b", branch], work);

    for (const p of SKILL_PATHS) fs.rmSync(path.join(work, p), { recursive: true, force: true });
    copySkill(dir, work);
    run("git", ["add", "-A", "--", ...SKILL_PATHS], work);

    const stat = run("git", ["diff", "--cached", "--stat"], work);
    if (!stat) {
      console.log(`No changes: your copy matches ${hasTag ? tag : "the current skill"}.`);
      return;
    }
    const diff = run("git", ["diff", "--cached", "-U2"], work);
    const hits = scanAdded(diff);
    if (hits.length) die(`personal details in your changes. Remove them and run this again. Nothing was sent.\n  ${hits.join("\n  ")}`, 2);

    if (!opts.yes) {
      console.log(diff);
      console.log(`\n${stat}\n\nDRY RUN: nothing was sent. Read the whole diff, since names can't be detected automatically.`);
      console.log('To send it, with the owner\'s agreement: npx srm-report contribute --yes --title "What improved" --body "Why"');
      return;
    }
    run("git", ["commit", "-m", opts.title], work);
    run("git", ["push", "-u", "origin", branch], work);
    const head = own ? branch : `${login}:${branch}`;
    const body = opts.body || `Improvements found while making a report.\n\n\`\`\`\n${stat}\n\`\`\``;
    console.log(run("gh", ["pr", "create", "--repo", `${OWNER}/${REPO}`, "--base", "main", "--head", head, "--title", opts.title, "--body", body], work));
  } catch (e) {
    die(((e.stderr && e.stderr.toString()) || e.message).trim());
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

const opts = parse(process.argv.slice(2));
const cmd = opts.help ? "help" : opts._[0] || "help";
if (cmd === "install") install(opts);
else if (cmd === "contribute") contribute(opts);
else if (cmd === "help") console.log(HELP);
else die(`unknown command "${cmd}"\n\n${HELP}`);
