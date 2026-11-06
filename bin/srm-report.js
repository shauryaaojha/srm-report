#!/usr/bin/env node
// Installs the srm-report skill for Claude Code, and sends a user's improvements back as a pull request.
const { execFileSync } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");

const OWNER = "shauryaaojha";
const REPO = "srm-report";
const PKG_ROOT = path.join(__dirname, "..");
const SKILL_PATHS = ["SKILL.md", "references", "scripts", "assets"];
const DEFAULT_DIR = path.join(os.homedir(), ".claude", "skills", REPO);
// Reports and templates must never travel. Only text that improves the skill goes in a pull request.
const BLOCKED_FILE = /\.(docx?|dotx|pdf|pptx?|xlsx?|png|jpe?g|gif|zip|tgz)$/i;
const SKIP = /(^|[\\/])(__pycache__|\.git)([\\/]|$)|\.pyc$/;
const PERSONAL = [
  [/\bRA\d{10,}\b/, "SRM registration number"],
  [/[\w.+-]+@[\w-]+\.[\w.]{2,}/, "email address"],
  [/(?<![\d.])[6-9]\d{9}(?![\d.])/, "phone number"],
];

const HELP = `srm-report: SRM report skill for Claude Code

  npx srm-report install [--dir PATH] [--force]
      Copy the skill to ${DEFAULT_DIR}. --force replaces an existing copy.

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

function install(opts) {
  const dir = path.resolve(opts.dir || DEFAULT_DIR);
  if (fs.existsSync(path.join(dir, "SKILL.md")) && !opts.force) {
    die(`${dir} already has this skill, and may hold your own improvements.\nSend those first with "npx srm-report contribute", or pass --force to replace it.`);
  }
  fs.mkdirSync(dir, { recursive: true });
  copySkill(PKG_ROOT, dir);
  const { version } = JSON.parse(fs.readFileSync(path.join(PKG_ROOT, "package.json"), "utf8"));
  fs.writeFileSync(path.join(dir, "VERSION"), version + "\n");
  console.log(`Installed srm-report ${version} to ${dir}`);
  console.log("Claude Code picks it up in the next session.");
  console.log('If you improve it while making a report, send it back: npx srm-report contribute');
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
  const dir = path.resolve(opts.dir || DEFAULT_DIR);
  if (!fs.existsSync(path.join(dir, "SKILL.md"))) die(`no skill found at ${dir}`);
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
