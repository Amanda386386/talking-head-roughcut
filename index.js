import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const skillDir = join(here, "skills", "talking-head-roughcut");
const skillFile = join(skillDir, "SKILL.md");

function parseSkill() {
  const raw = readFileSync(skillFile, "utf8");
  const m = /^---\r?\n([\s\S]*?)\r?\n---\r?\n?([\s\S]*)$/.exec(raw);
  if (!m) throw new Error("SKILL.md 缺少 YAML frontmatter");
  const fm = m[1];
  const name = (/^name:\s*(.+)$/m.exec(fm) || [])[1]?.trim();
  const description = (/^description:\s*(.+)$/m.exec(fm) || [])[1]?.trim();
  if (!name || !description) throw new Error("SKILL.md frontmatter 缺少 name/description");
  return { name, description, content: m[2].trim() };
}

export const inject = ["skills"];

export function apply(ctx) {
  const { name, description, content } = parseSkill();
  ctx.effect(() =>
    ctx.skills.register({
      name,
      description,
      content,
      source: "plugin",
      path: skillFile,
      resourceBase: { kind: "directory", path: skillDir },
    })
  );
}
