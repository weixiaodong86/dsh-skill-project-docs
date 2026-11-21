/**
 * DSH skill bundle: 软件研发全流程项目管理过程文档。
 *
 * 把技能以 packaged provider 的形式注册到 ctx.skills，
 * 技能正文与资源目录随 npm 包一起分发，安装后出现在会话技能目录中。
 *
 * @module dsh-skill-project-docs
 */
import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'

/** 与 @deepseek-ai/dsh-skill 的 BUNDLED_SKILL_RANK 一致：打包技能的优先级。 */
const SKILL_RANK = 600
const PROVIDER_NAME = 'software-project-docs-bundle'
const SKILL_DIR_URL = new URL('./skills/software-project-docs/', import.meta.url)
const SKILL_BODY_URL = new URL('SKILL.md', SKILL_DIR_URL)
const RESOURCE_BASE = {
  kind: 'directory',
  path: fileURLToPath(SKILL_DIR_URL),
}
const FALLBACK_DESCRIPTION =
  '生成与修订软件研发全流程项目管理过程文档（DOCX）。先交互确认文档清单，再依据项目资料按统一排版规范生成正式文档；模板模式附斜体填写说明与示例；修订按版本号出新文件并维护修订记录。'

/** 拆分 YAML frontmatter：只处理 `key: value` 形式的顶层标量。 */
function splitFrontMatter(raw) {
  if (!raw.startsWith('---')) return { meta: {}, body: raw }
  const end = raw.indexOf('\n---', 3)
  if (end < 0) return { meta: {}, body: raw }
  const meta = {}
  for (const line of raw.slice(3, end).split(/\r?\n/)) {
    if (!line || line.startsWith(' ') || line.startsWith('\t') || line.startsWith('#')) continue
    const at = line.indexOf(':')
    if (at < 0) continue
    const key = line.slice(0, at).trim()
    const value = line.slice(at + 1).trim().replace(/^["']|["']$/g, '')
    if (key) meta[key] = value
  }
  return { meta, body: raw.slice(end + 4).replace(/^\r?\n/, '') }
}

let cache

async function load() {
  if (cache) return cache
  const raw = await readFile(SKILL_BODY_URL, 'utf8')
  const { meta, body } = splitFrontMatter(raw)
  cache = {
    name: typeof meta.name === 'string' && meta.name ? meta.name : 'software-project-docs',
    description: typeof meta.description === 'string' && meta.description ? meta.description : FALLBACK_DESCRIPTION,
    whenToUse: typeof meta.whenToUse === 'string' && meta.whenToUse ? meta.whenToUse : undefined,
    body,
  }
  return cache
}

const invocation = { modelInvocable: true, userInvocable: true }

const provider = {
  name: PROVIDER_NAME,
  async list() {
    const skill = await load()
    return [summaryOf(skill)]
  },
  async get() {
    const skill = await load()
    return {
      name: skill.name,
      description: skill.description,
      ...(skill.whenToUse ? { whenToUse: skill.whenToUse } : {}),
      invocation,
      provider: PROVIDER_NAME,
      source: 'bundled',
      resourceBase: RESOURCE_BASE,
      content: skill.body,
    }
  },
}

function summaryOf(skill) {
  return {
    name: skill.name,
    description: skill.description,
    ...(skill.whenToUse ? { whenToUse: skill.whenToUse } : {}),
    invocation,
    provider: PROVIDER_NAME,
    source: 'bundled',
    resourceBase: RESOURCE_BASE,
    rank: SKILL_RANK,
    locator: SKILL_BODY_URL,
  }
}

/** Cordis 插件名。 */
export const name = 'skill-software-project-docs'
/** 依赖的技能注册表服务。 */
export const inject = ['skills']
/** 注册打包技能 provider。 */
export function apply(ctx) {
  ctx.skills.registerProvider(() => provider)
}
