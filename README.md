# 口播粗剪（talking-head-roughcut）

单机位口播视频的自动粗剪技能 + DeepSeek Harness 插件包。把原素材与口播稿对齐校准，自动去重录、删卡壳、清长气口，按原分辨率、原码率输出粗剪成片。

只做粗剪：不加字幕、花字、调色、BGM、音效。

## 能力

- 同一段台词连录多次，只保留最后一次完整录制；前面说一半又重说的，前面那半句整段删掉。
- 剪掉卡壳和 0.5 秒以上的长气口；任何 3 秒以上的静默必剪。
- 句尾语气词（呢/啊/了/的）和「左右」「大概」这类模糊限定词完整保留，不把「1000-2000左右」剪成确切的「1000-2000」。
- 切点交替换景别（1.0/1.09）遮掉单机位硬切的跳帧，且不剪到字。
- 按原分辨率、原帧率、原码率输出。

## 目录结构

```
.
├── index.js                  # 插件入口：注册 skill
├── package.json              # 插件包清单（dsh.bundle.patch）
├── cordis.patch.yml          # 插入插件行的补丁
├── locale/ icon.svg docs/    # 插件显示元数据、图标、详情页
└── skills/talking-head-roughcut/
    ├── SKILL.md              # 技能说明（五条铁律 + 流程 + 故障排查）
    └── scripts/
        ├── roughcut.py       # 主流程：对齐、去重录、切气口、渲染
        ├── make_subtitles.py # 可选：字幕（默认不加）
        └── environment.json  # 本机 ffmpeg / ffprobe / 模型 / python 路径
```

`skills/talking-head-roughcut/` 是一个**可独立搬运的技能文件夹**：任何支持 `SKILL.md` 约定的 Agent（DeepSeek Harness、Claude、Cursor 等）都能直接把它丢进自己的 skills 目录使用；外面的 `index.js` 等只是 DSH 的插件包装，让它在 DSH 里出现在「已安装插件」并可启停/卸载。

## 前置依赖

技能运行需要本机具备：

- `ffmpeg` + `ffprobe`
- `faster-whisper` 的 `large-v3` 模型（词级时间戳）
- 一个装好 `faster-whisper`、`av`、`numpy` 的 Python 环境

三者的绝对路径写在 `skills/talking-head-roughcut/scripts/environment.json` 里；也可以不改文件，改用环境变量 `KOU_BO_FFMPEG` / `KOU_BO_FFPROBE` / `KOU_BO_MODEL` 覆盖。

> 注意：`environment.json` 里现在是作者本机的路径，别的机器直接装会找不到 ffmpeg/模型，需要先改成自己机器的路径。

## 安装进 DeepSeek Harness

### 方式 A：作为插件包（推荐，可启停/卸载）

在 DSH 里让 Agent 执行，或直接在「插件 → 添加插件」里填仓库地址：

```text
install_bundle https://github.com/<你的用户名>/talking-head-roughcut.git
```

装完在「已安装插件」里能看到「口播粗剪」，技能名 `talking-head-roughcut` 自动注册。

也支持克隆到本地后从目录安装：

```powershell
git clone https://github.com/<你的用户名>/talking-head-roughcut.git
# 然后在 DSH 插件页选「从本地目录添加」，指向克隆出的文件夹
```

### 方式 B：只当普通技能用（不用插件包装）

把 `skills/talking-head-roughcut/` 整个文件夹复制到你的技能目录即可：

- DeepSeek Harness：`~/.dsh/skills/`（或项目里 `.dsh/skills/`）
- 其他遵循 SKILL.md 约定的 Agent：其对应的 skills 目录

## 使用

> 用 talking-head-roughcut 粗剪这条视频，视频在〈路径〉，文案在〈路径〉。

需要准备一个视频文件（.mov/.mp4，最好一整条长录）和一份口播稿（.docx/.txt/.md，脚本默认取最长正文段）。

产出在视频同级生成 `<视频名>_粗剪\`：成片 mp4、`剪点表.csv`、`粗剪报告.md`、`plan.json`。

## 修改与回滚

- 直接改仓库里的 `SKILL.md`（规则）或 `scripts/roughcut.py`（算法），提交后重新安装或等 HMR 生效。
- 常用参数：`--gap-cut`（气口阈值，默认 0.5）、`--keep-sentence`（断句留白，默认 0.12）、`--speed`（整体倍速）、`--no-zoom`（纯硬切）。
- 每次发现新问题，把「哪一分钟、什么现象」记进 `SKILL.md` 的故障排查表，规则就不会再犯。
