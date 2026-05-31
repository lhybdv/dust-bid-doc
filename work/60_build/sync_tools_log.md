# sync_tools_log.md

## 同步时间
2026/05/31

## 源路径
`/Users/liuhuiyu/.claude/plugins/cache/fengyun-claude-marketplace/tender/0.3.31/tools/tenderctl/`

## 目标路径
`/Users/liuhuiyu/projects/dust-bid-doc/tools/tenderctl/`

## 同步详情

| 文件 | 操作 | 原因 |
|------|------|------|
| `tenderctl.py` | COPY | 旧版本 112.8K → 新版本 172.9K (缺少 `normalize-writing-plan` 和 `--section-briefs-dir` 的早期版本) |
| `requirements.txt` | SKIP | identical |

## 验证结果

| 检查项 | 输出 |
|--------|------|
| `python tenderctl.py --help` | PASS — 包含 `normalize-writing-plan` |
| `python tenderctl.py claim-section --help` | PASS — 包含 `--section-briefs-dir` |
| `python tenderctl.py normalize-writing-plan --help` | PASS — 子命令正常 |

## 最终结果
**PASS**

> 注：从插件缓存目录 `/Users/liuhuiyu/.claude/plugins/cache/fengyun-claude-marketplace/tender/0.3.31/tools/tenderctl/` 同步到项目 `tools/tenderctl/`。旧版本缺失 `normalize-writing-plan` 子命令，已被覆盖。新版本包含 `claim-section` 的 `--section-briefs-dir` 选项。