# Onsite Report Consolidator

独立 Codex Skill，把明确指定的一次 onsite 审计 Excel 输出填入原 Checklist 模板，保留每个子项结果、Findings、Instruction、详情和导航。排除 GSC 第 7 系列；不负责审计、Drive 凭据、下载或上传。

输出主表名为 `(auto) SEO Checklist`。`Human check` 独立保留并加入 onsite 下拉选项；未检查使用 `-`。主表重命名会同步更新引用及返回链接，GSC 的值和下拉选项保持原样。

详情表只填原模板已有 schema，不新增 Source 栏位或底部汇总块。额外字段及来源追踪保存在 `.validation.json`。主表链接跳到对应明细起始位置，没有明细时跳到表头；详情表顶部返回对应主表主题行。每次重建这些导航链接并验证准确目标。

## 使用

同一套 Skill 支持 Windows / macOS，使用 Python 3.10+、openpyxl 和 Pillow。先在仓库外的 venv 安装 `requirements.txt`；Pillow 用于保留模板截图，缺少时会明确停止。两边完整的安装命令见 [跨平台使用说明](references/platforms.md)。不需要本机安装 Excel 或连接 Google。

```text
python -m pip install -r requirements.txt
python scripts/consolidate.py doctor
python scripts/consolidate.py inspect "本次指定的 ZIP 或 run folder"
python scripts/consolidate.py extract "本次指定的.zip" --destination "新的本地工作目录"
python scripts/consolidate.py build --manifest "本次 run folder/run.json"
python scripts/consolidate.py verify "本次 run folder/Autocheck_onsite_audit_Company_2026-10-15.xlsx"
```

在 Codex 中安装时，将本仓库内容放入个人 skills 目录下的 `onsite-report-consolidator/`，然后新会话使用 `$onsite-report-consolidator`。开发仓库本身不等于全局安装。详细输入和可扩展映射见 [mapping.md](references/mapping.md)，执行约束见 [SKILL.md](SKILL.md)。

最终文件 `Autocheck_onsite_audit_{companyname}_{date}.xlsx` 与本次源报告放在同一 run folder，另附 `.validation.json` 校验记录。同名停止，不覆盖。ZIP、源报告和原模板保持不变。旧文件名需要逐文件明确启用兼容模式。

目前包含 Robots、Sitemap、Pagination、Hreflang、Page titles、Meta descriptions、Headings 和 Images 的实际样本映射。17–20 支持合并和按子项拆分的来源详情表。HTTPS 目标表预留，首次真实输入需确认其 schema。未知格式由 agent 调整显式配置或增加测试过的 adapter；脚本不猜结果。

## 校验

```text
python -m unittest discover -s tests -v
```

测试及真实样本使用本地文件，不需要服务账号。新增测试覆盖中文／空格路径、Mac 文件名 Unicode 规范化、相对路径分隔符、独立解释器和截图保留。GitHub Actions 配置了 Windows／macOS × Python 3.10／3.12 的同一测试集；是否通过以实际运行结果为准。

样本中的客户数据、生成报告和 run manifest 不提交到仓库。模板原有 Summary 断引用、Google 导出公式和 heading/gid 链接会保留并单独记录，不能宣称整个工作簿无错误；原生 Excel 的重算／点击行为需另行验证。
