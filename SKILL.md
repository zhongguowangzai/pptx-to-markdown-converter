---
name: pptx-to-markdown
version: 1.0.0
description: >
  将 PowerPoint (.pptx) 精确转换为 Markdown，保留层级结构、表格、流程图。
  支持空间排序、矩阵表格检测、大纲重建，解决普通转换工具的信息丢失问题。
  触发词：PPT转MD、PPT转Markdown、pptx to markdown、课件转文本、提取PPT内容、PPT转文字。
license: MIT
author: 旺仔
tags:
  - pptx
  - markdown
  - 文档转换
  - powerpoint
  - 课件
agent_created: true
---
# PPTX → Markdown 精确转换

将 PowerPoint 文件精确转换为结构化 Markdown，解决普通转换工具的信息丢失问题（流程图、层级、侧边栏标注）。

## 工作流程

### 第一步：准备环境（仅首次）

检查 `python-pptx` 和 `lxml` 是否已安装，未安装则自动安装：

```bash
pip show python-pptx > /dev/null 2>&1 || pip install python-pptx lxml
```

### 第二步：运行转换

对每个 PPTX 文件运行 v3 脚本。输出到工作空间根目录，文件名自动生成：

```bash
python "$SKILL_DIR/scripts/pptx_convert_v3.py" "输入PPTX路径" "输出MD路径"
```

多文件时并行运行以提高速度。每个文件 timeout 设为 60000ms，大文件（>10MB）可延长。

### 第三步：快速抽检（不做全量复核）

转换完成后，只看统计指标：页数、空白页数、表格数、代码块闭合情况。不逐页人工审查。

### 第四步：合并（如多文件）→ 交付

合并后交付所有文件，直接结束——不要等待用户确认或追问。

## 转换能力清单

> **v3 核心改进：** 空间排序（Y带+X左→右）→ 解决碎片化文本；矩阵检测 → 网格转表格；大纲重建 → 层级标签+描述（替代错乱Mermaid）；上下文层级感知 → 正确嵌套；顶部标题保留

| 能力 | 方法 | 输出形式 |
|------|------|----------|
| **标题识别** | placeholder.idx=0 精确定位 | `## 标题` |
| **子标题识别** | placeholder.idx=2 + 上下文 | `### 子标题` |
| **侧边栏** | 右边缘坐标 >12in | `> 📖 教材 XX页` |
| **底部标注** | y>6.0 + 高度<1in | `> 💡 总结` |
| **问句检测** | 以？/? 结尾 | `**问句内容**` |
| **段落层级** | paragraph.level 缩进 | `  - 二级内容` |
| **流程图重建** | 短文本非占位符形状 → Mermaid | `graph TD` 代码块 |
| **表格提取** | shape.has_table → GFM 表格 | Markdown 表格 |
| **矩阵表格** (v3) | 均匀行×列布局检测 | Markdown 表格 |
| **层级大纲** (v3) | 左标签+右描述结构识别 | 缩进大纲（`- **标签**` + 描述） |
| **空间排序** (v3) | Y带分组 + X左→右排序 | 有序正文流 |
| **幻灯片分界** | 每页自动分隔 | `<!-- Slide N -->` + `---` |

## 流程图检测规则

满足全部条件才识别为流程图节点：
1. 非占位符形状
2. 非连接线（shape_type ≠ 8）
3. 位于内容区（Y: 1.5in ~ 6.5in）
4. 宽度 ≥ 0.5in，高度 ≥ 0.3in
5. 仅含 1 段文本，且 ≤ 15 字符

重建逻辑：按 Y 坐标排序，最上方为根节点，其余为子节点，根节点到每个子节点建立 → 连接。

## 限制与回退

- **仅支持 .pptx**（Office 2007+ 格式），不支持 .ppt
- **SmartArt** 无法深度解析（Office 将 SmartArt 渲染为图片，XML 中不可见）
- **嵌入图片** 无法提取（如需图片，建议用 markitdown + LibreOffice 截图方案）
- 若用户 PPT 以图片/截图为主（非文字型课件），建议改用 markitdown
- 若脚本报错 `ModuleNotFoundError`，按第二步安装缺失的包

## 脚本维护

- 当前主力脚本：`scripts/pptx_convert_v3.py`（推荐）
- 旧版兜底脚本：`scripts/pptx_deep_convert.py`（简单布局仍可用）
- v3 架构：`SlideParser` 类 → `parse()` → `_process_text_shapes()` → 三个检测分支（大纲/矩阵/流程图）
- 关键参数：
  - `_detect_outline_structure()`：left < 4.0 且短文本比例 > 20% + 右侧长描述 ≥ 2 → 大纲模式
  - `_build_outline()`：0.5in Y 带分组，4 级左坐标阈值（1.5/3.0/5.5）
  - `_check_uniform_columns()`：列数一致的矩阵 → 表格
- 若幻灯片被误判为不当模式，调整对应检测函数的阈值

---

## 依赖

- Python 3.8+
- python-pptx
- lxml

---

发布页：GitHub 仓库直链安装，兼容 ClawHub / WorkBuddy / OpenClaw 技能生态。
