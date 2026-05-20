---
name: coa-msds
description: >
  COA和MSDS文件处理统一技能。涵盖：COA完整一键生成、MSDS一键生成（含成分复制+超链接+含量检查）、
  CAS号填写、保质期填写。
  触发词：COA、MSDS、成分表、CAS号、保质期。
tools: Bash, Read, Write, Edit, Grep, Glob
---

# COA / MSDS 统一处理

## 核心规则

- **操作前先说明计划，等用户同意再执行**
- **多文件时先做一个确认**：第一个做完给用户验证，通过后再批量处理其余
- **展示步骤菜单**：列出所有可用步骤，让用户选择执行哪一步
- **编辑不改样式**：铁律 — `.xls` 用 COM，`.xlsx` 用 openpyxl，只写 `.value`，不动样式属性
  - `.xls` 必须用 COM（Python win32com），禁止 xlutils/xlwt 写操作
  - `.xlsx` 用 openpyxl，只写 `.value`，不动 `.font/.border/.fill/.alignment`
  - 编辑后验证：对比未编辑区域的样式是否一致
- **品牌品名格式**：ICE LERSKIN / CHIC.PEAK
  - 中文：英文品牌 + 空格 + 中文品名（品牌粘连自动补空格）
  - 英文：全大写 + 粘连拆分（`GLAZEALL` → `GLAZE ALL`）+ 品牌前缀补全
  - 品牌统一：`CHICPEAK`/`CHIC PEAK` → `CHIC.PEAK`
  - 品名不含规格（ml/g/片等），规格在独立列
- **保质期默认填 3 年**，不用每次确认
- **居中规则**：
  - COA：全部数据区域（Row 2 到末行）水平+垂直居中
  - MSDS：仅 Row 2-11 中有值的单元格居中，Row 12-15 维持模板原样

## 路径配置

| 路径 | 用途 |
|------|------|
| `templates/COA模板.xls` | COA 输出模板 (.xls) |
| `templates/MSDS模板.xlsx` | MSDS 输出模板 (.xlsx) |
| `resources/新建文本文档.txt` | 委托单位、联系人信息 |
| `resources/mfr_en.json` | 生产单位英文翻译表 |
| `resources/cp_color_map.json` | CP 颜色/性状映射表（53条） |
| `resources/COA基准/` | COA 填好的参考文件，用于最终校对 |
| `resources/MSDS基准/` | MSDS 填好的参考文件，用于最终校对 |
| `scripts/` | 所有处理脚本 |

> 输入数据（成分 `.xlsm` + 生产方对照表 `.xlsx`）由用户指定目录，不在 skill 内。

## 文件结构

```
coa-msds/
├── SKILL.md                           # 本文件：指令和工作流
├── templates/
│   ├── COA模板.xls                    # COA 输出模板
│   └── MSDS模板.xlsx                  # MSDS 输出模板
├── resources/
│   ├── 新建文本文档.txt                # 委托单位/联系人信息
│   ├── mfr_en.json                    # 生产单位英文翻译表
│   ├── cp_color_map.json              # CP 颜色/性状映射表
│   ├── COA基准/                       # COA 参考示例（校对用）
│   │   └── 多款-COA-出报告信息确认表.xls
│   └── MSDS基准/                      # MSDS 参考示例（校对用）
│       ├── MSDS_CP_韶关市古晴化妆品有限公司.xlsx
│       └── MSDS_ICE_广州肌因未来国际生物科技有限公司.xlsx
└── scripts/
    ├── coa_fill.py                    # COA 一键生成（含匹配+规格+性状+居中+保质期）
    ├── msds_fill.py                   # MSDS 一键生成（含分组+性状+成分复制+超链接+含量检查+居中）
    ├── fill_cas.py                    # CAS 号填写
    ├── fill_shelf_life.py             # 保质期填写
    ├── build_cp_color_map.py          # CP 色号映射表生成
    ├── verify_coa.py                  # COA 输出验证
    └── verify_msds_format.py          # MSDS 格式验证
```

## 工作流总览

```
COA 流程（.xls 必须用 COM）
  ① python coa_fill.py <数据目录> <输出路径>
     → 含：产品匹配 + 规格提取 + 性状推断 + 居中 + 保质期

MSDS 流程（.xlsx 用 openpyxl）
  ② python msds_fill.py <数据目录> <输出目录> [语言]
     → 含：分组 + 性状推断 + 成分复制 + Sheet1超链接 + 含量检查 + 居中

  ③ python fill_cas.py <file_or_dir>
     → CAS 号填充（详见 CAS 填充规则）

  ④ python fill_shelf_life.py <file> <column>
     → 保质期填写（默认 3 年）

验证
  ⑤ python verify_coa.py <基准> <输出>
  ⑥ python verify_msds_format.py <基准目录> <输出目录>
```

---

## 步骤 ①：COA 完整生成

**脚本：** `scripts/coa_fill.py`

### 场景

根据成分文件和生成方对照表，一键生成 COA 确认表。输出为 .xls（ICE + CP 两个 sheet）。

### 需要的数据源

| 数据源 | 路径 | 内容 |
|--------|------|------|
| 品牌成分文件夹 | 用户指定目录/CP成分/, ICE成分/ | 每个产品一个 .xlsm，A1=英文品名，sheet 内含成分表 |
| 生产方对照表 | 用户指定目录/厦门外贸商品对应生产方.xlsx | 列：商品编号、商品名称、生产单位、生产单位地址 |
| COA 模板 | templates/COA模板.xls | 输出模板，含表头、示例行、合并单元格和边框样式 |
| 委托/联系人信息 | resources/新建文本文档.txt | 委托单位中英文 + 地址 + 联系人 + 电话 + 邮箱 |
| COA 基准参考 | resources/COA基准/ | 填好的成品，用于生成后校对 |

### 处理流程

1. **扫描成分文件**：遍历品牌文件夹 .xlsm，读 sheet A1 提取英文品名。跳过 phantom sheet
2. **匹配生产方**：英文品名 × 生产方中文品名做关键词打分。CP 产品用类目关键词 + 产品码匹配
3. **推断性状**：从成分表分析色素(CI码)和香精推断物理状态、颜色、气味
4. **COM 填充模板**：先复制模板 → CP sheet，逐个写 `.Value`，不改格式。数据超出模板行数时延伸边框
5. **统一格式**：字体宋体、字号自动检测、颜色统一、行高统一、全部居中

### 性状推断规则

**物理状态** — 从品名关键词推断：

| 关键词 | 状态 | 关键词 | 状态 |
|--------|------|--------|------|
| TONER / LOTION / WATER | 液体 | SERUM | 精华 |
| SPRAY / AEROSOL | 气雾 | STICK / BALM | 棒状 |
| CREAM | 乳霜 | MOUSSE / CLEANSER | 慕斯 |
| GEL | 凝胶 | MASK / MUD | 膏状 |
| LIP GLAZE | 粘稠液体 | LIP MUD / BLUSH MUD | 泥状 |
| LIQUID BLUSH | 液体 | CUSHION | 液状 |
| SETTING COMPACT | 粉状 | OIL / ELIXIR | 液体 |

CP 产品颜色/性状从 `resources/cp_color_map.json` 查得（53条）。

**气味** — 从成分功能列检测：
- 含香精/FRAGRANCE → 稍有气味
- 不含香精 → 无
- 含特征香料原料 → 特征性气味
- CP 品牌彩妆统一"稍有气味"

### 关键坑

| 坑 | 说明 |
|----|------|
| phantom LIP STICK sheet | 所有 .xlsm 首 sheet 是共享模板残留，必须跳过 |
| .xls 必须用 COM | `xlutils.copy` + `xlwt.write` 会破坏格式 |
| COM Copy 语法 | `ws.Copy(After=ws)` 不生效，须用 `ws.Copy(None, ws)` |
| LIP STICK A1 相同 | 6 个文件 A1 相同，需从文件名正则提取真实变体 |
| 模板行数不足 | 模板通常 ~12 行，产品多时需延伸边框 |
| 模板示例残留 | 第一行数据有示例的"生产日期/批号/质检员"，必须清空 |

### 操作流程

1. 用户指定成分文件夹和生产方对照表路径
2. 说明产品数，等用户确认
3. 运行 `python scripts/coa_fill.py <数据目录> <输出路径>`
4. 对照 `resources/COA基准/` 校验收到的输出

---

## 步骤 ②：MSDS 一键生成

**脚本：** `scripts/msds_fill.py`

### 场景

按品牌→生产单位分组，每组生成一个 MSDS .xlsx 文件。填写 Sheet1 基本资料，并从源 .xlsm 整体复制成分表到各成分 sheet。

### 需要的数据源

| 数据源 | 路径 | 内容 |
|--------|------|------|
| MSDS 模板 | templates/MSDS模板.xlsx | 含 Sheet1 + 5 个成分表 sheet 占位 |
| 品牌成分文件夹 | 用户指定目录/CP成分/, ICE成分/ | 每个产品一个 .xlsm 文件 |
| 生产方对照表 | 用户指定目录/厦门外贸商品对应生产方.xlsx | 商品编号、商品名称、生产单位、地址 |
| 联系人/委托信息 | resources/新建文本文档.txt | 委托单位、地址、联系人、电话、邮箱 |
| 生产单位英文翻译 | resources/mfr_en.json | 中文厂名→英文厂名 |
| MSDS 基准参考 | resources/MSDS基准/ | 填好的成品，用于生成后校对 |

### 处理流程

1. **分组**：扫描成分文件，通过 COA 同款 `match_mfr` 匹配生产方（评分 ≥ 3），按品牌|生产单位分组
2. **品名规范化**：中文去规格/版本描述，英文全大写+粘连拆分+品牌前缀补全
3. **Sheet1 填写**：Row 2-7, 9, 11 基本信息，Row 12-15 产品信息
4. **性状推断**：与 COA 使用完全相同的推断逻辑
5. **Sheet 管理**：产品 < 5 删除多余 sheet，> 5 复制 sheet + 延伸合并单元格
6. **成分表复制**：从源 .xlsm 整页复制（A1 匹配 → 解合并 → 清空 → 复制数据+样式+行高+列宽+合并）
7. **公式自动修正**：扫描 D3/F3 `=100-SUM` 公式，按实际数据末行校准范围
8. **Sheet1 超链接**：Row 13 产品名自动链接到对应成分表 sheet
9. **居中**：Row 2-11 有值单元格自动水平+垂直居中

### 填写规则速查

| Row | 标签 | 填写方式 | 语言 |
|-----|------|----------|------|
| 2 | 委托单位 | DELEGATE_CN / DELEGATE_EN | 双语 |
| 3 | 委托单位地址 | DELEGATE_ADDR_CN / DELEGATE_ADDR_EN | 双语 |
| 4 | 生产单位 | mfr_cn / mfr_en | 双语 |
| 5 | 生产单位地址 | addr_cn / addr_en（含国家前缀） | 双语 |
| 6 | 联系人 | 中文姓名 | 单语言 |
| 7 | 电话 | 号码 | — |
| 8 | 紧急联系话语 | **不填** | — |
| 9 | 邮箱 | email | — |
| 10 | 品牌准确名称 | **不填** | — |
| 11 | 语言 | **每次问用户** | — |
| 12 | 页码标签 | 第一页/第二页... | 中文 |
| 13 | 产品名称 | cn / en | 双语 |
| 14 | 物理状态 | 英文 | 单语言 |
| 15 | 气味 | 英文 | 单语言 |

### 关键坑

| 坑 | 说明 |
|----|------|
| Sheet 名不能改 | sheet 名 `第N款成分` 保持不变 |
| 地址加国家 | 中文地址前加 `中国`，英文地址末尾加 `, China` |
| 成分表 A1 匹配 | 源 .xlsm 可能含 phantom 首 sheet，必须 A1 匹配 |
| 表头不覆盖 | copy_ingredients 后不再写 Row 1-2 |
| 合并单元格延伸 | 产品 > 5 时，Row 1-11 合并区域从 K 列延伸到 (2+n) 列 |
| 公式保留 | `copy_ingredients` 必须用 `data_only=False` |
| 性状推断与COA统一 | `match_mfr` 从 coa_fill 导入 |

### 操作流程

1. 确认成分文件夹路径和生产方对照表
2. 问用户 Row 11 语言设置
3. 运行 `python scripts/msds_fill.py <数据目录> <输出目录> [语言]`
4. 运行 `python scripts/fill_cas.py <输出目录>` 补填 CAS

---

## 步骤 ③：CAS 号填写

**脚本：** `scripts/fill_cas.py`

### 场景

MSDS 生成后，H 列（CAS 原料报送码）空白按规则自动填充。

### 填充规则

| 优先级 | 规则 | 处理方式 |
|--------|------|----------|
| 1 | 已有值（CAS 号 或 `XXXXXX-XXXXX-XXXX` 报送码） | **不动** |
| 2 | `/` 或 `-`（不适用标记） | **不动** |
| 3 | H 列合并且首行已有 CAS | **不拆不填**（厂商内部编码） |
| 4 | H 列合并且首行空白 | 拆开合并，逐行填 |
| 5 | 水 / AQUA / WATER | **跳过** |
| 6 | 香精 / FRAGRANCE / AROMA / PARFUM | 填 `Mixture` |
| 7 | 无 CAS 的聚合物/混合物 | 跳过 |
| 8 | 其余空白 | 从映射表匹配填 |

### 数据来源

- 脚本内置 `CAS_MAP` 字典，覆盖 140+ 化妆品常见成分的 INCI 名称 → CAS 编号

### 用法

```bash
python scripts/fill_cas.py "<file_or_dir>"              # 处理单个文件或递归目录
python scripts/fill_cas.py "<file_or_dir>" --dry-run    # 预览不保存
```

### 关键坑

| 坑 | 说明 |
|----|------|
| 报送码不是 CAS | `004183-01464-9628` 格式是化妆品原料报送码，必须跳过 |
| H 列合并保护 | H 列合并且首行已有编码时不拆不填 |
| 水不填 | 水虽有 CAS `7732-18-5`，但 MSDS 惯例不填 |

---

## 步骤 ④：保质期填写

**脚本：** `scripts/fill_shelf_life.py`

默认填 "3年"，支持 .xls（COM）和 .xlsx（openpyxl）。

### 用法

```bash
python scripts/fill_shelf_life.py "<file>" "<column>"                     # 默认填 3年
python scripts/fill_shelf_life.py "<file>" "<column>" --value "2年"       # 自定义
python scripts/fill_shelf_life.py "<file>" "<column>" --dry-run           # 预览
```

---

## 步骤 ⑤：生成后校对

COA 和 MSDS 生成完成后，对照 `resources/COA基准/` 或 `resources/MSDS基准/` 验证。

### 校对内容

| 校对项 | 检查方法 | 标准 |
|--------|----------|------|
| 表格样式 | 对比边框、填充色、对齐 | 与基准一致 |
| 字体字号 | 检查数据区域字体 | 宋体，与模板一致 |
| 数据位置 | 确认行列对应 | 与基准一致 |
| 合并单元格 | 检查合并区域延伸 | 与产品数量匹配 |
| 超链接 | Ctrl+Click Row 13 | 跳转到对应成分表 sheet |
