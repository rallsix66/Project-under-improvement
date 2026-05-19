---
name: coa-msds
description: >
  COA和MSDS文件处理统一技能。涵盖：COA完整一键生成、MSDS成分表替换、
  Sheet1产品名超链接、成分表含量合计检查、产品名拆分、CAS号填写、保质期填写。
  触发词：COA、MSDS、成分表、确认表、超链接、含量检查、百分比检查、COA生成、MSDS替换、保质期、产品名拆分、规格分离。
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

## 路径配置

| 路径 | 用途 |
|------|------|
| `templates/COA模板.xls` | COA 输出模板 (.xls) |
| `templates/MSDS模板.xlsx` | MSDS 输出模板 (.xlsx) |
| `resources/新建文本文档.txt` | 委托单位、联系人信息 |
| `resources/mfr_en.json` | 生产单位英文翻译表 |
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
│   ├── COA基准/                       # COA 参考示例（校对用）
│   │   └── 多款-COA-出报告信息确认表.xls
│   └── MSDS基准/                      # MSDS 参考示例（校对用）
│       ├── MSDS_CP_韶关市古晴化妆品有限公司.xlsx
│       └── MSDS_ICE_广州肌因未来国际生物科技有限公司.xlsx
└── scripts/
    ├── coa_fill.py                 # ① COA完整一键生成
    ├── msds_fill.py                # ② MSDS一键生成
    ├── verify_coa.py                  # ① COA输出验证
    ├── verify_msds_format.py          # ② MSDS格式验证
    ├── replace_msds.py                # ② MSDS成分表替换
    ├── extract_shape_color_scent.py   # ①② 性状颜色气味推断（从成分推断，不依赖外部性状表）
    ├── extract_chic_peak.py           # ①② CHIC.PEAK品牌提取
    ├── match_specs_v3.py              # ① 产品匹配+规格提取
    ├── fill_cas.py                    # ⑥ CAS号填写
    ├── hyperlink.py                   # ③ Sheet1超链接
    ├── check_percent.py               # ④ 含量合计检查
    ├── split_spec.py                  # ⑤ 产品名拆分
    └── fill_shelf_life.py             # ⑦ 保质期填写
```

## 工作流总览

```
COA 流程（.xls 必须用 COM，.xlsm/.xlsx 用 openpyxl）
  ① COA 完整生成 → scripts/coa_fill.py  一键生成（含产品匹配+规格提取+性状推断+保质期）

MSDS 流程（.xlsx/.xlsm 用 openpyxl）
  ② MSDS 一键生成 → scripts/msds_fill.py  Sheet1基本资料 + 成分表复制（按品牌→生产单位分组）
  ③ Sheet1超链接    → scripts/hyperlink.py    Row 13 产品名链接到对应成分表 sheet
  ④ 含量合计检查    → scripts/check_percent.py D/F 列求和 ≈ 100% 检查

通用
  ⑤ 产品名拆分   → scripts/split_spec.py     中文品名 → 产品名 + 规格
  ⑥ CAS号填写    → scripts/fill_cas.py
  ⑦ 保质期填写   → scripts/fill_shelf_life.py  默认填 "3年"
```

---

## 步骤 ①：COA 完整生成

**脚本：** `scripts/coa_fill.py`

### 场景

根据成分文件和生成方对照表，一键生成 COA 确认表。输出为 .xls（ICE + CP 两个 sheet + 备案表）。

### 需要的数据源

| 数据源 | 路径 | 内容 |
|--------|------|------|
| 品牌成分文件夹 | 用户指定目录/CP成分/, ICE成分/ | 每个产品一个 .xlsm，A1=英文品名，sheet 内含成分表 |
| 生产方对照表 | 用户指定目录/厦门外贸商品对应生产方.xlsx | 列：商品编号、商品名称、生产单位、生产单位地址 |
| COA 模板 | templates/COA模板.xls | 输出模板，含表头、示例行、合并单元格和边框样式 |
| 委托/联系人信息 | resources/新建文本文档.txt | 委托单位中英文 + 地址 + 联系人 + 电话 + 邮箱 |
| COA 基准参考 | resources/COA基准/ | 填好的成品，用于生成后校对 |

### 处理流程

1. **扫描成分文件**：遍历品牌文件夹 .xlsm，读 sheet A1 提取英文品名。跳过 phantom sheet（首 sheet 可能是 LIP STICK 共享模板残留）
2. **匹配生产方**：英文品名 × 生产方中文品名做关键词打分。CP 产品用类目关键词 + 产品码匹配
3. **推断性状**：从成分表分析色素(CI码)和香精推断物理状态、颜色、气味（详见下方规则）。不再依赖外部性状表。
4. **COM 填充模板**：先复制模板 → CP sheet，逐个写 `.Value`，不改格式。数据超出模板行数时延伸边框

### 性状推断规则（AI/关键词，无外部依赖）

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

**颜色** — 从 CI 色号推断：

| CI 码 | 色系 | CI 码 | 色系 |
|-------|------|-------|------|
| CI15850 / CI45410 / CI73360 / CI17200 | 红色系 | CI15985 / CI45380 | 橙色系 |
| CI19140 / CI47005 | 黄色系 | CI42090 / CI77007 | 蓝色系 |
| CI60725 / CI77742 | 紫色系 | CI77491/2/9 组合 | 棕色系 |
| 仅 CI77891 | 白色 | ≥2 种色系 | 多色 |

特殊产品色号从英文名关键词匹配（如 BANANA→淡黄色, GRAPE→淡紫色, PEACH→蜜桃色）。

**气味** — 从成分功能列检测：

| 成分 | 气味 |
|------|------|
| 含香精/FRAGRANCE | 稍有气味 |
| 不含香精 | 无 |
| 含特征香料原料 | 特征性气味 |

CP 品牌彩妆统一"稍有气味"。

### 关键坑

| 坑 | 说明 |
|----|------|
| phantom LIP STICK sheet | 所有 .xlsm 首 sheet 是共享模板残留，必须跳过（除非文件名本身是 LIP STICK 产品） |
| .xls 必须用 COM | `xlutils.copy` + `xlwt.write` 会破坏格式，必须用 `win32com.client` |
| COM Copy 语法 | `ws.Copy(After=ws)` 不生效，须用位置参数 `ws.Copy(None, ws)` |
| LIP STICK A1 相同 | 6 个文件 A1 相同，需从文件名正则提取真实变体 |
| 模板行数不足 | 模板通常 ~12 行，产品多时需延伸边框 |
| 模板示例残留 | 第一行数据有示例的"生产日期/批号/质检员"，必须清空 |
| 变体匹配 | 液体腮红/气垫/LIP STICK 有多变体，需额外加分 |
| 字号须检测 | 模板字号可能不统一（9pt + 11pt 混用），需从多数检测 |

### 操作流程

1. 用户指定成分文件夹和生产方对照表路径
2. 说明匹配逻辑和预期产品数，等用户确认
3. 运行 `python scripts/coa_fill.py <数据目录> <输出路径>`
4. 对照 `resources/COA基准/` 校验收到的输出样式和数据

---

## 步骤 ②：MSDS 一键生成

**脚本：** `scripts/msds_fill.py`

### 场景

按品牌→生产单位分组，每组生成一个 MSDS .xlsx 文件。填写 Sheet1 基本资料，并从源 .xlsm 整体复制成分表到各成分 sheet（含样式、行高、列宽、合并单元格）。

### 需要的数据源

| 数据源 | 路径 | 内容 |
|--------|------|------|
| MSDS 模板 | templates/MSDS模板.xlsx | 含 Sheet1 + 5 个成分表 sheet 占位 |
| 品牌成分文件夹 | 用户指定目录/CP成分/, ICE成分/ | 每个产品一个 .xlsm 文件 |
| 分组数据 | .json（脚本生成） | 品牌→生产单位分组 |
| 联系人/委托信息 | resources/新建文本文档.txt | 委托单位、地址、联系人、电话、邮箱 |
| 生产单位英文翻译 | resources/mfr_en.json | 中文厂名→英文厂名 |
| MSDS 基准参考 | resources/MSDS基准/ | 填好的成品，用于生成后校对 |

### 处理流程

1. **分组**：扫描成分文件 A1 英文品名，通过 COA 同款 `match_mfr` 匹配生产方对照表（评分 ≥ 3），按品牌|生产单位分组
2. **品名规范化**：中文去规格/版本描述，英文全大写+粘连拆分+品牌前缀补全
3. **Sheet1 填写**：Row 2-7, 9 基本信息，Row 11 语言设置，Row 12-15 产品信息
4. **性状推断**：与 COA 使用完全相同的推断逻辑（AI/关键词，无外部性状表依赖）
5. **Sheet 管理**：产品 < 5 删除多余 sheet，> 5 复制 sheet + 延伸合并单元格
6. **成分表复制**：从源 .xlsm 整页复制（A1 品名匹配 → 解合并 → 清空 → 复制数据+样式+行高+列宽+合并 → 清多余行）
7. **公式自动修正**：扫描 D3/F3 `=100-SUM` 公式，按实际数据末行校准范围
8. **Sheet1 超链接**：Row 13 产品名自动链接到对应成分表 sheet

### 填写规则速查

| Row | 标签 | 填写方式 | 语言 |
|-----|------|----------|------|
| 2 | 委托单位（名称-中文/英文） | DELEGATE_CN / DELEGATE_EN | 双语 |
| 3 | 委托单位地址（中文/英文） | DELEGATE_ADDR_CN / DELEGATE_ADDR_EN | 双语 |
| 4 | 生产单位（名称-中文/英文） | mfr_cn / mfr_en | 双语 |
| 5 | 生产单位地址（中文/英文） | addr_cn / addr_en（含国家前缀） | 双语 |
| 6 | 联系人 | 中文姓名 | 单语言 |
| 7 | 电话 | 号码 | — |
| 8 | 紧急联系话语 | **不填** | — |
| 9 | 邮箱 | email | — |
| 10 | 品牌准确名称 | **不填** | — |
| 11 | 语言 | **每次问用户** | — |
| 12 | 页码标签 | 第一页/第二页... | 中文 |
| 13 | 产品名称（中文/英文） | cn / en | 双语 |
| 14 | 物理状态 | 英文（跟随 Row 11） | 单语言 |
| 15 | 气味 | 英文（跟随 Row 11） | 单语言 |

### 生产单位英文翻译

见 `resources/mfr_en.json`，当前已翻译 11 个生产单位。新增生产单位时需手动添加。

### 关键坑

| 坑 | 说明 |
|----|------|
| Sheet 名不能改 | sheet 名 `第N款成分` 保持不变，Row 12 用 `第N页` 格式 |
| 地址加国家 | 中文地址前加 `中国`，英文地址末尾加 `, China` |
| 成分表 A1 匹配 | 源 .xlsm 可能含 phantom 首 sheet，必须用 A1 匹配找正确 sheet |
| 表头不覆盖 | copy_ingredients 后不再写 Row 1-2，保留参考文件原样 |
| 合并单元格延伸 | 产品 > 5 时，Row 1-11 合并区域从 K 列延伸到 (2+n) 列 |
| 公式保留 | `copy_ingredients` 必须用 `data_only=False` 加载 |
| 性状推断与COA统一 | MSDS 与 COA 共用完全相同的推断逻辑 + 生产方匹配逻辑（`match_mfr` 从 coa_fill 导入） |
| 公式范围按数据末行 | SUM 范围按列内实际有值的最后一行，不用 max_row |
| 匹配待确认 | 评分 < 3 的产品归入"待确认"组，需人工核实生产方对照表 |

### 操作流程

1. 确认成分文件夹路径和生产方对照表
2. 问用户 Row 11 语言设置
3. 运行 `python scripts/msds_fill.py <数据目录> <输出目录> [语言]`
4. 先跑一个 demo group，对照 `resources/MSDS基准/` 验证
5. 用户确认后批量处理全部组

---

## 步骤 ③：Sheet1 产品名超链接

**脚本：** `scripts/hyperlink.py`

### 场景

MSDS 文件已包含多个成分表 sheet 后，为 Sheet1 Row 13 的每个产品名添加指向对应成分表 sheet 的内部超链接。

### 两种映射方式

- **方式 A：Row 12 标签匹配** — Row 12 标签值与 sheet 名一一对应
- **方式 B：位置映射** — Row 12 用 `第N页`，sheet 名用 `第N款成分`，按列序映射

### 用法

```bash
python scripts/hyperlink.py "<folder>"           # 批量处理
python scripts/hyperlink.py "<folder>" --dry-run # 预览不保存
```

### 操作流程

1. 先说明：要处理多少文件、映射逻辑
2. 等用户同意后，先做一个文件验证
3. 用户确认没问题后，批量处理其余
4. 验证：Ctrl+Click Sheet1 产品名应跳转到对应成分表 sheet

---

## 步骤 ④：成分表含量合计检查

**脚本：** `scripts/check_percent.py`

### 场景

批量检查 MSDS 文件每个成分表 sheet 的 D 列（Wt% 标准百分比）和 F 列（Actual Wt% 实际百分比）合计是否 ≈ 100%。

### 判定规则（按优先级）

1. D3/F3 有 `=100-SUM(...)` 公式 → OK（水补齐型，自动保证 100%）
2. D 列全空、F 硬值合计 ≈100 → OK（源文件无 D 列）
3. D 硬值合计 ≈100、F 列含公式 → OK（信任 Excel 计算）
4. D 硬值合计 ≠100 → 异常（可能是混合物分量重复计）
5. F 全硬值且合计 ≠100 → 异常

### 用法

```bash
python scripts/check_percent.py "<base_dir>"           # 递归搜索 _new.xlsx
python scripts/check_percent.py "<base_dir>" --verbose # 显示每个 sheet 状态
```

### 结果解读

| 异常类型 | 含义 | 处理建议 |
|----------|------|----------|
| D/F 均为 =100-SUM 公式 | Row 3 水自动补齐 | 无需处理 |
| D=empty, F=sum≈100 | 源文件无 D 列 | 无需处理 |
| D=sum≈100, F=all-formula | F 全公式引用 | 无需处理 |
| D=sum≠100, F=sum≠100 | 混合物分量重复计数 | 打开 Excel 确认 F 列实际求和 |

---

## 步骤 ⑤：产品名拆分

**脚本：** `scripts/split_spec.py`

### 用法

```bash
python scripts/split_spec.py "<file.xlsx>" "<column>"             # 执行
python scripts/split_spec.py "<file.xlsx>" "<column>" --dry-run   # 预览
```

---

## 步骤 ⑥：CAS号填写

**脚本：** `scripts/fill_cas.py`

### 场景

MSDS 成分表生成后，部分成分的 CAS 列为空白，需要按规则自动填充。CAS 列位于成分表 sheet 的第 8 列（H 列），列头含 "CAS No"。

### 填充规则

| 优先级 | 规则 | 处理方式 |
|--------|------|----------|
| 1 | 已有值（CAS 号 或 `XXXXXX-XXXXX-XXXX` 报送码） | **不动** |
| 2 | `/` 或 `-`（不适用标记） | **不动** |
| 3 | 水 / AQUA / WATER | **跳过** |
| 4 | 香精 / FRAGRANCE / AROMA / PARFUM | 填 `Mixture` |
| 5 | 合并单元格且 CAS 空白 | **拆开合并**，每个成分行各自填对应 CAS |
| 6 | 非合并单元格且 CAS 空白 | 直接填对应 CAS |
| 7 | 无 CAS 的聚合物/混合物（如 ACRYLATES 共聚物） | 跳过 |

### CAS 数据来源

- 脚本内置 `CAS_MAP` 字典，覆盖 140+ 化妆品常见成分的 INCI 名称 → CAS 编号
- 未匹配的空白 CAS 会输出 `not_found` 报告，需人工补充到映射表

### 用法

```bash
python scripts/fill_cas.py "<file_or_dir>"              # 处理单个文件或递归目录
python scripts/fill_cas.py "<file_or_dir>" --dry-run    # 预览不保存
```

### 关键坑

| 坑 | 说明 |
|----|------|
| 报送码不是 CAS | `004183-01464-9628` 格式是化妆品原料报送码，不是 CAS，必须跳过 |
| 合并单元格拆分 | 多成分共用一个 CAS 合并单元格且为空时，必须先 unmerge 再逐行填写 |
| 多行 INCI 名 | 合并单元格导致 INCI 名称含换行符，需 normalize 后再匹配 |
| 水不填 | 水虽有 CAS `7732-18-5`，但 MSDS 惯例不填 |

---

## 步骤 ⑦：保质期填写

**脚本：** `scripts/fill_shelf_life.py`

默认填 "3年"，支持 .xls（COM）和 .xlsx（openpyxl）两种格式。

### 用法

```bash
python scripts/fill_shelf_life.py "<file>" "<column>"                     # 默认填 3年
python scripts/fill_shelf_life.py "<file>" "<column>" --value "2年"       # 自定义值
python scripts/fill_shelf_life.py "<file>" "<column>" --dry-run           # 预览
```

---

## 步骤 ⑧：生成后校对（对照基准验证）

COA 和 MSDS 生成完成后，必须对照 `resources/COA基准/` 或 `resources/MSDS基准/` 中的参考文件进行校对。

### 校对内容

| 校对项 | 检查方法 | 标准 |
|--------|----------|------|
| 表格样式 | 对比生成文件和基准文件的边框、填充色、对齐方式 | 样式与基准一致，填充未破坏原有格式 |
| 字体字号 | 检查数据区域字体名称和大小 | 与模板一致（通常 9pt/11pt），未因写入变形 |
| 数据位置 | 确认委托单位、产品名、规格等是否在正确行列 | 与基准对应位置一致 |
| 合并单元格 | 检查表头合并区域是否正确延伸 | 与产品数量匹配，无断裂 |
| 数据正确性 | 抽查品名、生产单位、成分表内容 | 来源数据准确写入，无错行/偏移 |
| 列宽行高 | 对比成分表 sheet 的列宽行高 | 与源 .xlsm 一致 |

### 校对流程

1. 生成输出文件后，打开对应的基准文件
2. 逐 sheet 对比结构（行列数、合并区域）
3. 随机抽查 3-5 个产品的数据
4. 发现差异时，定位脚本逻辑问题并修正
5. 修正后重新生成，再次校对
