# Sheet2Anki · 飞书表格 / Excel → Anki apkg（含图片）

把一张**带图片**的表格一键转换成 Anki `.apkg` 牌组文件：支持飞书在线表格（Lark）和本地 Excel（`.xlsx`）两种数据源，文字、嵌套牌组、标签、单元格内嵌图片全部自动处理。

## 这解决什么问题

### 典型场景：扫街拍店 → 登记品牌 → 间隔重复记忆

1. **扫街采集**：逛街时用手机拍摄门店 / Logo / 陈列，直接在**飞书 App 里拍照插入表格单元格**，顺手记下品类、品牌名、价格带。
2. **在线登记**：飞书表格作为多人协作的品牌数据库，补全品牌介绍、标签（如 `潮牌` `高端` `韩系`）。
3. **一键出卡**：运行本工具，每行生成一张卡片——正面是拍的照片，背面是品牌资料，按品类自动归入嵌套牌组。
4. **随时复习**：把 `.apkg` 导入 **Brandki** 或任意 Anki 生态软件（Anki Desktop / AnkiDroid / AnkiMobile），利用间隔重复（Spaced Repetition）在通勤、排队时刷卡片，把扫街见过的品牌真正记住。

### 为什么 CSV 行不通？

CSV 是**纯文本**格式，只能保存单元格里的文字。照片在飞书 / Excel 里是「嵌入在单元格上的图片对象」，一旦导出 CSV，**图片会全部丢失**，只剩空列。

本工具绕开 CSV：

- **飞书模式**：通过开放接口读取单元格里的图片 token 并逐张下载，不需要把整个表格导出为 Excel，也不需要额外的导出授权；
- **Excel 模式**：直接解析 `.xlsx` 压缩包内部的 `xl/drawings` + `xl/media`，按图片锚点定位到所在单元格，完全离线运行。

## 示例表格

第 1 行为表头，一行 = 一张卡片，`Front` 列放**嵌入单元格的图片**：

| Deck1 | Deck2 | Deck3 | Note Type | Front | Back | Fields1_英文名 | Fields2_中文名 | Fields3_品牌级次 | Tags1 | Tags2 |
|---|---|---|---|---|---|---|---|---|---|---|
| 零售 | 女装 | 中淑装 | Basic | 📷 *（嵌入图片：读丽 logo）* | 入卡旗下实体女装品牌，主打松弛轻熟风，面向 25–40 岁女性。 | DULI | 读丽 | C | 韩系 | 低价 |
| 零售 | 运动户外 | 国际零售 | Basic | 📷 *（嵌入图片：鬼冢虎 logo）* | 鬼冢虎（Onitsuka Tiger）1949 年创立于日本神户，经典与街头融合的鞋履品牌。 | Onitsuka Tiger | 鬼冢虎 | D | 潮牌 | 高端 |

生成结果：

- 牌组自动拼成嵌套路径 `零售::女装::中淑装`、`零售::运动户外::国际零售`
- 卡片正面显示图片，背面显示介绍，下方附加英文名 / 中文名 / 品牌级次
- 标签 `韩系 低价`、`潮牌 高端` 随卡带入

可以运行下面的命令生成一份同款的空白模板（含占位图片）：

```bash
python3 examples/make_template.py
# 生成 examples/example_brands.xlsx，替换成真实照片即可使用
```

## 表头约定

表头大小写不敏感，中英文均可，未识别的列会被忽略（可放心保留备注列）：

| 表头写法 | 含义 |
|---|---|
| `Deck1`、`Deck2`、`Deck3`… / `牌组*` | 牌组层级，多列从左到右用 `::` 拼接成嵌套牌组，留空的层级自动跳过 |
| `Front` / `正面` | 卡片正面，支持文字、图片、图文混排 |
| `Back` / `背面` | 卡片背面 / 答案 |
| `Note Type` / `卡片类型` | `Basic`（默认）或 `Cloze` / `填空`（填空题在 Front 里写 `{{c1::挖空内容}}`） |
| `Tags1`、`Tags2`… / `标签*` | 标签，按空格、逗号（`,` `，`）、分号（`;` `；`）拆分，自动去掉开头的 `#` |
| `Fields1_英文名`、`字段2_xxx`… | 额外字段，下划线（或冒号/空格）后的部分作为 Anki 字段名 |

## 安装

需要 Python 3.9+：

```bash
pip install -r requirements.txt
# 依赖：genanki（生成 apkg）、openpyxl（读取 xlsx）
```

飞书模式另外需要 `lark-cli`（Trae 的 Lark 插件自带，会被自动发现；也可用环境变量 `LARK_CLI` 指定路径）并完成飞书登录；纯 Excel 模式无任何外部依赖。

## 使用

### 方式一：从飞书在线表格生成

```bash
python3 scripts/feishu_to_anki.py \
  --url "https://example.feishu.cn/sheets/XXXXXXXX" \
  -o brands.apkg
```

图片通过单元格图片 token 逐张下载并打包，**无需导出整个表格**。

### 方式二：从本地 Excel 生成（离线）

```bash
python3 scripts/feishu_to_anki.py \
  --xlsx ./brands.xlsx \
  --sheet-name "Anki卡片" \
  -o brands.apkg
```

`.xlsx` 里的嵌入图片（含 WPS / Excel / 飞书导出的文件）会自动提取。

### 全部参数

| 参数 | 说明 |
|---|---|
| `--url` / `--xlsx` | 二选一：飞书表格 URL，或本地 xlsx 路径 |
| `--sheet-name` | 子表名称，默认取第一张表 |
| `-o, --output` | 输出 apkg 路径，默认 `./<表名>.apkg` |
| `--deck` | Deck 列留空时使用的兜底牌组名 |
| `--keep-media` | 保留中间下载 / 提取的图片（默认用完即清） |

运行结束会打印 JSON 摘要：卡片数、牌组路径、图片数、跳过的空行数。

## 导入复习

- **Brandki**：在 App 内选择导入 `.apkg`，图片与牌组结构一并带入。
- **Anki 桌面版**：`文件 → 导入`；**AnkiDroid / AnkiMobile**：把文件发送到设备后选择用 Anki 打开。
- **增量更新**：卡片 GUID 按「表格 + 行号」稳定生成，表格更新后重新生成并再次导入，会**更新原卡片而不是重复创建**，复习记录也能保留。注意不要随意调整行顺序。

## 作为 Trae Skill 使用

本仓库附带 `SKILL.md`，在 Trae 中可直接作为 Skill 安装：

```bash
mkdir -p .trae/skills/feishu-sheet-to-anki
cp SKILL.md .trae/skills/feishu-sheet-to-anki/
cp -r scripts .trae/skills/feishu-sheet-to-anki/
```

之后直接说「把这个飞书表格生成 Anki」即可自动执行。

## 限制

- 仅支持单行表头的扁平表格；
- 只识别**嵌入单元格的图片**，悬浮在表格画布上的浮动图片不会被取到；
- 不包含学习进度（新卡状态），分享牌组这通常正是期望行为，迁移进度请用 AnkiWeb 同步；
- 超大图片会原样打包，建议在表格里先压缩以减小 apkg 体积。

## License

[MIT](./LICENSE)
