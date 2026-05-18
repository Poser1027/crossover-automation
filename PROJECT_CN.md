# Hengli Orbital Motor Crossover Tool

## ⚡ 协作指令（必读）

**所有回答必须以节省 token 为优先：**
- 回答尽量精简，去掉客套和不必要的解释
- 代码修改只贴**改动的代码块**，不重复贴整个文件
- 让用户检查代码时，**贴完整的代码段供比对**，避免遗漏部分修改
- 不主动展示推理过程，直接给结论和动作
- 修改多处时，所有改动一次性给出，不分多轮

---

## 项目背景

将竞品品牌的液压马达型号解析并匹配到恒立（Hengli）对应产品。用户输入竞品型号代码，系统自动解码规格参数，再在恒立产品数据库中找出最佳匹配型号。

---

## 技术栈

- 语言：Python 3.9+
- 依赖：openpyxl
- 运行方式：命令行 CLI 或作为模块调用

---

## 文件结构

| 文件 | 说明 |
|------|------|
| `crossover_competitor.py` | 主入口：接收竞品型号 → 输出恒立匹配结果 |
| `decode_competitor.py` | 竞品型号解码（解析品牌、系列、参数） |
| `spec_matching.py` | 恒立产品数据库加载与评分匹配逻辑 |
| `Hengli_Orbital_Motor_Master.xlsx` | 恒立产品数据库 |
| `competitor_code_extractor.xlsx` | 竞品型号解码规则表 |

---

## 系列映射

```python
SERIES_MAPPING = {
    ("Char-Lynn (Eaton)", "2000 Series") → Hengli HSP
    ("Char-Lynn (Eaton)", "T Series")    → Hengli HRD
}
```

如需新增品牌/系列，在 `SERIES_MAPPING` 字典中添加对应条目。

---

## 关键参数 & 阈值

| 参数 | 当前值 | 说明 |
|------|--------|------|
| `MIN_SCORE` | 50.0 | 低于此分数的候选直接丢弃 |
| 排量锁定容差 | ±10% | 三个候选必须在此区间内 |
| shaft 直径精确匹配 | ±0.1 mm | 优先精确匹配 |
| shaft 直径松匹配 | ±1.0 mm | 无精确匹配时才回退 |
| `WEAK_PRIMARY_THRESHOLD` | 50.0 | 主匹配分数低于此值时触发兜底 |
| `FALLBACK_TOP_N` | 3 | 兜底候选数量上限 |
| `top_n`（默认） | 3 | 主匹配候选数量上限 |

---

## 匹配逻辑（已更新）

### 排量评分权重（`_score_model`）

| 偏差 | 得分 |
|------|------|
| ≤ 5% | 50（满分） |
| 5–15% | 50 → 20 线性扣 |
| 15–30% | 20 → 2 线性扣 |
| > 30% | 接近 0 |

### 三候选筛选（`find_matches`）

1. 所有 score < 50 的候选直接丢弃
2. 以最高分模型的排量为**锚点**
3. 只在锚点 ±10% 内取候选
4. 同一排量值只保留最高分那个 → 三个候选**排量各不相同**
5. 不够 3 个就只返回实际数量，不补凑

### shaft 直径选择（`_pick_shaft`）

- 先找 ±0.1 mm 精确匹配（处理浮点误差）
- 没有才退到 ±1.0 mm 松匹配
- 都没有才回退到 series 默认值

### shaft 描述正则

兼容两种格式：
- 带 Ø 前缀：`Ø32 straight...`
- 不带前缀：`31.75 (1.250) dia straight shaft...`

正则：`(?:ø\s*)?([0-9]+(?:\.[0-9]+)?)\s*(?:mm|\(|straight|spline|taper)`

### shaft 直径解码（`decode_competitor.py`）

竞品描述中 mm 数值可能在前（`31.75 (1.250)`）或在括号内（`1 inch (25.4 mm)`）：
- 优先匹配 `(数字 mm)` 格式
- 没有才退到"括号前的数字"

---

## 默认填充逻辑

解码结果中以下字段为 `None` 时，系统自动补默认值：

| 字段 | 默认值 |
|------|--------|
| `special` | `"A"` |
| `paint` | `"N"` |
| `rotation` | `"CW"` |

---

## 使用方式

```bash
# CLI
python crossover_competitor.py "M02 02 049 AC 02 AA 01 0 00 1 0 00 00 00 AA AA F"
python crossover_competitor.py M02049AC02AA0100010000000AAAAF
python crossover_competitor.py   # 交互式输入

# 编程调用
from crossover_competitor import crossover_competitor, print_crossover
result = crossover_competitor("M02049AC02AA0100010000000AAAAF")
print_crossover(result)
```

---

## 当前进度

- [x] 竞品型号字符串解码
- [x] 系列映射（Char-Lynn 2000 → HSP，T → HRD）
- [x] 主匹配 + 兜底匹配
- [x] shaft 直径精确匹配（修复 1 inch → 25mm 错配）
- [x] shaft 直径解码（修复 mm/inch 顺序歧义）
- [x] shaft 描述正则兼容 Ø 前缀与无前缀两种格式
- [x] 排量扣分权重加重
- [x] 三候选排量多样性（不再出现仅排量不同的重复型号）
- [x] 分数 < 50 直接丢弃
- [ ] （待补充）

---

## 当前问题 & 反馈

（暂无）

---

## 决策记录

- 使用 `SERIES_MAPPING` 字典做系列映射，而非硬编码 if/else，便于扩展
- 兜底逻辑设计为"先主搜，弱则扩展"，避免直接全库搜索带来的噪音
- 默认值（special/paint/rotation）在 crossover 层填充，不污染 decode 结果
- shaft 直径用"精确优先 + 松匹配兜底"两段制，避免 25.4mm 误匹配到 25mm
- 三候选必须排量各异，仅其他规格不同（mount/shaft/port 等），避免给出冗余建议
- 分数阈值 50 是 fit_rating "Functional" 的下限，低于此建议人工介入而非自动推荐

---

## 修改代码时的常见坑

1. **改完代码后清缓存**：删除 `__pycache__/` 目录，否则 Python 仍加载旧 `.pyc`
2. **多处改动要逐一确认**：让用户检查时贴完整代码块，避免只改了一处就以为完成
3. **正则不能假设格式**：恒立和竞品的描述格式不一致（有/无 Ø 前缀、mm/inch 顺序），需用宽松正则 + 后续过滤
