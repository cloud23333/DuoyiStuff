# 预测算法迭代路线图

本文档记录日销预测模型的后续可迭代方向，按"成本低→收益稳→工程量大"分四个阶段。阶段 1 已于 2026-04-20 落地（EWMA 替换固定窗口 + 策略双重确认 + 阈值常量化）。阶段 3 已于 2026-06-01 落地（重构为「预测分布 → 分位数决策」两层：pure-stdlib NegBin/Poisson 逆 CDF + negbin/poisson/hurdle EWMA 估计器，holdout bake-off 选定 hurdle；`service_level` 单参数替代离散策略并吸收热销/全局乘子；`--global-gap-multiplier` 改为 `--service-level-offset`）。详见 `docs/superpowers/specs/2026-05-31-distribution-quantile-forecast-design.md`。

---

## 当前算法概要（阶段 1 落地后）

输入：`(SKC, SKUID)` 每日销量序列、备货天数、是否热销款、当日库存。

核心步骤：

1. **缺货尾部掩码**（`_mask_stockout_tail`）：当前无库存时，剥离末尾连续零销量。
2. **基础统计量**：median、mean、`ewma_short`（half-life=2）、`ewma_long`（half-life=5）、`trimmed_mean`、`volatility`、`isolated_spike`。
3. **慢销分支**：median=0 且 mean>0 时走 Poisson 均值兜底。
4. **策略选择**（双重确认）：
   - 保守：`ewma_short < baseline×0.6` 且 `ewma_long ≤ baseline`
   - 保守：`volatility > 1.2` 或 `isolated_spike`
   - 激进：`ewma_short > baseline×1.5` 且 `ewma_long ≥ baseline`
   - 激进：热销款且 `ewma_long ≥ baseline`
   - 其余：正常
5. **日均销量公式**：
   - 保守：`min(median, ewma_long, trimmed_mean)`
   - 激进：`0.7·ewma_short + 0.3·ewma_long`，上限 `max(median×2, ewma_short×1.2)`
   - 正常：`0.5·median + 0.5·ewma_long`
6. **缺口计算**：`ceil(max(0, 日均×备货天数 − 可用库存))`，热销 ×1.2 再乘全局倍率。

---

## 阶段 2：季节性与冷启动（中成本，3–5 天）

补两块当前完全没有的盲区。

### 2.1 周内效应（day-of-week）
- 每个 SKU 估一组 7 维因子：`dow_factor[i] = mean_i / overall_mean`
- 预测输出按目标日 dow 调整
- 数据不足（<21 天或单 dow 样本<3）时退化到**类目全局因子**

### 2.2 促销/节日日历
- 在 `daily_sales` 输入加一列"是否活动日"标注
- 训练时对活动日**降权**（如权重 0.3），避免把爆单当常态趋势
- 或单独走一条"事件后回落"的衰减曲线

### 2.3 新品冷启动
- `len(values) < 7` 时不走现有 4 策略
- 改用**同 SKC 兄弟 SKU 均值**或**类目先验**作贝叶斯先验
- 随着样本增长平滑过渡到标准策略

**预期收益**：覆盖当前盲区，周末爆单不再被当异常尖峰砍掉，新品首周备货合理化。

---

## 阶段 3：点估计 → 分布决策（高成本，1–2 周）✅ 已落地（2026-06-01）

把"预测日均 × 备货天数"升级成真正的库存决策。落地实现见 `forecast_distribution.py` / `forecast_decision.py` / `forecast_level.py`。以下为原始设计条目（均已实现）。

### 3.1 输出预测分布
- 用 Poisson 或 NegBin 拟合日销序列
- 得到 `mean + dispersion`，不只是点估计
- 备货期销量用分布的**分位数**（P70/P85/P95）替代点估计

### 3.2 服务水平可调
- 用**单一参数** `service_level ∈ [0.5, 0.95]` 替代 4 个离散策略
- "保守/正常/激进/慢销"变成 4 个 service_level 档位
- 策略过渡连续化，避免边界抖动

### 3.3 合并热销/全局乘数
- 把 `HOT_STYLE_GAP_MULTIPLIER=1.2` 并入 `service_level`（热销款 service_level 调高一档）
- 把 `--global-gap-multiplier` 改为全局 `service_level` 偏移
- 避免多层乘数叠加后语义模糊

### 3.4 gap 公式透明化
- 新公式：`gap = quantile(demand, service_level) − available_stock`
- 每行报表输出"选择 P?? 分位"，可解释性↑
- 业务侧按品类设 service_level，开发侧不再改算法

**预期收益**：决策语义从"拍脑袋乘数"变成"服务水平 → 安全库存"，跨部门沟通有统一刻度。

---

## 阶段 4：闭环评估与自适应（长期工程）

让算法能自己变好。

### 4.1 固化回测流程
- 把 `eval_forecast.py` 做成 `make eval` / CI 任务
- 按 SKC 输出 **MAE / bias / 缺货率**三件套
- 每次算法改动 PR 自动跑回测对比
- `eval_policy.py` 实现**策略成本回测**：对 `(service_level_offset, alpha)` 旋钮网格回放历史，输出每格的发货量 / 缺货量 / 滞销量与履约率，供运营在效率前沿上挑选更保守的设置，而不是凭感觉拍。需要注意的限制是：仓里没有历史库存快照，所以回测假设每个 holdout 窗口从**零库存起步、单次补货** `target` 件，绝对成本偏乐观，但不同旋钮设置之间的**相对排序**是可靠的（运营要选的正是相对排序）。

### 4.2 按 SKC 自适应参数
- 同一套算法，不同 SKC 学不同的 α（EWMA 衰减）与 service_level
- 基于回测误差**定期重拟合**（周级或月级）
- 参数存 `data/input/sku_profiles.json`

### 4.3 异常 SKU 告警
- 连续 N 天预测偏差超阈值的 SKU 自动进**数据质量报告**
- 人工复核清单按偏差金额排序

### 4.4 A/B 对照监控
- `scripts/compare_outputs.py` 升级为**固定基线 vs 新策略**的差异监控
- 每次算法变更后自动生成差异分布报告
- 防止算法无声回归

**预期收益**：从"靠工程师拍参数"变成"数据驱动参数"。规模化后必须做。

---

## 延后项（需要先拿到数据）

### 中间段 OOS 掩码
- 原阶段 1 范围，因缺历史库存数据被延后
- 实现需要库存快照序列（当前只有当日库存）
- 拿到数据后可放入阶段 2

### 退货数据
- 当前算法只看销售侧，不考虑退货
- 退货会在未来几天释放可用库存
- 需要订单状态变更流水

### 价格/活动弹性
- 促销降价会放大销量，非促销期会回落
- 当前没有价格信号输入
- 属于阶段 3 之后的长期项

---

## 推荐下一步

按阶段排序，谨慎组合建议：

1. **阶段 2.1（dow 因子）** — 独立、改动小、验证成本低
2. **阶段 2.3（冷启动先验）** — 与 dow 因子解耦，可并行
3. **阶段 3.2+3.3（service_level 统一参数）** — 需要业务侧对齐
4. **阶段 4.1（回测 CI）** — 前三项落地后必须有，否则优化无法验证

阶段 2.2（促销标注）依赖数据源改造，优先级放在 dow/冷启动之后。
