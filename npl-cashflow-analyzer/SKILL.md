---
name: npl-cashflow-anomaly-detector
description: Detects anomalies in NPL ABS cashflow data by comparing actual vs projected recovery amounts. Uses statistical methods (deviation%, Z-score, rolling trend) and LLM-based judgment for borderline cases. Triggered when user asks about cashflow anomalies, data quality issues, or during scheduled pipeline checks.
metadata:
  type: skill
  category: data-quality
  tags: [abs, npl, cashflow, anomaly-detection, data-quality]
  version: 1.0.0
  author: AI金盾 Team
---

# 现金流异常检测 Skill

## 功能
对比实际回款与评级机构预测，识别异常偏差，判断是数据错误还是真实风险信号。

## 触发条件
- 用户说："检查现金流异常" "哪些项目回款偏离预期" "数据质量检查"
- 新数据入库后自动触发
- 每日调度检查时触发

## 执行步骤

### Step 1: 多维度异常检测
```sql
-- 维度1: 单期偏差率 > ±30%
SELECT p.code, c.period_label, c.actual_amount, c.projected_amount, c.deviation_pct
FROM cashflows c JOIN products p ON c.product_id=p.id
WHERE p.is_invested=1 AND c.is_paid=1 
  AND ABS(c.deviation_pct) > 30
ORDER BY ABS(c.deviation_pct) DESC LIMIT 20

-- 维度2: 连续3期同一方向偏离
-- (需要Python窗口函数或逐产品分析)

-- 维度3: 累计回收率突然跳变(本期与上期差异>10%)
SELECT p.code, c.period_label, c.recovery_rate,
       LAG(c.recovery_rate) OVER (PARTITION BY c.product_id ORDER BY c.period_label) as prev_rate
FROM cashflows c JOIN products p ON c.product_id=p.id
WHERE p.is_invested=1 AND c.is_paid=1
```

### Step 2: 异常分类（规则引擎）
```
偏差率 30-50%       → level: watch   (关注)
偏差率 50-80%       → level: warn    (警告)  
偏差率 >80%         → level: critical (严重)
连续3期同向偏离      → 升级一级
单期极端偏离+后续恢复 → level: info   (可能是数据修正,不告警)
```

### Step 3: LLM复核（仅对warn和critical级别）
```
System: 你是银行风控数据审核专家。
分析以下异常回款记录，判断是"数据录入错误"还是"真实风险信号"。

异常记录:
{product_code} {period}: 实际{actual}万 vs 预计{projected}万, 偏差{deviation}%
历史均值: {avg}万, 标准差: {std}万

判断标准:
- 偏离超过历史均值3σ → 可能是真实异常
- 恰好是整数/整数倍 → 可能是录入错误
- 本期异常但下期恢复正常 → 可能是时间错配(延期回款)
- 同银行多个产品同时异常 → 可能是系统性问题

返回JSON: {"judgment":"data_error|real_risk|time_shift|unknown","confidence":0.85,"reason":"简述"}
```

### Step 4: 生成报告
```
现金流异常检测报告 (2026-05-15)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
严重异常 (需立即处理):
  ABS-0035 Q3-2025: 实际0.7万 vs 预计0.5万 (+46%) → LLM判断: 数据录入错误

警告:
  ABS-0048 H1-2025: 连续3期低于预计 → LLM判断: 真实恶化, 建议启动处置

统计: 严重1个, 警告3个, 关注12个, 信息5个
```

## 输出文件
- `reports/cashflow_anomaly_{date}.md`
- 异常项目的 `alert_level` 更新到 `project_stages` 表

## 集成方式
- 作为Claude Code Skill使用: `/npl-cashflow-anomaly-detector`
- 自动触发: 新数据入库后调用本Skill的SQL检查
