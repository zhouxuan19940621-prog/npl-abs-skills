---
name: npl-weekly-report-generator
description: Generates a comprehensive weekly ABS NPL portfolio report including recovery summary, rating migration, cashflow anomalies, and next-week focus items. Uses LLM to write executive summary and action recommendations. Triggered when user asks for weekly report or on scheduled Monday runs.
metadata:
  type: skill
  category: reporting
  tags: [abs, npl, report, weekly, llm]
  version: 1.0.0
  author: AI金盾 Team
---

# 周度简报生成 Skill

## 功能
每周一自动生成ABS投后管理周报，包含数据统计、异常标记、趋势分析和操作建议。

## 触发条件
- 用户说："生成周报" "本周回收情况" "weekly report"
- 每周一上午9:00自动触发

## 报告结构

### 一、本周回收概览
```sql
-- 本周回款总额
SELECT SUM(c.actual_amount) as total 
FROM cashflows c JOIN products p ON c.product_id=p.id
WHERE p.is_invested=1 AND c.is_paid=1
  AND c.start_date BETWEEN '上周一' AND '本周一'

-- 本周回款vs预计偏差
SELECT SUM(c.actual_amount) as actual, SUM(c.projected_amount) as projected
FROM cashflows c JOIN products p ON c.product_id=p.id  
WHERE p.is_invested=1
  AND c.start_date BETWEEN '上周一' AND '本周一'

-- 按资产类型分组
SELECT at.name, SUM(c.actual_amount) as total
FROM cashflows c JOIN products p ON c.product_id=p.id
JOIN asset_types at ON p.asset_type_id=at.id
WHERE p.is_invested=1 AND c.is_paid=1
  AND c.start_date BETWEEN '上周一' AND '本周一'
GROUP BY at.name
```

### 二、评级变动汇总
调用评级迁移监控Skill的结果，汇总本周评级变化。

### 三、异常项目列表
调用现金流异常检测Skill的结果，列出本周新增异常。

### 四、下周关注
- 下周预计兑付的项目列表
- 到期需检查的项目
- 风险评分最高的5个项目

### 五、LLM生成执行摘要
```
System: 你是银行资产管理部门负责人。基于以下本周ABS投后数据，撰写一段200字的执行摘要：
- 本周回收总额: {amount}
- 评级恶化: {downgrade_count}个项目
- 异常回款: {anomaly_count}条
- 下周到期兑付: {due_count}个项目

要求: 数据准确、风险突出、建议可行。
```

## 输出文件
- `reports/weekly_{date}.md` — Markdown版本
- `reports/weekly_{date}.pdf` — PDF版本(可选)
- 如果配置了邮件，自动发送

## 集成方式
- 手动触发: `/npl-weekly-report-generator`
- 自动触发: 调度中心每周一9:00调用
