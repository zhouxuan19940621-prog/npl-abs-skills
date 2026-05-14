---
name: npl-rating-migration-monitor
description: Monitors NPL ABS product rating migrations over time. Detects consecutive downgrades, generates early-warning alerts, and recommends intervention actions. Triggered when user asks about rating changes, risk deterioration, or during daily pipeline checks.
metadata:
  type: skill
  category: risk-management
  tags: [abs, npl, rating, monitoring, early-warning]
  version: 1.0.0
  author: AI金盾 Team
---

# 评级迁移监控 Skill

## 功能
跟踪每笔ABS产品的风险评级变化，识别连续恶化趋势，自动生成预警。

## 触发条件
- 用户说："检查评级变动" "有哪些项目评级恶化了" "评级迁移情况"
- 每日调度检查自动触发

## 执行步骤

### Step 1: 查询数据
```sql
SELECT p.code, p.name, p.current_rating, p.risk_trend, 
       f.current_rating as prev_rating, f.forecast_date
FROM products p 
JOIN project_stages s ON p.id=s.product_id
LEFT JOIN forecasts f ON p.id=f.product_id 
  AND f.forecast_date = (SELECT MAX(forecast_date) FROM forecasts WHERE product_id=p.id)
WHERE p.is_invested=1 
  AND p.current_rating IN ('关注','次级','可疑')
ORDER BY p.current_rating, p.risk_trend
```

### Step 2: 分析迁移模式
- 正常→关注：标记为watch，检查回收率是否也下降
- 关注→次级：标记为warn，触发处置建议
- 次级→可疑：标记为critical，需人工复核
- 连续2期同一方向恶化：升级告警级别

### Step 3: 生成报告
输出格式：
```
评级迁移报告 (2026-05-15)
━━━━━━━━━━━━━━━━━━━━━━
恶化项目 (需关注):
  ICBC-2023-001: 正常→关注, 回收率下降12%, 建议: 增加催收频率
  CCB-2023-015: 关注→次级, 连续2期恶化, 建议: 启动处置预案

改善项目:
  BOC-2023-022: 关注→正常, 回收率回升至35.3%

统计:
  恶化: 3个  改善: 1个  稳定: 96个
```

### Step 4: 如果调用大模型
将迁移数据发给LLM，附加系统提示词：
```
你是银行风控专家。基于以下评级迁移数据，撰写300字分析报告：
- 重点分析连续恶化项目的原因
- 提出具体处置建议
- 评估整体组合风险变化趋势
```

## 输出文件
- 保存到 `reports/rating_migration_{date}.md`
- 如果告警≥2条，发送通知

## 集成方式
- 作为Claude Code Skill使用: `/npl-rating-migration-monitor`
- 或集成到调度中心: 每日检查时自动调用本Skill的SQL和逻辑
