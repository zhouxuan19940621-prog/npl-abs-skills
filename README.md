# NPL ABS 风控智能体 Skills

不良资产证券化投后风控管理的Claude Code Skills集合。

## Skills列表

| Skill | 命令 | 功能 |
|-------|------|------|
| 评级迁移监控 | `/npl-rating-migration-monitor` | 跟踪风险评级变化，检测连续恶化 |
| 现金流异常检测 | `/npl-cashflow-analyzer` | 实际vs预测回款偏差检测 |
| 周度简报生成 | `/npl-auto-forecast` | 自动生成投后管理周报 |

## 快速开始

### 1. 初始化数据库
```bash
python seed_data.py
```
生成 `abs_demo.db`，包含8家银行、4种资产类型、100笔已投产品。

### 2. 复制Skills
```bash
cp -r npl-* ~/.claude/skills/
```

### 3. 使用
在Claude Code中说：
- "检查评级变动"
- "现金流异常检测"
- "生成周报"

## 数据库Schema

```
banks / asset_types / products / product_tranches
cashflows / forecasts / project_stages / pipeline_log
```
详见 `schema.sql`

## 注意事项

- Skills中的SQL查询基于以上Schema
- 需要AI模型支持（Claude/GPT/DeepSeek均可）
- 示例数据仅供测试，生产环境请替换为真实数据
