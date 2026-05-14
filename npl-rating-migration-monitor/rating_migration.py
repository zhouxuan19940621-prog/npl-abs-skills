"""
评级迁移监控 — 可执行脚本
用法: python rating_migration.py [--db abs_demo.db] [--alert]
"""

import sqlite3, argparse, json
from datetime import datetime

def check_migration(db_path: str, alert: bool = False):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # 查询当前评级和上次预测评级
    rows = c.execute("""
        SELECT p.code, p.name, b.short_name as bank, at.name as asset_type,
               p.current_rating, p.risk_trend,
               f.current_rating as prev_rating, f.forecast_date,
               COALESCE((SELECT recovery_rate FROM cashflows WHERE product_id=p.id AND is_paid=1
                         ORDER BY end_date DESC LIMIT 1), 0) as latest_recovery
        FROM products p
        JOIN banks b ON p.bank_id=b.id
        JOIN asset_types at ON p.asset_type_id=at.id
        LEFT JOIN forecasts f ON p.id=f.product_id
          AND f.forecast_date = (SELECT MAX(forecast_date) FROM forecasts WHERE product_id=p.id)
        WHERE p.is_invested=1
    """).fetchall()

    # 分析迁移
    downgrades = []
    upgrades = []
    stable = []

    for r in rows:
        prev = r["prev_rating"] or r["current_rating"]
        curr = r["current_rating"]
        levels = {"正常": 0, "关注": 1, "次级": 2, "可疑": 3}

        if levels.get(curr, 0) > levels.get(prev, 0):
            downgrades.append(dict(r))
        elif levels.get(curr, 0) < levels.get(prev, 0):
            upgrades.append(dict(r))
        else:
            stable.append(dict(r))

    print(f"\n{'='*60}")
    print(f"  评级迁移报告 - {datetime.now().strftime('%Y-%m-%d')}")
    print(f"{'='*60}")
    print(f"\n  恶化: {len(downgrades)}个  改善: {len(upgrades)}个  稳定: {len(stable)}个")

    if downgrades:
        print(f"\n  恶化项目:")
        for d in downgrades:
            print(f"    {d['code']} {d['name'][:20]}: {d['prev_rating']}→{d['current_rating']} | 回收率{d['latest_recovery']:.1f}% | {d['bank']}")

    if upgrades:
        print(f"\n  改善项目:")
        for u in upgrades:
            print(f"    {u['code']} {u['name'][:20]}: {u['prev_rating']}→{u['current_rating']} | 回收率{u['latest_recovery']:.1f}%")

    # 生成JSON输出
    result = {
        "date": datetime.now().strftime('%Y-%m-%d'),
        "summary": {"downgrades": len(downgrades), "upgrades": len(upgrades), "stable": len(stable)},
        "downgrades": downgrades,
        "upgrades": upgrades
    }

    path = f"rating_migration_{datetime.now().strftime('%Y%m%d')}.json"
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n  报告已保存: {path}")

    conn.close()
    return result

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="评级迁移监控")
    parser.add_argument("--db", default="abs_demo.db", help="数据库路径")
    parser.add_argument("--alert", action="store_true", help="仅输出告警项目")
    args = parser.parse_args()
    check_migration(args.db, args.alert)
