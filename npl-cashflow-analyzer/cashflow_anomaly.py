"""
现金流异常检测 — 可执行脚本
用法: python cashflow_anomaly.py [--db abs_demo.db] [--threshold 30]
"""

import sqlite3, argparse, json
from datetime import datetime

def detect_anomalies(db_path: str, threshold: float = 30.0):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    rows = c.execute("""
        SELECT p.code, p.name, b.short_name as bank, at.name as asset_type,
               c.period_label, c.actual_amount, c.projected_amount, c.deviation_pct,
               c.recovery_rate,
               AVG(c.deviation_pct) OVER (PARTITION BY c.product_id) as avg_deviation
        FROM cashflows c
        JOIN products p ON c.product_id=p.id
        JOIN banks b ON p.bank_id=b.id
        JOIN asset_types at ON p.asset_type_id=at.id
        WHERE p.is_invested=1 AND c.is_paid=1
          AND ABS(c.deviation_pct) > ?
        ORDER BY ABS(c.deviation_pct) DESC
    """, (threshold,)).fetchall()

    # 分类
    critical = [r for r in rows if abs(r["deviation_pct"]) > 80]
    warnings = [r for r in rows if 50 < abs(r["deviation_pct"]) <= 80]
    watches  = [r for r in rows if threshold < abs(r["deviation_pct"]) <= 50]

    print(f"\n{'='*60}")
    print(f"  现金流异常检测 - {datetime.now().strftime('%Y-%m-%d')}")
    print(f"{'='*60}")
    print(f"\n  严重({len(critical)}个)  警告({len(warnings)}个)  关注({len(watches)}个)")

    if critical:
        print(f"\n  【严重异常】")
        for r in critical:
            direction = "超额回收" if r["deviation_pct"] > 0 else "低于预期"
            print(f"    {r['code']} {r['period_label']}: 实际{r['actual_amount']:.1f}万 vs 预计{r['projected_amount']:.1f}万 ({r['deviation_pct']:+.1f}%) {direction}")

    if warnings:
        print(f"\n  【警告】")
        for r in warnings[:10]:
            print(f"    {r['code']} {r['period_label']}: 偏差{r['deviation_pct']:+.1f}%")

    result = {
        "date": datetime.now().strftime('%Y-%m-%d'),
        "threshold": threshold,
        "summary": {"critical": len(critical), "warning": len(warnings), "watch": len(watches)},
        "critical": [dict(r) for r in critical],
        "warnings": [dict(r) for r in warnings]
    }

    path = f"cashflow_anomaly_{datetime.now().strftime('%Y%m%d')}.json"
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n  报告已保存: {path}")

    conn.close()
    return result

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="现金流异常检测")
    parser.add_argument("--db", default="abs_demo.db", help="数据库路径")
    parser.add_argument("--threshold", type=float, default=30.0, help="偏差阈值(%)")
    args = parser.parse_args()
    detect_anomalies(args.db, args.threshold)
