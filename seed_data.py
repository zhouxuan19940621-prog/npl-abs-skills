"""
生成示例数据 — 供Skills测试使用
运行: python seed_data.py
输出: abs_demo.db (SQLite)
"""

import sqlite3, random, json

DB_PATH = "abs_demo.db"
random.seed(42)

conn = sqlite3.connect(DB_PATH)
conn.execute("PRAGMA journal_mode=WAL")
c = conn.cursor()

# ── 建表 ──────────────────────────────────────────────
with open("schema.sql", "r", encoding="utf-8") as f:
    # 只执行CREATE TABLE部分，跳过注释和扩展
    for stmt in f.read().split(";"):
        stmt = stmt.strip()
        if stmt.upper().startswith("CREATE"):
            try:
                c.execute(stmt)
            except:
                pass

# ── 银行 ──────────────────────────────────────────────
banks = [("ICBC","工商银行"),("CCB","建设银行"),("ABC","农业银行"),
         ("BOC","中国银行"),("BOCOM","交通银行"),("CMB","招商银行"),
         ("SPDB","浦发银行"),("CIB","兴业银行")]
for code, name in banks:
    c.execute("INSERT INTO banks(code,name,short_name) VALUES(?,?,?)", (code,name,name))

# ── 资产类型 ───────────────────────────────────────────
for row in [("mortgage","房屋抵押类","mortgage","semi_annual",2),
            ("credit_card","信用卡类","credit","quarterly",3),
            ("micro_loan","小微贷款类","credit","quarterly",3),
            ("consumer_loan","消费贷类","credit","quarterly",3)]:
    c.execute("INSERT INTO asset_types(code,name,category,normalize_frequency,normalize_multiplier) VALUES(?,?,?,?,?)", row)

# ── 产品 (100笔已投) ──────────────────────────────────
asset_ids = [1,1,1,2,2,3,3,4]
for i in range(100):
    bank_id = random.randint(1,8)
    aid = random.choice(asset_ids)
    amt = round(random.uniform(5,50), 2)
    rating = random.choices(["正常","关注","次级","可疑"], weights=[60,25,10,5])[0]
    c.execute("""INSERT INTO products(code,name,bank_id,asset_type_id,issue_date,total_amount,
        our_investment,is_invested,outstanding_principal,current_rating,risk_trend,data_quality_score,status)
        VALUES(?,?,?,?,?,?,?,1,?,?,?,?,'active')""",
        (f"ABS-{i+1:04d}", f"第{i+1}期ABS产品", bank_id, aid, f"202{random.randint(1,4)}-01-01",
         amt, amt*0.3, round(amt*random.uniform(0.3,0.9),2),
         rating, random.choice(["改善","稳定","恶化","波动"]), round(random.uniform(60,100),1)))

print(f"已生成: {len(banks)}家银行, 4种资产类型, 100笔产品")
conn.commit()
conn.close()
print("数据库: abs_demo.db 创建完成")
