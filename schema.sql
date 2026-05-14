-- ============================================================
-- ABS投后风控智能体 · 数据库Schema v1.0
-- PostgreSQL 15+
-- 6表 + 2视图 + 索引
-- ============================================================

-- ── 扩展 ──────────────────────────────────────────────────
-- CREATE EXTENSION IF NOT EXISTS "uuid-ossp";
-- CREATE EXTENSION IF NOT EXISTS vector;  -- 向量检索(后续)

-- ═══════════════════════════════════════════════════════════
-- TABLE 1: 银行
-- ═══════════════════════════════════════════════════════════
CREATE TABLE banks (
    id              SERIAL PRIMARY KEY,
    code            VARCHAR(20)  NOT NULL UNIQUE,       -- 银行代码 ICBC/CCB/ABC...
    name            VARCHAR(100) NOT NULL,              -- 银行全称
    short_name      VARCHAR(50),                        -- 简称 工商银行
    total_products  INT DEFAULT 0,                      -- 该行发行ABS总数
    invested_count  INT DEFAULT 0,                      -- 我们投资的数量
    created_at      TIMESTAMP DEFAULT NOW(),
    updated_at      TIMESTAMP DEFAULT NOW()
);

-- ═══════════════════════════════════════════════════════════
-- TABLE 2: 底层资产类型
-- ═══════════════════════════════════════════════════════════
CREATE TABLE asset_types (
    id                  SERIAL PRIMARY KEY,
    code                VARCHAR(20)  NOT NULL UNIQUE,   -- mortgage / credit_card / micro_loan
    name                VARCHAR(100) NOT NULL,          -- 房屋抵押类 / 信用卡类
    category            VARCHAR(20)  NOT NULL,          -- mortgage(抵押类) / credit(信用类)
    normalize_frequency VARCHAR(10)  NOT NULL,          -- semi_annual / quarterly
    normalize_multiplier INT DEFAULT 1,                 -- 标准化乘数
    description         TEXT
);

-- 初始数据
INSERT INTO asset_types (code, name, category, normalize_frequency, normalize_multiplier) VALUES
('mortgage',      '房屋抵押类',  'mortgage', 'semi_annual', 2),
('credit_card',   '信用卡类',    'credit',   'quarterly',   3),
('micro_loan',    '小微贷款类',  'credit',   'quarterly',   3),
('consumer_loan', '消费贷类',    'credit',   'quarterly',   3),
('hybrid',        '混合类',      'credit',   'quarterly',   3),
('auto_loan',     '车贷类',      'credit',   'quarterly',   3);

-- ═══════════════════════════════════════════════════════════
-- TABLE 3: 产品 (ABS)
-- ═══════════════════════════════════════════════════════════
CREATE TABLE products (
    id                  SERIAL PRIMARY KEY,
    code                VARCHAR(30)  NOT NULL UNIQUE,   -- 产品代码 ICBC-NPL-2023-001
    name                VARCHAR(200) NOT NULL,          -- 产品全称
    bank_id             INT REFERENCES banks(id),
    asset_type_id       INT REFERENCES asset_types(id),
    issue_date          DATE,                           -- 发行日期
    total_amount        DECIMAL(18,2),                  -- 总规模(元)
    our_investment      DECIMAL(18,2),                  -- 我们投资金额
    is_invested         BOOLEAN DEFAULT FALSE,          -- 是否已投资
    outstanding_principal DECIMAL(18,2),                -- 剩余本金
    current_rating      VARCHAR(10) DEFAULT '未评级',    -- 当前风险评级
    risk_trend          VARCHAR(10),                    -- 风险趋势
    data_quality_score  DECIMAL(5,1) DEFAULT 100,       -- 数据质量分
    status              VARCHAR(20) DEFAULT 'active',   -- active/closed/defaulted
    created_at          TIMESTAMP DEFAULT NOW(),
    updated_at          TIMESTAMP DEFAULT NOW()
);

-- ═══════════════════════════════════════════════════════════
-- TABLE 4: 产品分层结构
-- ═══════════════════════════════════════════════════════════
CREATE TABLE product_tranches (
    id                  SERIAL PRIMARY KEY,
    product_id          INT REFERENCES products(id) ON DELETE CASCADE,
    tranche_level       VARCHAR(20)  NOT NULL,          -- senior(优先) / mezzanine(夹层) / junior(劣后)
    tranche_order       INT NOT NULL,                   -- 兑付顺序 1优先→2夹层→3劣后
    amount              DECIMAL(18,2),                  -- 该层规模
    payment_frequency   VARCHAR(15)  NOT NULL,          -- monthly / quarterly / semi_annual / annual
    start_date          DATE,                           -- 该层开始兑付日期
    end_date            DATE,                           -- 该层结束兑付日期
    is_paid_off         BOOLEAN DEFAULT FALSE,          -- 是否已兑付完毕
    created_at          TIMESTAMP DEFAULT NOW()
);

-- ═══════════════════════════════════════════════════════════
-- TABLE 5: 现金流 (每期回款明细)
-- ═══════════════════════════════════════════════════════════
CREATE TABLE cashflows (
    id                  SERIAL PRIMARY KEY,
    product_id          INT REFERENCES products(id) ON DELETE CASCADE,
    tranche_id          INT REFERENCES product_tranches(id),
    period_label        VARCHAR(20)  NOT NULL,          -- 期次标签 H1-2024 / Q3-2025 / M6-2024
    start_date          DATE NOT NULL,                  -- 起始日期
    end_date            DATE NOT NULL,                  -- 截止日期
    frequency           VARCHAR(15)  NOT NULL,          -- 原始频率 monthly/quarterly/semi_annual

    actual_amount       DECIMAL(18,2),                  -- 实际回款额
    projected_amount    DECIMAL(18,2),                  -- 评级机构预计额
    deviation_pct       DECIMAL(8,4),                   -- 偏差率 (actual-projected)/projected*100

    cumulative_actual   DECIMAL(18,2),                  -- 累计实际回收
    cumulative_projected DECIMAL(18,2),                 -- 累计预计回收
    recovery_rate       DECIMAL(8,4),                   -- 当期回收率 = actual/total_amount*100

    is_paid             BOOLEAN DEFAULT FALSE,          -- 是否已回款
    is_estimated        BOOLEAN DEFAULT FALSE,          -- 是否为预估值(非实际到账)
    remarks             TEXT,

    created_at          TIMESTAMP DEFAULT NOW()
);

-- ═══════════════════════════════════════════════════════════
-- TABLE 6: 预测结果
-- ═══════════════════════════════════════════════════════════
CREATE TABLE forecasts (
    id                  SERIAL PRIMARY KEY,
    product_id          INT REFERENCES products(id) ON DELETE CASCADE,
    forecast_date       DATE NOT NULL,                  -- 预测执行日期
    model_version       VARCHAR(20),                    -- 模型版本
    forecast_periods    INT DEFAULT 5,                  -- 预测期数

    current_rating      VARCHAR(10),                    -- 预测时的评级
    risk_trend          VARCHAR(20),                    -- 趋势
    trend_strength      DECIMAL(5,3),                   -- 趋势强度
    data_quality_score  DECIMAL(5,1),                   -- 数据质量分

    current_cum_rate    DECIMAL(8,4),                   -- 当前累计回收率
    predicted_final_rate DECIMAL(8,4),                  -- 预测最终回收率
    confidence          DECIMAL(5,3),                   -- 模型置信度

    details             JSONB,                          -- 每期预测明细 [{period, forecast, cumulative, lower, upper}]
    stress_results      JSONB,                          -- 压力测试结果
    chart_path          VARCHAR(500),                   -- 图表文件路径

    created_at          TIMESTAMP DEFAULT NOW()
);

-- ═══════════════════════════════════════════════════════════
-- INDEXES
-- ═══════════════════════════════════════════════════════════
CREATE INDEX idx_products_bank      ON products(bank_id);
CREATE INDEX idx_products_asset     ON products(asset_type_id);
CREATE INDEX idx_products_invested  ON products(is_invested);
CREATE INDEX idx_products_rating    ON products(current_rating);

CREATE INDEX idx_tranches_product   ON product_tranches(product_id);
CREATE INDEX idx_tranches_level     ON product_tranches(tranche_level);

CREATE INDEX idx_cashflows_product  ON cashflows(product_id);
CREATE INDEX idx_cashflows_dates    ON cashflows(start_date, end_date);
CREATE INDEX idx_cashflows_freq     ON cashflows(frequency);
CREATE INDEX idx_cashflows_paid     ON cashflows(is_paid);

CREATE INDEX idx_forecasts_product  ON forecasts(product_id);
CREATE INDEX idx_forecasts_date     ON forecasts(forecast_date);

-- ═══════════════════════════════════════════════════════════
-- VIEW 1: 标准化现金流 (按asset_type规则聚合)
-- ═══════════════════════════════════════════════════════════
CREATE OR REPLACE VIEW normalized_cashflows AS
SELECT
    c.id,
    c.product_id,
    p.code                     AS product_code,
    p.name                     AS product_name,
    p.bank_id,
    b.name                     AS bank_name,
    at.code                    AS asset_type_code,
    at.name                    AS asset_type_name,
    at.category                AS asset_category,
    at.normalize_frequency     AS compare_frequency,

    -- 标准化周期标签
    CASE at.normalize_frequency
        WHEN 'semi_annual' THEN CONCAT('H', EXTRACT(YEAR FROM c.start_date), '-',
            CASE WHEN EXTRACT(MONTH FROM c.start_date) <= 6 THEN '1' ELSE '2' END)
        WHEN 'quarterly'   THEN CONCAT('Q', EXTRACT(YEAR FROM c.start_date), '-',
            EXTRACT(QUARTER FROM c.start_date))
        ELSE c.period_label
    END AS normalized_period,

    -- 标准化金额
    CASE at.normalize_frequency
        WHEN 'semi_annual' THEN
            SUM(c.actual_amount) OVER (
                PARTITION BY c.product_id,
                CASE WHEN EXTRACT(MONTH FROM c.start_date) <= 6
                     THEN CONCAT(EXTRACT(YEAR FROM c.start_date), '-H1')
                     ELSE CONCAT(EXTRACT(YEAR FROM c.start_date), '-H2')
                END
            )
        WHEN 'quarterly' THEN
            SUM(c.actual_amount) OVER (
                PARTITION BY c.product_id,
                CONCAT(EXTRACT(YEAR FROM c.start_date), '-Q', EXTRACT(QUARTER FROM c.start_date))
            )
        ELSE c.actual_amount
    END AS normalized_actual,

    CASE at.normalize_frequency
        WHEN 'semi_annual' THEN
            SUM(c.projected_amount) OVER (
                PARTITION BY c.product_id,
                CASE WHEN EXTRACT(MONTH FROM c.start_date) <= 6
                     THEN CONCAT(EXTRACT(YEAR FROM c.start_date), '-H1')
                     ELSE CONCAT(EXTRACT(YEAR FROM c.start_date), '-H2')
                END
            )
        WHEN 'quarterly' THEN
            SUM(c.projected_amount) OVER (
                PARTITION BY c.product_id,
                CONCAT(EXTRACT(YEAR FROM c.start_date), '-Q', EXTRACT(QUARTER FROM c.start_date))
            )
        ELSE c.projected_amount
    END AS normalized_projected,

    c.start_date,
    c.end_date,
    c.frequency                 AS original_frequency,
    c.actual_amount             AS original_actual,
    c.projected_amount          AS original_projected,
    c.is_paid,
    c.is_estimated

FROM cashflows c
JOIN products p       ON c.product_id = p.id
JOIN banks b          ON p.bank_id = b.id
JOIN asset_types at   ON p.asset_type_id = at.id;

-- ═══════════════════════════════════════════════════════════
-- VIEW 2: 产品投资总览
-- ═══════════════════════════════════════════════════════════
CREATE OR REPLACE VIEW portfolio_summary AS
SELECT
    p.id,
    p.code,
    p.name,
    b.name                     AS bank_name,
    at.name                    AS asset_type_name,
    at.category                AS asset_category,
    at.normalize_frequency     AS compare_frequency,

    p.total_amount,
    p.our_investment,
    p.outstanding_principal,
    p.current_rating,
    p.risk_trend,

    -- 最新累计回收率
    (SELECT c.recovery_rate FROM cashflows c
     WHERE c.product_id = p.id AND c.is_paid = TRUE
     ORDER BY c.end_date DESC LIMIT 1) AS latest_recovery_rate,

    -- 最新预测
    (SELECT f.predicted_final_rate FROM forecasts f
     WHERE f.product_id = p.id
     ORDER BY f.forecast_date DESC LIMIT 1) AS latest_predicted_rate,

    -- 总实际回收
    COALESCE((SELECT SUM(c.actual_amount) FROM cashflows c
     WHERE c.product_id = p.id AND c.is_paid = TRUE), 0) AS total_recovered,

    -- 下期待回款
    (SELECT c.projected_amount FROM cashflows c
     WHERE c.product_id = p.id AND c.is_paid = FALSE
     ORDER BY c.start_date LIMIT 1) AS next_payment,

    p.status

FROM products p
JOIN banks b          ON p.bank_id = b.id
JOIN asset_types at   ON p.asset_type_id = at.id
WHERE p.is_invested = TRUE;

-- ═══════════════════════════════════════════════════════════
-- 示例: 查询标准化后的现金流对比
-- ═══════════════════════════════════════════════════════════
-- SELECT
--   normalized_period,
--   asset_type_name,
--   compare_frequency,
--   SUM(normalized_actual)   AS total_actual,
--   SUM(normalized_projected) AS total_projected,
--   ROUND((SUM(normalized_actual) - SUM(normalized_projected))
--         / NULLIF(SUM(normalized_projected), 0) * 100, 2) AS deviation_pct
-- FROM normalized_cashflows
-- WHERE is_paid = TRUE
-- GROUP BY normalized_period, asset_type_name, compare_frequency
-- ORDER BY normalized_period;
