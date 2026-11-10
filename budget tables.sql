USE business_management_system_db;

-- Default packing / delivery cost (kisima item / method ekakata cost ekak nathnam use wenawa)
CREATE TABLE IF NOT EXISTS budget_settings (
    setting_key   VARCHAR(50) PRIMARY KEY,
    setting_value DECIMAL(12,2) NOT NULL DEFAULT 0
);

INSERT IGNORE INTO budget_settings (setting_key, setting_value) VALUES
    ('default_packing_cost', 0),
    ('default_delivery_cost', 0);

-- Item ekakata cost price (per unit) saha packing cost (per parcel)
CREATE TABLE IF NOT EXISTS item_costs (
    item_code    VARCHAR(50) PRIMARY KEY,
    cost_price   DECIMAL(12,2) NOT NULL DEFAULT 0,
    packing_cost DECIMAL(12,2) NULL
);

-- Delivery method ekakata delivery cost (courier / post office ...)
CREATE TABLE IF NOT EXISTS delivery_costs (
    method VARCHAR(50) PRIMARY KEY,
    cost   DECIMAL(12,2) NOT NULL DEFAULT 0
);

-- Parcel (order) ekakata wenama cost (default wala wenuwata)
CREATE TABLE IF NOT EXISTS parcel_costs (
    order_id      INT PRIMARY KEY,
    product_cost  DECIMAL(12,2) NULL,
    packing_cost  DECIMAL(12,2) NULL,
    delivery_cost DECIMAL(12,2) NULL,
    extra_cost    DECIMAL(12,2) NULL,
    updated_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

-- Dawasakata ad cost
CREATE TABLE IF NOT EXISTS daily_ad_costs (
    ad_date    DATE PRIMARY KEY,
    amount     DECIMAL(12,2) NOT NULL DEFAULT 0,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

-- Masayakata thiyena expenses saha savings (rental fee, electricity, savings ...)
CREATE TABLE IF NOT EXISTS monthly_expenses (
    id          INT AUTO_INCREMENT PRIMARY KEY,
    name        VARCHAR(100) NOT NULL,
    category    ENUM('Expense', 'Savings') NOT NULL DEFAULT 'Expense',
    amount      DECIMAL(12,2) NOT NULL,
    start_month DATE NOT NULL,
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Order ekakata items godak (server eka automatically hadanawa, mehema run karanna ona naha)
CREATE TABLE IF NOT EXISTS order_items (
    id           INT AUTO_INCREMENT PRIMARY KEY,
    order_id     INT NOT NULL,
    item_code    VARCHAR(50)  DEFAULT '',
    product_name VARCHAR(150) DEFAULT '',
    quantity     INT          DEFAULT 1,
    price        DECIMAL(10,2) DEFAULT 0,
    INDEX idx_order (order_id)
);