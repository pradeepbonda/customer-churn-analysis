-- Customer Churn Analysis
-- Table: customers (loaded from data/telco_customer_churn.csv by src/run_analysis.py)
-- Each query starts with a "-- name:" line. The runner saves each result to
-- outputs/query_results/<name>.csv

-- name: 01_overall_churn
-- How many customers do we have, and what share left?
SELECT
    COUNT(*)                                             AS customers,
    SUM(CASE WHEN churn = 'Yes' THEN 1 ELSE 0 END)       AS churned,
    ROUND(100.0 * AVG(CASE WHEN churn = 'Yes' THEN 1.0 ELSE 0 END), 1) AS churn_rate_pct
FROM customers;

-- name: 02_churn_by_contract
-- Contract type is the strongest single driver.
SELECT
    contract,
    COUNT(*)                                             AS customers,
    SUM(CASE WHEN churn = 'Yes' THEN 1 ELSE 0 END)       AS churned,
    ROUND(100.0 * AVG(CASE WHEN churn = 'Yes' THEN 1.0 ELSE 0 END), 1) AS churn_rate_pct
FROM customers
GROUP BY contract
ORDER BY churn_rate_pct DESC;

-- name: 03_churn_by_tenure_band
-- New customers leave far more often than long-standing ones.
SELECT
    CASE
        WHEN tenure <= 12 THEN '1. 0-12 months'
        WHEN tenure <= 24 THEN '2. 13-24 months'
        WHEN tenure <= 48 THEN '3. 25-48 months'
        ELSE                   '4. 49-72 months'
    END                                                  AS tenure_band,
    COUNT(*)                                             AS customers,
    SUM(CASE WHEN churn = 'Yes' THEN 1 ELSE 0 END)       AS churned,
    ROUND(100.0 * AVG(CASE WHEN churn = 'Yes' THEN 1.0 ELSE 0 END), 1) AS churn_rate_pct
FROM customers
GROUP BY tenure_band
ORDER BY tenure_band;

-- name: 04_churn_by_internet_service
SELECT
    internet_service,
    COUNT(*)                                             AS customers,
    ROUND(100.0 * AVG(CASE WHEN churn = 'Yes' THEN 1.0 ELSE 0 END), 1) AS churn_rate_pct,
    ROUND(AVG(monthly_charges), 2)                       AS avg_monthly_charge
FROM customers
GROUP BY internet_service
ORDER BY churn_rate_pct DESC;

-- name: 05_churn_by_payment_method
SELECT
    payment_method,
    COUNT(*)                                             AS customers,
    ROUND(100.0 * AVG(CASE WHEN churn = 'Yes' THEN 1.0 ELSE 0 END), 1) AS churn_rate_pct
FROM customers
GROUP BY payment_method
ORDER BY churn_rate_pct DESC;

-- name: 06_support_services_effect
-- Customers with internet who do or do not have tech support and online security.
SELECT
    tech_support,
    online_security,
    COUNT(*)                                             AS customers,
    ROUND(100.0 * AVG(CASE WHEN churn = 'Yes' THEN 1.0 ELSE 0 END), 1) AS churn_rate_pct
FROM customers
WHERE internet_service <> 'No'
GROUP BY tech_support, online_security
ORDER BY churn_rate_pct DESC;

-- name: 07_contract_x_tenure
-- Combining the two strongest drivers.
SELECT
    contract,
    CASE WHEN tenure <= 12 THEN '0-12 months' ELSE '13+ months' END AS tenure_group,
    COUNT(*)                                             AS customers,
    ROUND(100.0 * AVG(CASE WHEN churn = 'Yes' THEN 1.0 ELSE 0 END), 1) AS churn_rate_pct
FROM customers
GROUP BY contract, tenure_group
HAVING COUNT(*) >= 50
ORDER BY churn_rate_pct DESC;

-- name: 08_high_risk_segment
-- Month-to-month, first year, fiber optic: small share of customers, large share of churn.
SELECT
    COUNT(*)                                                         AS segment_customers,
    ROUND(100.0 * COUNT(*) / (SELECT COUNT(*) FROM customers), 1)    AS pct_of_all_customers,
    ROUND(100.0 * AVG(CASE WHEN churn = 'Yes' THEN 1.0 ELSE 0 END), 1) AS segment_churn_rate_pct,
    SUM(CASE WHEN churn = 'Yes' THEN 1 ELSE 0 END)                   AS segment_churned,
    ROUND(100.0 * SUM(CASE WHEN churn = 'Yes' THEN 1 ELSE 0 END)
          / (SELECT COUNT(*) FROM customers WHERE churn = 'Yes'), 1) AS pct_of_all_churners
FROM customers
WHERE contract = 'Month-to-month'
  AND tenure <= 12
  AND internet_service = 'Fiber optic';

-- name: 09_monthly_revenue_lost
-- Monthly recurring revenue that walked out the door.
SELECT
    churn,
    COUNT(*)                                             AS customers,
    ROUND(SUM(monthly_charges), 0)                       AS monthly_revenue,
    ROUND(AVG(monthly_charges), 2)                       AS avg_monthly_charge
FROM customers
GROUP BY churn;

-- name: 10_above_average_payers
-- Subquery: churn among customers paying more than the average monthly charge.
SELECT
    CASE WHEN monthly_charges > (SELECT AVG(monthly_charges) FROM customers)
         THEN 'Above average charge' ELSE 'At or below average' END AS charge_group,
    COUNT(*)                                             AS customers,
    ROUND(100.0 * AVG(CASE WHEN churn = 'Yes' THEN 1.0 ELSE 0 END), 1) AS churn_rate_pct
FROM customers
GROUP BY charge_group
ORDER BY churn_rate_pct DESC;
