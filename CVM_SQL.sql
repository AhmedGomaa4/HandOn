-- Aggregating sales data for specific conditions
WITH t1 AS (
    SELECT
        SUM(Net_Sales) AS total_sales,
        SUM(Net_Sales - Cost + compensation) AS gross_profit,
        SUM(CASE WHEN [Customer_Type] = 'installment' THEN [Net_Sales] ELSE 0 END) AS MC_Sales,
        [Customer].[group] AS customer_group,
        COUNT(DISTINCT [RMS_SettlementTransKey]) AS num_invoices,
        Customer_Phone
    FROM sales
    INNER JOIN customer ON Customer.customer_id = sales.Customer_Phone
    WHERE
        -- Filter by Customer_Type ('cash' or 'installment')
        Customer_Type IN ('cash', 'installment')
        
        -- Filter by Budget_Channel
        AND Budget_Channel IN ('Arkan', 'branches', 'btechx', 'cc br', 'online br')
        
        -- Filter by Quantity within the range of 1 and 2
        AND [Quantity] BETWEEN 1 AND 2
        
        -- Filter by Sold_Date within the specified date range
        AND [Sold_Date] BETWEEN '2023-06-12' AND '2023-06-25'
        
    GROUP BY [Customer].[group], Customer_Phone
    HAVING
        -- Additional filtering conditions based on aggregated data
        SUM(Quantity) < 7
        AND COUNT(DISTINCT [RMS_SettlementTransKey]) < 4
),

-- Summarizing t1 data for 'control' group
t2 AS (
    SELECT
        [group],
        ROUND(SUM(total_sales), 0) AS total_sales,
        ROUND(SUM(gross_profit), 0) AS gross_profit,
        ROUND(SUM(MC_Sales), 0) AS MC_Sales,
        SUM(num_invoices) AS num_invoices
    FROM t1
    GROUP BY [group]
    HAVING [group] = 'control'
)
-- Final query: Selecting data from t2
SELECT * FROM t2;
