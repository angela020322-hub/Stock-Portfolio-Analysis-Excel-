# Table of Contents
- [Stock Portfolio Analysis (Excel)](#stock-portfolio-analysis-excel)
- [Data Disclaimer](#data-disclaimer)
- [Overview](#overview)
- [Project Scope / Analytical Coverage](#project-scope-analytical-coverage)
- [Project Results](#project-results)
  - [Transaction & Portfolio Activity](#transaction--portfolio-activity)
  - [Market Data & Price Monitoring](#market-data--price-monitoring)
  - [Stock Screening & Selection](#stock-screening--selection)
  - [Link to Excel Model](#link-to-excel-model)
  - [Link to Project PDF](#link-to-project-pdf)
- [Stakeholders & Potential Applications](#stakeholders--potential-applications)
- [Workbook Overview](#workbook-overview)
  - [Transaction History](#transaction-history)
  - [Real Time Market Price](#real-time-market-price)
  - [Search Engine](#search-engine)
- [Vocabulary](#vocabulary)
  - [Transaction Vocabulary](#transaction-vocabulary)
  - [Portfolio Vocabulary](#portfolio-vocabulary)
  - [Performance & Risk Vocabulary](#performance--risk-vocabulary)
  - [Market & Classification Vocabulary](#market--classification-vocabulary)
  - [Sector Definition](#sector-definition)
- [Formulas & Functions](#formulas--functions)
  - [Formulas](#formulas)
  - [Functions](#functions)
  - [Function Logistics](#function-logistics)
# Stock Portfolio Analysis (Excel)
Portfolio information may be distributed across transaction records, market prices, holdings, and company-level data, making it difficult to obtain a consistent view of portfolio performance and risk.
# Data Disclaimer
All stock transaction data in this project are entirely illustrative and created solely for demonstration purposes. The transactions do not represent actual personal investments, trading activity, or brokerage records. They are used to demonstrate the analytical model, calculation logic, and Excel capabilities of the project.
# Overview
Portfolio information may be distributed across transaction records,  market prices, holdings, and company-level data, making it difficult  to obtain a consistent view of portfolio performance and risk.  This project is an Excel-based Stock Portfolio Analysis Model that  transforms transaction and market data into portfolio holdings,  cost basis, P&L, performance, allocation, and risk metrics.  The project follows an analyst-oriented workflow: defining business  questions, structuring data requirements, building analytical logic,  calculating KPIs, and communicating findings through dashboards.  
# Project Scope/ Analytical Coverage
The project covers the following areas:
1) **Transaction Tracking** — Records buy and sell activities, including transaction dates, quantities, prices, and transaction values.
2) **Portfolio Holdings** — Calculates current holdings, position sizes, market values, and portfolio weights.
3) **Cost Basis** — Tracks remaining cost basis based on transaction activity.
4) **P&L Analysis** — Calculates realized and unrealized P&L at the security and portfolio levels.
5) **Performance Analysis** — Evaluates portfolio and security-level returns and benchmark performance.
6) **Portfolio Allocation** — Analyzes exposure across individual securities and sectors.
7) **Risk Analysis** — Evaluates selected portfolio risk metrics, concentration, and exposure.
8) **Dashboard Reporting** — Consolidates key portfolio KPIs and analytical results into an interactive dashboard.
9) **Excel Automation** — Uses Excel formulas, PivotTables, and VBA to automate calculations and repetitive reporting tasks.
# Project Results
## Transaction & Portfolio Activity
Tracks buy and sell transactions and provides the underlying activity used for portfolio analysis.
<img width="1362" height="509" alt="Screenshot (9)" src="https://github.com/user-attachments/assets/54714058-35b1-48f7-b7a9-fa033aaa9e50" />
## Market Data & Price Monitoring
Provides market price information used to calculate portfolio value, P&L, and performance.
<img width="1366" height="403" alt="Screenshot (8)" src="https://github.com/user-attachments/assets/945558e1-1da4-410b-8a74-6fcd89fd5584" />
## Stock Screening & Selection
Screens securities based on selected criteria to support stock analysis and selection.
<img width="1366" height="504" alt="Screenshot (7)" src="https://github.com/user-attachments/assets/43e820ec-663c-4ee4-94c8-fb4eb28199f4" />
## Link to Excel Model
[Stock Analyst Excel Model.xlsx](https://github.com/user-attachments/files/31810486/Stock.Analyst.Excel.Model.xlsx)
## Link to Project PDF

# Stakeholders & Potential Applications
The model is designed around different stakeholder information needs and can be adapted to portfolio monitoring, investment reporting, and financial analysis use cases.
<img width="1294" height="579" alt="Screenshot (11)" src="https://github.com/user-attachments/assets/33030f4d-ab58-449d-bb69-45f6e6438b31" />
# Workbook Overview
The workbook is organized into three main worksheets, moving from transaction-level portfolio information to market data and interactive security-level analysis.
## Transaction History
Provides a historical view of portfolio transactions and summarizes investment activity and portfolio diversification.
1) **Transaction History Chart** — Visualizes historical buying and selling activity over time.
2) **Investment Summary** — Summarizes key investment information and portfolio-level metrics.
3) **Investment Diversity Chart** — Shows how investments are distributed across different securities.
4) **Investment Diversity Diagram** — Provides a visual representation of portfolio diversification.
## Real Time Market Price
Provides the latest available market price for each ticker currently included in the portfolio.
1) **Real-Time Price by Ticker** — Displays the current/reference market price for each portfolio security and provides the price inputs used for portfolio valuation and gain/loss analysis.
## Search Engine
Provides an interactive security-level analysis tool that allows users to search for a specific ticker and review its portfolio position, purchase history, performance, allocation, and market information.
1) **Search Area** — Allows users to select or enter a ticker for detailed analysis.
2) **Stock Summary** — Summarizes key information and current portfolio metrics for the selected security.
3) **Purchase History** — Displays the historical purchase activity for the selected ticker.
4) **Current Gain & Loss Chart** — Visualizes the current gain or loss associated with the selected security.
5) **Portfolio Weight (%) Chart** — Shows the selected security's percentage of total portfolio value.
6) **Portfolio Weight (%) Diagram** — Provides a visual representation of the selected security's contribution to the overall portfolio.
7) **Sector Weight (%) Chart** — Shows portfolio exposure by sector.
8) **Sector Weight (%) Diagram** — Provides a visual representation of sector diversification.
9) **Real-Time Price Summary** — Displays the latest/reference market price and related information for the selected ticker.
10) **Real-Time Price Chart** — Visualizes the selected ticker's market price information.
11) **Real-Time Price Diagram** — Provides an additional visual representation of the selected security's current market price.
# Vocabulary
## Transaction Vocabulary
1) ***Ticker***<br>
**Definition:** A unique symbol used to identify a publicly traded stock.<br>
**Example from Project:** Used to identify and track the stock being transacted and to connect transaction, market-price, and stock-summary data.<br>
2) ***Date***<br>
**Definition:** The date on which a transaction occurs. <br>
**Example from Project:** Used to record the timing of each Buy or Sell transaction and to calculate the holding period of a stock. <br>
4) ***Shares***<br>
**Definition:** The number of shares bought or sold in a transaction. <br>
**Example from Project:** Used to calculate Total Cost, current shares owned, and Current Value based on the transaction history. <br>
5) ***Unit Price*** <br>
**Definition:** The price paid or received for one share in a transaction. <br>
**Example from Project:** Used with Shares to calculate the Total Cost of each transaction and to determine the cost basis of holdings. <br>
6) ***Total Cost*** <br>
**Definition:** The total dollar amount associated with a transaction, calculated from the number of shares and the price per share.<br>
**Example from Project:** Calculated as Shares × Unit Price and used to track the amount invested through transaction activity. <br>
7) ***Buy*** <br>
**Definition:** A transaction in which shares of a stock are purchased. <br>
**Example from Project:** Increases the number of shares owned and adds the transaction cost to the Running Cost Basis. <br>
8) ***Sell*** <br>
**Definition:** A transaction in which shares of a stock are sold. <br>
**Example from Project:** Decreases the number of shares owned and reduces the Running Cost Basis associated with the position. <br>
9) ***Transaction*** <br>
**Definition:** A recorded purchase or sale of a financial asset. <br>
**Example from Project:** Each row in the Transaction History table represents one Buy or Sell transaction and provides the underlying data for portfolio calculations. <br>
## Portfolio Vocabulary
1) ***Portfolio*** <br>
**Definition:** A collection of financial assets held by an investor or investment entity. <br>
**Example from Project:** Represents the complete set of stocks and other assets tracked by the Excel model and is used for overall performance and allocation analysis. <br>
2) ***Holding*** <br>
**Definition:** An asset or number of shares currently owned in a portfolio. <br>
**Example from Project:** Used to determine how many shares of each stock remain after accounting for all Buy and Sell transactions. <br>
3) ***Position*** <br>
**Definition:** The amount of a particular security held within a portfolio. <br>
**Example from Project:** Each stock position is analyzed using its shares owned, invested amount, current value, P&L, and portfolio weight. <br>
4) ***Total Invested*** <br>
**Definition:** The total amount of capital invested in a stock or portfolio based on the model's investment calculation. <br>
**Example from Project:** Used as the investment basis for comparing the amount invested with the Current Value and calculating Gain/Loss. <br>
5) ***Current Value*** <br>
**Definition:** The current market value of an asset or position based on its current price and quantity held. <br>
**Example from Project:** Calculated using Current Price × Stock Owned to determine the current value of each stock position. <br>
6) ***Market Value*** <br>
**Definition:** The value of an asset or position based on its current market price. <br>
**Example from Project:** Used to represent the current financial value of a stock position and to support portfolio allocation analysis. <br>
7) ***Portfolio Weight*** <br>
**Definition:** The percentage of a portfolio's total market value represented by a particular position. <br>
**Example from Project:** Calculated as Stock Current Value ÷ Current Portfolio Value to measure the relative size of each position. <br>
8) ***Portfolio Allocation*** <br>
**Definition:** The distribution of a portfolio's value across individual assets, sectors, or other categories. <br>
**Example from Project:** Used to analyze how the portfolio is distributed across different stocks and sectors and to identify concentration. <br>
## Performance & Risk Vocabulary
1) ***P&L (Profit & Loss)*** <br>
**Definition:** The financial gain or loss generated by an investment or portfolio over a specified period or based on its current value. <br>
**Example from Project:** Used to evaluate the financial performance of individual stock positions and the overall portfolio. <br>
2) ***Gain/Loss in Dollars*** <br>
**Definition:** The absolute dollar amount of profit or loss generated by an investment. <br>
**Example from Project:** Calculated as Current Value − Invested Amount and used to compare the dollar contribution of different positions. <br>
3) ***Gain/Loss in Percentage*** <br>
**Definition:** The gain or loss expressed as a percentage of the invested amount. <br>
**Example from Project:** Calculated as (Current Value − Invested Amount) ÷ Invested Amount to compare performance across positions with different investment sizes. <br>
4) ***Realized P&L*** <br>
**Definition:** The profit or loss resulting from a position that has been sold. <br>
**Example from Project:** Used to evaluate the financial outcome of completed Sell transactions. <br>
5) ***Unrealized P&L*** <br>
**Definition:**The profit or loss on an asset that is still being held and has not yet been sold. <br>
**Example from Project:** Used to evaluate the current performance of open positions based on their current market value. <br>
6) ***Return*** <br>
**Definition:** The gain or loss from an investment expressed relative to the capital invested. <br>
**Example from Project:** Used as a percentage-based performance measure to evaluate the return generated by a stock or portfolio. <br>
7) ***ROC (Rate of Change)*** <br>
**Definition:** A percentage measure of how much the price or value of an asset has changed between two points in time. <br>
**Example from Project:** Used as a market-performance indicator to evaluate the relative price movement of a stock. <br>
8) ***Beta*** <br>
**Definition:** ** A measure of a stock's sensitivity to movements in the overall market. <br>
**Example from Project:** Used as a risk indicator to evaluate the market sensitivity of individual stocks within the portfolio. <br>
## Market & Classification Vocabulary
1) ***Sector*** <br>
**Definition:** A broad classification that groups companies based on the primary type of business they operate. <br>
**Example from Project:** Used to classify stocks and analyze portfolio allocation and exposure across sectors. <br>
2) ***Industry*** <br>
**Definition:** A more specific business classification describing the type of products or services a company provides. <br>
**Example from Project:** Used to provide more detailed business classification when analyzing individual stocks. <br>
3) ***Current Price*** <br>
**Definition:** The most recent market price available for a stock. <br>
**Example from Project:** Used to calculate Current Value, Gain/Loss, and portfolio-level valuation. <br>
4) ***52-Week High*** <br>
**Definition:** The highest market price reached by a stock during the previous 52 weeks. <br>
**Example from Project:** Used to provide context for the stock's current price relative to its recent annual trading range. <br>
5) ***52-Week Low*** <br>
**Definition:** The lowest market price reached by a stock during the previous 52 weeks. <br>
**Example from Project:** Used to provide context for the stock's current price relative to its recent annual trading range. <br>
6) ***P/E Ratio*** <br>
**Definition:** A valuation ratio comparing a company's stock price with its earnings per share. <br>
**Example from Project:** Used as an additional market and valuation indicator when reviewing individual stocks. <br>
7) ***Benchmark*** <br>
**Definition:** A standard market index or reference used to evaluate the performance of an investment or portfolio. <br>
**Example from Project:** Can be used to compare portfolio performance against a broader market reference such as the S&P 500. <br>
## Sector Definition
1) ***Information Technology (IT)*** <br>
**Definition:** Companies primarily involved in software, hardware, semiconductors, technology services, and related technology products. <br>
2) ***Health Care*** <br>
**Definition:**Companies involved in pharmaceuticals, biotechnology, medical devices, healthcare services, and related products. <br>
3) ***Financials*** <br>
**Definition:** Companies involved in banking, insurance, investment, lending, and other financial services. <br>
4) ***Consumer Discretionary*** <br>
**Definition:** Companies providing non-essential consumer goods and services, including retail, automobiles, travel, entertainment, and leisure. <br>
5) ***Consumer Staples*** <br>
**Definition:** Companies providing essential consumer products such as food, beverages, household goods, and personal care products. <br>
6) ***Communication Services*** <br>
**Definition:** Companies providing essential consumer products such as food, beverages, household goods, and personal care products. <br>
7) ***Industrials*** <br>
**Definition:** Companies involved in manufacturing, machinery, transportation, aerospace, construction, and industrial services. <br>
8) ***Energy*** <br>
**Definition:** Companies involved in energy production, exploration, distribution, equipment, and related services. <br>
9) ***Utilities*** <br>
**Definition:** Companies providing essential services such as electricity, natural gas, water, and other utility services. <br>
10) ***Materials*** <br>
**Definition:** Companies involved in chemicals, metals, mining, construction materials, and other raw materials. <br>
11) ***Real Estate*** <br>
**Definition:** Companies involved in real estate ownership, development, management, and related services, including REITs. <br>
12) ***Crypto*** <br>
**Definition:** Digital assets that use blockchain technology, such as cryptocurrencies, rather than representing traditional equity sectors. <br>
# Formulas & Functions

## Formulas

<img width="1112" height="572" alt="Screenshot (12)" src="https://github.com/user-attachments/assets/38c92cd8-3608-44f6-b385-65708379a329" />

## Functions
## Functions & Formulas

1. **IF** — Returns one value when a condition is TRUE and another value when it is FALSE.
2. **IFERROR** — Returns a specified value when a formula results in an error.
3. **LET** — Assigns names to values or intermediate calculations within a formula.
4. **FILTER** — Filters a range or array based on specified criteria.
5. **SCAN** — Performs a calculation sequentially across an array and returns each intermediate result.
6. **SEQUENCE** — Generates a sequence of consecutive numbers in an array.
7. **LAMBDA** — Creates custom reusable functions using defined parameters and calculation logic.
8. **SPLIT** — Divides text into separate values based on a specified delimiter.
9. **VALUE** — Converts text that represents a number into a numeric value.
10. **INDEX** — Returns a value from a specified position within a range or array.
11. **ROWS** — Returns the number of rows in a range or array.
12. **MAP** — Applies a specified calculation to each value in an array.
13. **UNIQUE** — Returns a list of unique values from a range or array.
14. **SUM** — Adds numbers together and returns their total.
15. **SUMPRODUCT** — Multiplies corresponding values in arrays and returns the sum of the resulting products.
16. **MINIFS** — Returns the smallest value that meets one or more specified criteria.
17. **TODAY** — Returns the current date.
18. **DATEDIF** — Calculates the difference between two dates in specified units.
19. **SORT** — Sorts the values in a range or array based on specified criteria.
20. **XLOOKUP** — Searches for a value and returns the corresponding value from another range or array.
21. **MAX** — Returns the largest value in a range or array.
22. **GOOGLEFINANCE** — Retrieves financial market and security information from Google Finance.
23. **SUMIF** — Adds values that meet a specified condition.
24. **MAXIFS** — Returns the largest value that meets one or more specified criteria.
## Function Logistics
### Function 01: Running Cost Basis
```
=IF(B2="","",
LET(
t,B2,
shares,FILTER($C$2:C2,$B$2:B2=t),
prices,FILTER($D$2:D2,$B$2:B2=t),
states,SCAN("0|0",SEQUENCE(ROWS(shares)),
LAMBDA(acc,i,
LET(
x,SPLIT(acc,"|"),
oldShares,VALUE(INDEX(x,1,1)),
oldBasis,VALUE(INDEX(x,1,2)),
s,INDEX(shares,i),
p,INDEX(prices,i),
newShares,oldShares+s,
newBasis,
IF(
s>0,
oldBasis+s*p,
IF(
newShares<=0,
0,
oldBasis+s*(oldBasis/oldShares)
)
),
newShares&"|"&newBasis
))),
last,SPLIT(INDEX(states,ROWS(states)),"|"),
VALUE(INDEX(last,1,2))
))
```
**Purpose** <br>

Calculate the running total cost basis for each stock ticker by tracking purchases and sales sequentially. <br>

**Logic**
1) Identify the current Ticker — LET <br>
Stores the ticker from B2 as variable t.
2) Retrieve transactions for the Ticker — FILTER <br>
Extracts the relevant Shares and Unit Price records for the current ticker.
3) Create transaction sequence — ROWS + SEQUENCE <br>
Counts the filtered transactions and generates a sequential index for each one.
4) Process transactions sequentially — SCAN + LAMBDA <br>
Calculates the running Shares and Cost Basis while carrying the previous result forward.
5) Retrieve previous Shares and Cost Basis — SPLIT + INDEX + VALUE <br>
Separates the previous running values and converts them into numbers.
6) Retrieve current Shares and Unit Price — INDEX <br>
Gets the Shares and Price for the current transaction.
7) Calculate new Shares — + <br>
Adds the current transaction Shares to the previous Shares.
8) Handle Buy vs. Sell — IF <br>
Determines whether to add purchase cost or reduce Cost Basis.
9) Calculate Buy Cost Basis — + + * <br>
Adds Shares × Unit Price to the previous Cost Basis.
10) Calculate Sell Cost Basis — / + * + + <br>
Uses the current average cost per share to reduce the Cost Basis.
11) Reset fully sold positions — IF <br>
Sets Cost Basis to 0 when remaining Shares are ≤ 0.
12) Store the running result — & <br>
Combines Shares and Cost Basis into one value for the next SCAN iteration.
13) Return the latest Cost Basis — INDEX + SPLIT + VALUE <br>
Retrieves the final running state and returns the latest Cost Basis.

**In Short** <br>
- LET → identify & organize variables
- FILTER → retrieve ticker transactions
- ROWS / SEQUENCE → create transaction sequence
- SCAN / LAMBDA → process transactions sequentially
- SPLIT / INDEX / VALUE → retrieve running values
- IF → handle Buy / Sell / fully sold
- SUMMATION / MULTIPLICATION / DIVISION → calculate Shares & Cost Basis
- & → carry the running state
- INDEX / SPLIT / VALUE → return final Cost Basis <br>

### Function 02: Total Invested
```
=SUM(MAP(UNIQUE(FILTER('Transection History'!B2:B,'Transection History'!B2:B<>"")),LAMBDA(t,INDEX(FILTER('Transection History'!G2:G,'Transection History'!B2:B=t),ROWS(FILTER('Transection History'!G2:G,'Transection History'!B2:B=t))))))
```

**Purpose** <br>

Calculates the total invested amount by summing the latest total cost for each unique stock ticker. <br>

**Logic** <br>
1) Identify unique Tickers — FILTER + UNIQUE <br>
Retrieves all non-blank tickers and creates a unique list. <br>
2) Process each Ticker — MAP + LAMBDA <br>
Iterates through each unique ticker individually. <br>
3) Retrieve all Total Cost values — FILTER <br>
Filters the Total Cost column (G) for the current ticker. <br>
4) Identify the latest transaction — ROWS + INDEX <br>
Uses the number of filtered transactions to retrieve the last Total Cost value.
5) Sum the latest value for each Ticker — SUM <br>
Adds the latest Total Cost from every unique ticker to calculate Total Invested. <br>

**In Short** <br>
- FILTER / UNIQUE → identify unique tickers <br>
- MAP / LAMBDA → process each ticker <br>
- FILTER → retrieve its transaction costs <br>
- ROWS / INDEX → get the latest Total Cost <br>
- SUM → calculate Total Invested <br>

### Function 03: Amount Invested <br>
```
=IFERROR(
LET(
sector,L26,
tickers,UNIQUE(
FILTER(
'Transection History'!B$2:B,
'Transection History'!F$2:F=sector,
'Transection History'!B$2:B<>""
)
),
tickerCosts,MAP(
tickers,
LAMBDA(ticker,
LET(
shares,FILTER('Transection History'!C$2:C998,'Transection History'!B$2:B=ticker),
costs,FILTER('Transection History'!E$2:E998,'Transection History'!B$2:B=ticker),
n,ROWS(shares),
idx,SEQUENCE(n),
runningShares,SCAN(0,idx,LAMBDA(acc,i,acc+INDEX(shares,i))),
runningCost,SCAN(0,idx,LAMBDA(acc,i,
LET(
prevShares,IF(i=1,0,INDEX(runningShares,i-1)),
q,INDEX(shares,i),
cost,INDEX(costs,i),
IF(
q>0,
acc+cost,
IF(
prevShares=0,
0,
acc+q*(acc/prevShares)
)
)
))),
INDEX(runningCost,n)
)
)
),
SUM(tickerCosts)
),
0)
```
**Purpose** <br>

Calculates the total amount of invested for a selected sector by determine the remaining cost basis of each stock in that sector and summing them. <br>

**Logic** <br>
1) Identify the selected Sector — LET <br>
Stores the sector from L26 as sector. <br>
2) Identify unique Tickers within the Sector — FILTER + UNIQUE <br>
Filters the Transaction History by sector and removes blank tickers and duplicates. <br>
3) Process each Ticker separately — MAP + LAMBDA <br>
Calculates the invested amount for each stock in the sector. <br>
4) Retrieve each Ticker's transactions — FILTER <br>
Retrieves its Shares and transaction Costs. <br>
5) Create transaction sequence — ROWS + SEQUENCE <br>
Determines the number and order of transactions. <br>
6) Calculate Running Shares — SCAN + LAMBDA + INDEX
Sequentially adds each transaction's Shares to determine the remaining Shares Owned.
7) Calculate Running Cost Basis — SCAN + LAMBDA + IF + INDEX <br>
Sequentially updates Cost Basis based on each Buy or Sell transaction. <br>
8) Handle Buy transactions — IF + + <br>
Adds the transaction Cost to the previous Cost Basis. <br>
9) Handle Sell transactions — IF + / + * + + <br>
Reduces Cost Basis using the current average cost per share. <br>
10) Handle fully sold positions — IF <br>
Sets Cost Basis to 0 when the previous Shares Owned are 0. <br>
11) Return each Ticker's final Cost Basis — INDEX <br>
Retrieves the last Running Cost Basis for each ticker. <br>
12) Calculate Sector Amount Invested — SUM <br>
Adds the final Cost Basis of all tickers in the selected sector. <br>
13) Handle calculation errors — IFERROR <br>
Returns 0 if the calculation produces an error. <br>

**In Short** <br>
- LET → identify Sector <br>
- FILTER / UNIQUE → identify Sector's unique Tickers <br>
- MAP / LAMBDA → calculate each Ticker <br>
- FILTER → retrieve transactions <br>
- ROWS / SEQUENCE → create transaction sequence <br>
- SCAN / LAMBDA → calculate Running Shares & Cost Basis <br>
- IF → handle Buy / Sell / fully sold positions <br>
- INDEX → retrieve transaction & final values <br>
- SUM → calculate total Amount Invested <br>
- IFERROR → return 0 if an error occurs <br>

### Function 04: Real Time Value
```
=IFERROR(SUMPRODUCT(('Real Time Market Price'!$G$2:$G998=L26)*('Real Time Market Price'!$C$2:$C998)*('Real Time Market Price'!$B$2:$B)),0)
```

**Purpose** <br>

Calculates the current market value of all stocks within the selected sector using real-time market prices and current holdings. <br>
<br>

**Logic** <br>

1) Identify the selected Sector — L26 <br>
Uses the sector listed in L26 as the filtering criterion. <br>
2) Filter stocks by Sector — = <br>
Checks whether each stock's sector matches the selected sector. <br>
3) Retrieve Current Price — C <br>
Uses the current market price for each matching stock. <br>
4) Retrieve Shares Owned — B <br>
Uses the current shares held for each matching stock. <br>
5) Calculate each Stock's Current Value — * <br>
Multiplies Current Price × Shares Owned. <br>
6) Calculate Sector Real Time Value — SUMPRODUCT <br>
Sums the Current Value of all stocks belonging to the selected sector. <br>
7) Handle calculation errors — IFERROR <br>
Returns 0 if the calculation produces an error. <br>

**In Short** <br>

- = → identify stocks in the selected sector <br>
- C × B → calculate each stock's current value <br>
- SUMPRODUCT → total the sector's real-time value <br>
- IFERROR → return 0 if an error occurs <br>

### Function 05: Purchase History Tracker
```
=SORT(
  FILTER(
    'Transection History'!A2:G1000,
    'Transection History'!B2:B=$B$1
  ),
  1,
  True
)
```
**Purpose** <br>

<br>

**Logic** <br>

<br>

