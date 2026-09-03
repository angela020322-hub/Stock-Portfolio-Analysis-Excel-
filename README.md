# Table of Contents
1. [Stock Portfolio Analysis (Excel)](#stock-portfolio-analysis-excel)
2. [Data Disclaimer](#data-disclaimer)
3. [Overview](#overview)
4. [Project Scope / Analytical Coverage](#project-scope-analytical-coverage)
5. [Project Results](#project-results)
6. [Stakeholders & Potential Applications](#stakeholders--potential-applications)
7. [Workbook Overview](#workbook-overview)
8. [Vocabulary](#vocabulary)

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
[Stock_Analyst_Excel_Model.xlsx](https://github.com/user-attachments/files/31704987/Stock_Analyst_Excel_Model.xlsx)
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
**Definition:**
**Example from Project:**
2) ***Holding*** <br>
**Definition:**
**Example from Project:**
3) ***Position*** <br>
**Definition:**
**Example from Project:**
4) ***Total Invested*** <br>
**Definition:**
**Example from Project:**
5) ***Current Value*** <br>
**Definition:**
**Example from Project:**
6) ***Market Value*** <br>
**Definition:**
**Example from Project:**
7) ***Portfolio Weight*** <br>
**Definition:**
**Example from Project:**
8) ***Portfolio Allocation*** <br>
**Definition:**
**Example from Project:**
## Performance & Risk Vocabulary
1) ***P&L (Profit & Loss)*** <br>
**Definition:**
**Example from Project:**
2) ***Gain/Loss in Dollars*** <br>
**Definition:**
**Example from Project:**
3) ***Gain/Loss in Percentage*** <br>
**Definition:**
**Example from Project:**
4) ***Realized P&L*** <br>
**Definition:**
**Example from Project:**
5) ***Unrealized P&L*** <br>
**Definition:**
**Example from Project:**
6) ***Return*** <br>
**Definition:**
**Example from Project:**
7) ***ROC (Rate of Change)*** <br>
**Definition:**
**Example from Project:**
8) ***Beta*** <br>
**Definition:**
**Example from Project:**
## Market & Classification Vocabulary
1) ***Sector*** <br>
**Definition:**
**Example from Project:**
2) ***Industry*** <br>
**Definition:**
**Example from Project:**
3) ***Current Price*** <br>
**Definition:**
**Example from Project:**
4) ***52-Week High*** <br>
**Definition:**
**Example from Project:**
5) ***52-Week Low*** <br>
**Definition:**
**Example from Project:**
6) ***P/E Ratio*** <br>
**Definition:**
**Example from Project:**
7) ***Benchmark*** <br>
**Definition:**
**Example from Project:**
## Sector Definition
1) ***Information Technology (IT)*** <br>
**Definition:**
**Example from Project:**
2) ***Health Care*** <br>
**Definition:**
**Example from Project:**
3) ***Financials*** <br>
**Definition:**
**Example from Project:**
4) ***Consumer Discretionary*** <br>
**Definition:**
**Example from Project:**
5) ***Consumer Staples*** <br>
**Definition:**
**Example from Project:**
6) ***Communication Services*** <br>
**Definition:**
**Example from Project:**
7) ***Industrials*** <br>
**Definition:**
**Example from Project:**
8) ***Energy*** <br>
**Definition:**
**Example from Project:**
9) ***Utilities*** <br>
**Definition:**
**Example from Project:**
10) ***Materials*** <br>
**Definition:**
**Example from Project:**
11) ***Real Estate*** <br>
**Definition:**
**Example from Project:**
12) ***Crypto*** <br>
**Definition:**
**Example from Project:**
