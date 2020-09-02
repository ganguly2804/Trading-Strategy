# Trading-Strategy
An excel and VBA based strategical trading algorithm starter project 

Contents:
AlgoTrade.xlsm - The main excel file
Pull Stocks Data.xlsm - An excel file demonstrating fetching of stocks data (useful for older excel versions)

Requirements:
APIBridge
MS-Excel
Additional Excel Plug-in
(Instructions: https://mycoder.pro/apibridge/jump-start-system-trading-with-excel/)

Excel version used: Microsoft Excel for Office 365
Steps:
1. Start trading on APIBridge
2. Setup APIBridge Signals in Excel
3. Change the default formula for A-SigType to =IF(I2, "LX", IF(J2,"SX",IF(G2, "LE", IF(H2,"SE","NA"))))
4. Change the default formula for A-Symbol to BTC/INR (text or string datatype)
5. Run Sub