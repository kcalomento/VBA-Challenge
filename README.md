Multiple Year Stock Market Data Analysis with VBA
-----

**Description**:
-----
This project uses VBA scripting specifically in Excel to automate the analysis of stock market data in each quarter. The script processes a large dataset of stock tickers and outputs key financial metrics, including quarterly changes, percent changes, and total stock volumes. It also identifies the stock with the greatest percent increase, percent decrease, and total volume in each quarter.
Included below are interactive instructions if you wish to indulge in the fun!
* Please note that the data in this workbook is from the Year 2022. 

**How The Metrics Are Calculated**:
----

**Quarterly Change**: The quarterly change for each stock ticker is determined by comparing the opening price at the beginning of the quarter with the closing price at the end of the quarter.

**Percent Change**: The percentage change is computed by dividing the difference between the opening price and closing price by the opening price for each stock, then multiplying by 100.

**Total Stock Volume**: The total stock volume is calculated by summing the stock volumes for each ticker across the entire dataset.

**Greatest Values**: The script identifies the stock with the greatest percentage increase, the greatest percentage decrease, and the highest total stock volume.

**Want to see the magic happen for yourself? You can run it!**
----
* **DISCLAIMER**: While every effort has been made to ensure that the VBA scripts run correctly and efficiently, please use them at your own discretion. The author cannot be held responsible for any issues that may arise while using these scripts. It is recommended to test the code on a backup or non-critical data before applying it to important datasets.
* **Installation**:

  - To run this project, you will need Excel 2010 or later with Macros enabled (directions for Macros settings are below).
     - _Note: Make sure you are on the xlsm file or have converted the xlsx into xlsm._
  - With Excel open, navigate to the ribbon and select **File** -> **Options** -> **Trust Center** -> **Trust Center Settings**
  - Select all the options as the picture shows below:
  ![image](https://github.com/user-attachments/assets/c3c8627a-99de-41fc-afae-64f49f76a2c7)

You're ready!

1) Find and open the **m2_multiple_year_stock_data_diy.xlsm** file in the repository to start off with a blank slate.
2) Navigate to the ribbon. Then, select **Developer** -> **Visual Basic**.
     - By clicking on the folder icon, you will see "**VBAProject (m2_multiple_year_stock_data_diy)**". This is the name of the workbook.
     - Under "Modules", you will see five different modules with subs titled for different calculations.
     - **Module 1**: quarterly_ticker()
     - **Module 2**: quarterly_and_percent_change()
     - **Module 3**: total_stock_volume()
     - **Module 4**: greatest_values()
     - **Module 5**: * _All VBA Code Script Run at Once_ *
  
     - _Note: If you do not see any modules, feel free to locate the scripts in the repository and paste them into new modules._
3) Begin with sheet "**Q1**" and run each script (or the entire script with Module 5). Then, proceed with running it on each sheet: "**Q2**", "**Q3**", and "**Q4**".
     - Click the green "play/run" button.
     - The same VBA scripts will loop across the entire worksheet. How neat!
4) Now, watch all the data unleash in front of your eyes!

Check Your Work / See Finished Workbook:
----
Head back over to the repository to see the correct outputs for this project. Hope you had fun!

Notes
----
If you would like to change colors in the conditionally formatting, feel free to use this link: http://dmcritchie.mvps.org/excel/colors.htm.
