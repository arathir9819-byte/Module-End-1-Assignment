# Module-End-1-Assignment
Data Cleaning, Analysis & Visualization with Excel

Data Processing: The dataset consisted of four tables: Customer_Dim ,Product_Dim ,Store_Dim and Sales_Fact .                                                                                                         
Using Power Query:
Removed duplicate data.                                                  
To find the missing values in Unit_price ,Discount,Quantity columns by connecting formulas and median.                                                                            
Corrected Inconsistent text formats,Standardised data types(using Trim(),Clean() and Capitalize Each Word).                                                                  
Removed the column which contains irrelevant data(sub_Category).                                                                                                           
Load the Cleaned Data.

Created a new Worksheet named "Combined Sales_Fact": Retrieved three columns using VLOOKUP on the reference of Product_ID (Product_Name , Category and Product_cost) .     Using INDEX and MATCH added the columns Customer_Name, Customer_Age, Store_Region and Store_type.                                                             
Used the TEXT() function for Format Transaction date (TEXT([@[Order_Date]],"mmm-yyyy")).                                                                                     
To Classify the transactions into High sales and Low sales used IF Function (IF([@[Total_Sales_Amount]]>=1000,"High Sales","Low Sales")) and using Conditional Formatting Highlighted the cells containing High sales.

Created a new Worksheet named "Formula_Analysis and Statistics" : Using Excel Formulas find the Maximum sales, Minimum Sales,Total transactions, Total Sales Amount, Total Quantity sold, Average Sales Amount  and  using COUNTIF and SUMIF find the number of transactions and sales amount of Electronics category.                         
To find Descriptive Statistics of Sales amount column used Analysis Tool Pack.

Create a Worksheet named "Pivot_Final_Sheet" to summarise the transaction data: created Pivot tables to analyse sales by category,region,store type and month.         Developed charts (Column,Bar,Pie and Line) and  Sparklines to represent the trends of monthly wise quantity and sales in product.                                        
Finally Created an Interactive Dashboard using Pivot charts and Slicers.

Key Findings:
* Highest revenue-generating product category : Sports
* Highest Performing Region :North
* Best Performing Store type : Online
* Total Sales Revenue :  $21,74,394.70
* Highest Sales Month : March
  
