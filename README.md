# Late Product Deliveries Analysis
---

![Python](https://img.shields.io/badge/Programming-Python-blue?logo=python&logoColor=white&style=flat-square) <br>
![Microsoft SQL Server](https://img.shields.io/badge/Database-Microsoft_SQL_Server-darkblue?logo=microsoft-sql-server&logoColor=white&style=plastic) <br>
![Power BI](https://img.shields.io/badge/Analytics-Power_BI-yellow?logo=powerbi&logoColor=white&style=flat-square) <br>
![VBScript](https://img.shields.io/badge/Scripting-Visual_Basic-blue?logo=visual-studio&logoColor=white&style=flat-square)


<hr style="border: 2px solid gray;">

### Project Objective
---
This project focuses on creating a robust data pipeline that operates daily, collecting information from both SQL databases and local files. The primary objective is to generate a Power BI report that scrutinizes transporter delivery data, specifically identifying instances of delayed product deliveries to different customers.

The implementation of this comprehensive project has resulted in the identification of non-compliance issues, enhancing transporter performance, and effectively mitigating customer loss. As a direct consequence, the department has experienced a notable increase in profits.


### Project Structure
---

![Late-Product-Deliveries-Analysis-architecture](https://github.com/CarolMmai/Late-Product-Deliveries-Analysis/blob/main/Late_Product_Deliveries_Analysis_Architecture.JPG)



### Dataset Information
---

This dataset is composed of synthetic data designed to simulate an actual historical fact table extracted from SQL, capturing information sourced from SAP. It encompasses essential details for a comprehensive late product delivery report, including dates and times associated with various measurements in the product delivery process.

The dimension tables, structured in a star schema format alongside the fact table, include a Power BI hard-coded date table, customer details with unique identifiers, and information pertaining to product source locations and product classifications.


### Project Approach
---

To achieve the project objective, the following ETL and Power BI report development steps were followed....

1. **Data Extraction:**
Used Visual Basic Scripting and for loops to extract historical data from SAP efficiently, reducing report sizes and minimizing the risk of SAP session timeouts.

2. **Data Staging(temporary and permanent):**
The generated reports were stored temporarily, where a Python script combined them into a single dataframe and wrote the consolidated data to a SQL Server database table using the "replace" mode. Additionally, the script efficiently manages local machine resources by deleting SAP-generated files after reading them into the dataframe. It also logs process outputs to detect and handle failures effectively. The Python script executes a SQL stored procedure, planned for development in a later stage of the process

3. **Data Modeling:**
The consolidated data is incorporated into the SQL server table. Dimension tables, manually created from flat files, are imported and structured in a star schema format.

4. **Data Transformation, Stored Procedures and Views:**
A stored procedure was created to copy SQL data from the original database table to another database, generating a duplicate table for transformations while preserving the integrity of the raw files. Subsequently, desired views were crafted from the transformed table. 

5. **Power BI report development:**
Utilized Power BI to connect with SQL views, constructing a report tailored to the user's specifications. The development encompassed additional transformations in Power Query, along with the creation of measures and calculated columns through DAX.

(All the code used is available in the uploaded files section)


### The Resultant Power BI report
---

<p align="center">
  <img src="https://github.com/CarolMmai/Late-Product-Deliveries-Analysis/blob/main/Late%20Product%20Delivery%20Report.gif" width="800" height="450" alt="Late Product Deliveries Report">
</p>


### Value added to the business
---
1. Reduction in late deliveries due to knowledge of being monitored
2. Gradual increase in profits
3. Less contractual cancellations
4. Increase in process efficiencies