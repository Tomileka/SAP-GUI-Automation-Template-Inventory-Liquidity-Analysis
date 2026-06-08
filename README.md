# SAP-GUI-Automation-Template-Inventory-Liquidity-Analysis
This is the code to automate SAP transaction J3RFLVMOBVED to get the stock by material from multiple plants. 

### 📊 Project Overview: Inventory Liquidity Analysis

**Goal:** Automate the consolidation of fragmented inventory data across **54 plants** for the entire year of **2025**.

#### 🛠️ Key Challenges & Solutions:
*   **SAP Automation:** Developed **SAP GUI Scripting** to handle bulk data exports, eliminating manual repetitive tasks.
*   **Data Standardization:** Built a custom **VBA liquidity classification engine** to standardize raw data into meaningful business categories.
*   **ETL Pipeline:** Managed a high volume of files (**648 total**) using **Power Query**, ensuring a stable and repeatable consolidation process.

#### 🧰 Tech Stack:
*   **SAP GUI Scripting:** Bulk data extraction automation.
*   **VBA (Excel Macros):** Data processing & Liquidity logic implementation.
*   **Power Query (M):** ETL pipeline and data aggregation.
*   **Excel:** Dynamic dashboard and reporting.

#### 📈 Results:
*   **Efficiency:** Reduced the reporting cycle from **month to just 3 days**.
*   **Visibility:** Provided a unified, real-time view of inventory health across all manufacturing sites.

---

RU: Это скрипт для автоматизации выгрузки из транзакции SAP J3RFLVMOBVED - запасы материалов с документами для нескольких заводов. Особенность отчета по запасам в том, что он делает выгрузку по месяцам с кумулятивным эффектом. 
Если вам нужно сделать выгрузку за весь год по 12 месяцев по 54 завода, как мне, то эта программа поможет вам сэкономить один месяц работы аналитика вместо проставления дат и кодов заводов руками.

**Note: This repository contains an anonymized version of a real-world business project. All sensitive data has been replaced with synthetic datasets to comply with NDA requirements, while preserving the original logic and analytical methodology.**

<img width="242" height="622" alt="Scheme" src="https://github.com/user-attachments/assets/3411a1f8-0e05-4923-a03f-5c5de82bafab" />

**It gives cummulative results strating from the 1st of January. Example: if you need the report for January, February, and March, it will upload as follow: 1st file (01 Jan 2025 - 31 Jan 2025), 2nd file (01 Jan 2025 - 28 Feb 2025), 3rd file (01 Jan 2025 - 31 March 2025).**

Например, если нужен отчет за три месяца 2025 за январь, февраль и март, то он выгрузит согласно этим датам: 1 файл (01.01.25 - 31.12.25), 2ой файл (01.01.25 - 28.02.25) и 3ий файл (01.01.25 - 31.03.25).

**This is the image of the path in SAP where you can launch this script amended with your plants' name and other data.** 
Это путь в SAP через который вы можете запустить скрипт как обычный макрос, предварительно исправив код под ваши данные:
![SAP_Scripting](https://github.com/user-attachments/assets/830db254-0be5-4f00-bc9a-87f02772c458)

The data to amend / Данные в коде для корректировки: 
1) the codes of the plants as it is in your system / коды заводов, как в вашей системе. Example: "0201", "0255"
2) months you needed for report / месяца для отчета. Example: "02", "05"
3) Check the year / Год отчета
4) Path to the folder where you would like to store the results / Путь к месту назначения отчетов.
   
#### ⚙️ Process Flow
<!-- На мобильных устройствах схема может не отображаться, поэтому прячем код под спойлер -->
<details>
  <summary><b>Click to expand the Process Diagram (Best viewed on Desktop)</b></summary>

  ```mermaid
  graph TD
      A[SAP GUI Scripting<br/>Automation of data export] -->|648 CSV/XLS Files| B(Excel Macro - VBA)
      
      subgraph Data_Transformation_Layer
      B --> C{Liquidity Logic<br/>Classification}
      C -->|Processing| D[Anonymized Data<br/>with Inventory Classes]
      end

      D -->|Folder Import| E[Power Query - ETL]
      
      subgraph Aggregation_Layer
      E --> F[Combine 54 Plants]
      F --> G[Consolidate 12 Months]
      end

      G --> H[Final Analytics Report<br/>Interactive Dashboard]

      style A fill:#f9f,stroke:#333,stroke-width:2px
      style E fill:#bbf,stroke:#333,stroke-width:2px
      style H fill:#dfd,stroke:#333,stroke-width:4px
```


