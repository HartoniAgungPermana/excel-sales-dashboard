# 📊 **Microsoft Excel Sales Performance Dashboard – Portfolio Project**

![alt text](<resources/Dashboard Demo.gif>)

This project is an **Excel-based Sales Performance Dashboard** built for a fictional retail company, **FlipMart**.  
It was my **first hands-on data analysis project**, created before I moved on to SQL, Tableau, and Python.  

The purpose of this project was to explore how **Excel Pivot Tables, Pivot Charts, and Slicers** can be applied to analyze business data and build an interactive dashboard. The dashboard was connected to the dataset using an **external data connection**, simulating a real-world BI workflow where dashboards often rely on external sources.  

The dashboard highlights key performance indicators (KPIs):  
- 💰 **Total Profit**  
- 🛒 **Total Sales**  
- 📦 **Quantity Sold**  
- 📑 **Number of Transactions**  

It also provides insights into:  
- 📈 Profit trends over time  
- 🥧 Profit share by product category  
- 🌍 Top countries by profit  
- 👥 Top customers and sub-categories by profit  

This project represents the **starting point of my data analytics journey**. While the analysis focuses mainly on profit, it helped me build a strong foundation in Excel analytics and dashboard design — skills I later expanded using SQL, Tableau, and Python.  

# 📂 **Dataset**

The dataset contains sales records from **2019–2022**.  
You can access it here: [FlipMart_Dataset.xlsx](https://github.com/HartoniAgungPermana/excel-sales-dashboard/blob/main/FlipMart_Dataset.xlsx)

# 📊 **Dashboard Overview**

![alt text](<resources/Dashboard Overview.png>)

# 🛠️ **Excel Features**

- 🔗 Connected the dashboard workbook to the dataset via **external connection**, simulating a **real-world BI setup** while keeping the dataset in a separate file.  
- 📊 Built **Pivot Tables** for dynamic, accurate, and quick aggregations to generate insights.  
- 📉 Used **Pivot Charts** for clear data visualization, customizing chart types for maximum impact.  
- 🎛️ Added **Slicers** to filter data by different dimensions, enabling deeper exploration of measures. 

# 🎨 **Design Choices**

- 🧼 Adopted a **clean layout** with a white background and blue widgets to maintain a **professional look**.  
- 🪪 Displayed **KPI cards** at the top to emphasize key metrics and allow quick scanning.  
- 📈 Used a **line chart** for profit over time, making it easy to identify trends, declines, and growth.  
- 🥧 Implemented a **pie chart** to visualize profit share across categories.  
- 📊 Applied **horizontal bar charts** for ranking analysis, providing clear comparisons across dimensions.  

# 🔍 **Key Insights**

- 📈 Technology consistently generated the highest share of profit (~45%) across the period.  
- 🌍 The United States and China were the largest contributors to profit.  
- 👥 A small group of customers accounted for a significant share of sales, showing reliance on key accounts.  
- 🕒 Despite fluctuations, overall profit showed an upward trend from 2019 to 2022.  

➡️ *Additional insights can be uncovered by applying slicers (e.g., Market, Segment, Order Date).*  

# ⚙️ **How to Use**

1. 📥 **Download the files**  
   - This project uses an **external connection** between the dashboard and the dataset.  
   - Please download **both files** from this repository:  
     - [FlipMart_Dashboard.xlsx](FlipMart_Dashboard.xlsx)
     - [FlipMart_Dataset.xlsx](FlipMart_Dataset.xlsx)

2. 📁 **Keep both files in the same folder**  
   - Place both files in the **same directory** to ensure the external connection works properly.  
   - If Excel prompts you to update the data source path, simply point it to the dataset file in that folder.  

3. 🔓 **Enable content**  
   - Upon opening the dashboard, Excel may display a warning about external connections or macros.  
   - Click **“Enable Content”** to activate the full functionality.  

4. 🖱️ **Explore the dashboard**  
   - Use the **slicers** (Market, Segment, Order Date) to filter the data.  
   - KPIs and charts will update dynamically based on your selections.  

5. 🔧 **Turn off protection for deeper exploration**  
   - To simulate a **real-world BI workflow**, the workbook and all worksheets in `FlipMart_Dashboard.xlsx` are protected, and analysis sheets are hidden.  
   - You can disable protection and **unhide** the analysis sheets if you want to explore the underlying structure in detail.  

---

📌 *Note:* This setup was designed to simulate a **real-world BI process** where dashboards are connected to external data sources.  
If you want to view only the dashboard, you can download just `FlipMart_Dashboard.xlsx`. It will still work, but without the dataset you won’t be able to dig deeper into the Pivot Tables.  
