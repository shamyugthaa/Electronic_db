
# 📊 Electronic Database Data Analysis Report

## 📌 Project Overview
This project focuses on analyzing an electronic product database using Microsoft Excel. The dataset contains product details, brand names, prices, quantities, categories, price ranges, and date information.

The main objective is to organize the data, categorize products based on price, and generate useful insights for business decision-making.

## 🎯 Objectives
- Analyze electronic and other product database records.
- Identify products based on their price ranges.
- Compare product prices and quantities.
- Analyze product categories and brands.
- Organize product IDs to extract date-related information.
- Present data in a structured format for reporting.

## 📂 Dataset Description

The dataset contains the following columns:

| Column Name | Description |
|---|---|
| Product ID | Unique identifier for each product |
| Product Name | Name of the product |
| Brand Name | Manufacturer or brand |
| Price ($) | Price of the product in dollars |
| Quantity | Number of units available |
| Category | Product category |
| Price range | Classification based on price |
| Right | Country code extracted from Product ID |
| Left | Day extracted from Product ID |
| Month | Month extracted from Product ID |

## 📋 Dataset

| Product ID | Product Name | Brand Name | Price ($) | Quantity | Category | Price range | Right | Left | Month |
|---|---|---|---:|---:|---|---|---|---:|---|
| 28-JAN-US | Laptop | Dell | 1000 | 30 | Electronics | High price | US | 28 | JAN |
| 15-FEB-US | Sneakers | Nike | 80 | 15 | Fashion | Standard price | US | 15 | FEB |
| 03-MAR-US | Coffee Maker | Keurig | 130 | 40 | Kitchen | Standard price | US | 03 | MAR |
| 11-APR-US | Smartphone | Samsung | 900 | 25 | Electronics | High price | US | 11 | APR |

## 🛠️ Tools Used
- **Microsoft Excel** – Data entry, cleaning, and analysis
- **Excel IF Function** – Product price classification
- **Excel Text Functions** – Extracting date and country information
- **GitHub** – Project documentation and version control

## 🧮 Excel Formula Used

### 1. Product Price Classification

The IF function classifies products based on their price.

```excel
=IF([@[Price ($)]]>=500,"High price","Standard price")
```

**Classification Rules:**
- Price greater than or equal to $500 → High price
- Price below $500 → Standard price

### 2. Extract Month from Product ID

```excel
=MID([@[Product ID]],4,3)
```

### 3. Extract Day from Product ID

```excel
=LEFT([@[Product ID]],2)
```

### 4. Extract Country Code from Product ID

```excel
=RIGHT([@[Product ID]],2)
```

## 📈 Data Analysis and Findings

### 1. Total Products
- Total product records: **4**
- Unique brands: **4**
- Product categories: **3**

### 2. Price Analysis
- Highest-priced product: Laptop – Dell ($1,000)
- Lowest-priced product: Sneakers – Nike ($80)
- Total of listed product prices: $2,110
- Average product price: $527.50

### 3. Quantity Analysis
- Total quantity: **110 units**
- Highest quantity: Coffee Maker – Keurig (40 units)
- Lowest quantity: Sneakers – Nike (15 units)

### 4. Category Analysis
- Electronics: 2 products
- Fashion: 1 product
- Kitchen: 1 product

### 5. Price Range Analysis
- High price: 2 products
- Standard price: 2 products

## 💡 Key Insights
1. Electronics is the most represented category in the dataset.
2. Dell Laptop has the highest listed product price.
3. Sneakers have the lowest listed product price.
4. Coffee Makers have the highest quantity, with 40 units.
5. Half of the products fall into the high-price category, while the remaining half fall into the standard-price category.

## 🚀 Conclusion

This project demonstrates how Microsoft Excel can be used to organize and analyze a product database. IF functions help classify products according to price, while text functions extract useful information from product IDs.

The analysis provides a basic understanding of product pricing, quantities, brands, and categories. These insights can support inventory management and business reporting.

## 👤 Author
**Data Analyst | Excel Data Analysis Project**

## 🏷️ Skills Demonstrated
`Microsoft Excel` `Data Analysis` `Data Cleaning` `IF Function` `Text Functions` `Data Classification` `GitHub` `Markdown`
