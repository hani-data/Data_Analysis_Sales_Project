# تحلیل داده‌های فروش فروشگاه پوشاک – پروژه اکسل با داشبورد و ماکرو
این پروژه شامل تحلیل داده‌های واقعی فروش یک فروشگاه پوشاک کوچک است که داده‌های آن از ۱۱ محصول مختلف جمع‌آوری شده است.  
تمام مراحل تحلیل، تمیزسازی داده، مصورسازی و پیش‌نیازهای تحلیل‌های بعدی در این فایل‌ها قرار دارند.
### 🧾 توضیحات فایل اکسل
فایل اکسل شامل ۴ شیت اصلی است:  
Data main →  داده‌های خرید شامل تاریخ شمسی و ستون تبدیل‌شده به تاریخ میلادی (با استفاده از ماکرو) و نیز باقی اطلاعات خرید کالاها.  
Sales_data → جزئیات فروش واقعی محصولات در بازه‌های زمانی مختلف.  
Sales_Dashboard → داده‌های آماده برای مصورسازی و گزارش‌های خلاصه.  
Readme → توضیحات فنی مراحل و ساختار فایل‌ها داخل خود اکسل.
### 🧩 ماکرو مورد استفاده
در این پروژه از کد VBA برای تبدیل تاریخ شمسی به میلادی در ستون مربوطه استفاده شده است تا در مراحل بعدی بتوان از داده‌های زمانی در ابزارهایی مانند Power BI، SQL و Python بهره گرفت.

### 🖼 تصاویر داشبورد
سه تصویر از داشبورد اکسل در این مخزن آپلود شده است که هر کدام بخش متفاوتی از تحلیل را نشان می‌دهند:  
تصویر اول – اسلایسر نام محصول و ماه فروش که نشان می‌دهد در هر ماه دقیقاً چه کالاهایی فروش رفته‌اند.

![product_slicer](images/Dashboard1_product_slicer.png)

تصویر دوم – جدول و نمودار خطی سود محصولات. شامل پیوت تیبل با مجموع فروش کل هر محصول، میانگین سود واقعی و میانگین درصد سود هر محصول. این مقادیر با رنگ‌بندی Conditional Formatting به سه دسته‌ی رنگی سبز (پرسود)، زرد (میان‌سود) و قرمز (کم‌سود) تقسیم شده‌اند. نمودار خطی نیز همین دسته‌بندی را نمایش می‌دهد.

![profit_analysis](images/Dashboard2_profit_analysis.png)

تصویر سوم – فروش ماهانه کل اقلام. پیوت تیبل و نمودار میله‌ای که نشان می‌دهد در هر ماه مجموع اقلام فروخته‌شده چقدر بوده و چگونه فروش در برخی ماه‌ها به دلیل تخفیفات یا تقاضای فصلی دچار 
نوسان شده.

![monthly_sales_trend](images/Dashboard3_monthly_sales_trend.png)
### 🎯 هدف پروژه
هدف اصلی این پروژه، آماده‌سازی داده‌های فروش برای استفاده در تحلیل‌های پیشرفته‌تر با ابزارهای:

Python  
برای اتوماسیون تحلیل‌ها،  
SQL  
برای ذخیره و پرس‌وجوی داده‌ها،  
Power BI  
برای داشبورد مدیریتی و گزارش‌های تصویری.
# Clothing Store Sales Data Analysis (Excel Dashboard & Macro)
This Excel-based project contains real sales data from a small clothing store (11 products in total).  
The goal is to clean, visualize, and prepare this dataset for deeper analysis using Power BI, SQL, and Python in future steps.
### 🧾 Excel File Overview
The workbook includes four main sheets:  
Data main → Purchase data with both Persian and converted Gregorian dates (via macro) and the rest of the information about purchasing goods.  
Sales_data → Detailed product sales across different months.  
Sales_Dashboard → Preprocessed and aggregated data ready for visualization.  
Readme → Technical documentation of project steps inside Excel.
### 🧩 Macro Description
A VBA macro is used to convert Persian dates to Gregorian in the Data main sheet.  
This provides compatibility with tools such as Power BI, Python, and SQL for later integration and automation.
### 🖼 Dashboard Screenshots
Three dashboard screenshots have been uploaded to illustrate different analytical views:  
Screenshot 1 – Product & Month Slicer. Displays which products were sold in each month.  
Screenshot 2 – Pivot Table + Line Chart of Profit Categories.  
Shows total sales, average actual profit, and average profit % for each product.  
Conditional formatting highlights products as profitable (green), moderate (yellow), and low-profit (red) categories.  
Screenshot 3 – Monthly Sales Volume Overview. Bar chart showing total sold units by month — seasonal fluctuations caused by discounts or demand changes are clearly visible.
### 🎯 Project Goal
To build a clean and dynamic sales data model in Excel, serving as a foundation for next-stage analysis in:  
Python (automation)  
SQL (data storage & queries)  
Power BI (visual reporting)
### ✍️ Author
Developed by [@HaniData](https://github.com/HaniData)  
_Last updated: February 2026_
