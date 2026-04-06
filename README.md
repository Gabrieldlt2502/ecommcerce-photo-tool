# E-commerce Inventory Visual Auditor

## The Business Problem
In high-volume e-commerce, ensuring that the images on the marketplace (e.g., Mercado Libre) match the physical stock in the warehouse is a massive manual task. Discrepancies lead to customer returns and lost revenue.

##  The Solution
I developed this Python tool to automate the **Quality Assurance (QA)** process. It generates a comprehensive Excel report that side-by-side compares:
1. **Marketplace Images:** Fetched in real-time via URL.
2. **Local Warehouse Images:** Indexed automatically from local directories.

##Technical Features
* **Multi-Source Indexing:** Scans local file systems to build a fast lookup index for thousands of images.
* **Dynamic Report Generation:** Uses `openpyxl` to build a complex, formatted Excel "Comparator" sheet with custom row heights, borders, and conditional layouts.
* **Image Processing:** Uses `Pillow (PIL)` to handle diverse image formats (WebP, PNG, JPG) and standardize them for report embedding.
* **Automated Caching:** Implements a temporary download cache to optimize performance and reduce redundant network requests.

## 📈 Impact
* **Accuracy:** Eliminated human error in visually identifying mismatched SKUs.
* **Speed:** Reduced the auditing time for a 1,000-item inventory from days to minutes.
