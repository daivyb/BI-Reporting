# DIME Physical & Financial Project Tracking

This repository contains the data architecture and ETL scripts used to build a comprehensive Business Intelligence dashboard for tracking the physical and financial performance of various projects (DIME). 

This project solves the challenge of integrating data from two disparate systems with different granularities into a unified Kimball Star Schema, enabling executive decision-making through Power BI.

## 📌 Project Overview

**The Challenge:**
- **Financial Data (EMMA):** Handled budgets and expenditures. The granularity was mixed across administrative levels (Executing Units) and technical levels (Project > Component > Subcomponent).
- **Physical Progress Data (Sciforma):** Handled the actual construction/implementation progress. The granularity was strictly at the lowest Work Breakdown Structure (WBS) level: the **Task** (Tarea).

**The Solution:**
To allow cross-filtering and joint analysis without duplicating data or creating false intersections, I implemented a Kimball Star Schema architecture:
1. **ETL Pipeline (Python):** Processed the raw Excel/XLSB dumps from both systems, normalized the WBS structures (from standard dot-notation `1.1.1.1` to internal codes `144-01010101`), and flattened the hierarchy.
2. **Star Schema:** Created a central `Dim_Tecnica` dimension that serves as the backbone connecting the Financial Fact Table (`Fact_Ejecucion`) and the Physical Fact Table (`Fact_Avance_Fisico`).
3. **Advanced DAX:** Since the physical progress is only recorded at the Task level, complex DAX `AVERAGEX` iterators were developed to dynamically roll up the progress across the hierarchical levels in Power BI matrix visuals.

## Repository Structure

- `etl_star_schema.py`: The core Python script using `pandas` and `pyxlsb` to extract, transform, and load the data into the Star Schema Excel output.
- `DAX_Measures.md`: Documentation of the complex DAX formulas used in the Power BI model to calculate hierarchical averages, cumulative S-curves, and financial/physical gaps.
- `data_samples/`: Synthetic/anonymized samples showing the schema of the resulting dimensions and fact tables.

*(Note: Raw corporate data, full outputs, and the frontend React POC have been excluded from this repository for confidentiality).*

## Tech Stack
- **Data Engineering:** Python, Pandas, OpenPyXL, PyXLSB.
- **Data Modeling:** Kimball Methodology (Star Schema), Snowflake concepts.
- **Business Intelligence:** Power BI, Advanced DAX (Context Transition, Nested Iterators, Auto-Exist optimization).

## Key Highlights
- **S-Curve Generation:** Solved the DAX "Sum over time, Average over hierarchy" problem to accurately plot planned goals vs. actual progress.
- **Handling Sparse Data:** Implemented DAX protections against Cartesian products that caused "phantom" components to appear at 100% gap when using multi-table hierarchies.
