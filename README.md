# DataDog Analysis Tools

Simple guide for generating different types of reports from your data.

## 📋 Use Cases

### 1. 🎯 Generate Complete Comprehensive Report
**What it does:** Creates a full analysis with charts, error categorization, and detailed metrics for all services.

**How to run:**
```bash
cd /Users/shtlpmac027/Documents/DataDog
source venv/bin/activate
python run_individual_analysis.py
```

**Output:** 
- Individual analysis files in `individual_analysis/` folder
- Comprehensive Excel report in `combined_reports/` folder
- Charts and visualizations for each service

---

### 2. 📊 Generate Comparative Daily Analysis (TXT files)
**What it does:** Creates daily comparison reports showing metrics changes between two dates.

**How to run:**
```bash
cd /Users/shtlpmac027/Documents/DataDog
source venv/bin/activate
python scripts/simple_individual_analyzer.py --compare 24/09,25/09
```

**Output:** 
- Daily analysis TXT files in `individual_analysis/[service]/daily_analysis_[date1]_vs_[date2].txt`
- Each file contains: Latency, Throughput, LLM Cost, Reliability, User Activity metrics
- Shows changes and status (STABLE, IMPROVING, DEGRADING, etc.)

---

### 3. 📈 Aggregate Daily Analysis into Excel Report
**What it does:** Combines all daily analysis TXT files into a formatted Excel with color coding.

**How to run:**
```bash
cd /Users/shtlpmac027/Documents/DataDog
source venv/bin/activate
c
```

**Output:** 
- `formatted_daily_analysis.xlsx` in root directory
- Color-coded status indicators (Green/Red/Yellow)
- Bold service names
- Professional table format
- Multiple sheets for different date comparisons

---

## 🗂️ File Structure

```
DataDog/
├── source_data/           # Raw data files
├── individual_analysis/   # Generated analysis files
│   ├── preparesubmission/
│   ├── qna/
│   ├── relevantdoc/
│   ├── search/
│   └── summary/
├── combined_reports/     # Comprehensive reports
├── scripts/             # Analysis scripts
└── formatted_daily_analysis.xlsx  # Final Excel report
```

## 🚀 Step-by-Step Instructions

### **Step 1: Activate Virtual Environment**
```bash
cd /Users/shtlpmac027/Documents/DataDog
source venv/bin/activate
```

### **Step 2: Choose Your Analysis Type**

#### **Option A: Complete Comprehensive Report**
```bash
# This generates everything - individual analysis + comprehensive report
python run_individual_analysis.py
```
**What happens:**
- Analyzes all files in `source_data/` folder
- Creates individual analysis for each service
- Generates comprehensive Excel report
- Creates charts and visualizations

#### **Option B: Daily Comparative Analysis (TXT files only)**
```bash
# This generates only the daily analysis TXT files
python scripts/simple_individual_analyzer.py --compare 24/09,25/09
```
**What happens:**
- Creates `daily_analysis_[date1]_vs_[date2].txt` files
- Shows metrics changes between dates
- Status indicators (STABLE, IMPROVING, DEGRADING)

#### **Option C: Aggregate Daily Analysis into Excel**
```bash
# This combines existing TXT files into Excel
python scripts/format_daily_analysis.py
```
**What happens:**
- Reads existing daily analysis TXT files
- Creates `formatted_daily_analysis.xlsx`
- Color-coded status indicators
- Professional table format

### **Step 3: Check Your Outputs**

#### **After Complete Analysis:**
- **Individual files:** `individual_analysis/[service]/daily_analysis_*.txt`
- **Charts:** `individual_analysis/[service]/*.png`
- **Comprehensive report:** `combined_reports/analysis_report_*.xlsx`

#### **After Excel Aggregation:**
- **Excel file:** `formatted_daily_analysis.xlsx` (in root directory)

## 📝 Important Notes

- **Always activate virtual environment first:** `source venv/bin/activate`
- **Data files:** Place your Excel files in `source_data/` folder
- **Outputs:** Check `individual_analysis/` and `combined_reports/` folders
- **Excel features:** Color coding (Green=Good, Red=Bad, Yellow=Stable)
- **File formats:** Scripts work with `.xlsx` files in `source_data/` folder
