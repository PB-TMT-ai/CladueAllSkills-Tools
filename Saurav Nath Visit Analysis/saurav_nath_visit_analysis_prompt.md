# Visit Data Quality Analysis: Saurav Nath Performance Evaluation

## Context & Objective
You are analyzing field visit data for Saurav Nath, a Sales Manager at JSW ONE TMT covering the Jharkhand region (primarily East Singhbum and Saraikela Kharsawan districts). The analysis period spans November 2025 to February 2026. Your goal is to conduct a comprehensive quality and performance analysis to evaluate visit productivity, territory coverage effectiveness, and identify data quality issues.

## Data Structure
The dataset contains the following key fields:
- **Temporal**: Check-in Date/Time, Check-out Date/Time
- **Temporal Format**: DD/MM/YYYY, HH:MM AM/PM
- **Customer**: Account Name, Account SF Id, Customer Type (Retailer PB, Influencer), Account Record Type
- **Location**: Auto State, Auto District, Auto Taluka, Pin Code, Coordinates (Lat/Long)
- **Visit Details**: Meeting Outcome, Check-in Comments, Check-out Comments, Created By
- **Context**: Influencer flag, Monthly Potential (MT), Owner Role

## Analysis Requirements

### 1. VISIT PRODUCTIVITY ANALYSIS
**Target Benchmark**: 6 visits per day

**Required Metrics:**
- Daily visit count (actual vs target of 6)
- Average visits per working day by month
- Visit distribution by time of day (morning/afternoon/evening slots)
- Average visit duration (check-in to check-out)
- Productive hours per day (total time in field)
- Visit completion rate (visits with both check-in and check-out)

**Excel Output - Sheet: "Productivity_Monthly"**
```
Columns:
- Month (Nov '25, Dec '25, Jan '26, Feb '26)
- Total Visits (COUNT formula)
- Working Days (COUNTIF unique dates)
- Avg Visits/Day (=Total Visits/Working Days)
- Days Meeting Target (COUNTIF >=6)
- % Days on Target (=Days Meeting Target/Working Days)
- Avg Duration (min) (AVERAGE of duration calculations)
- Total Field Hours (SUM of all durations)
- Variance from Target (=Avg Visits/Day - 6)
```

### 2. TERRITORY COVERAGE ANALYSIS

**Required Metrics:**
- Geographic spread: Unique districts, talukas, pin codes visited
- Customer segmentation: Distribution across Retailer PB vs Influencer vs Retailer types
- Account penetration: Unique customers visited vs total visits (repeat rate)
- Coverage concentration: % of visits in top 3 pin codes vs balanced distribution
- New customer acquisition: First-time visits vs repeat visits

**Excel Output - Sheet: "Territory_Coverage"**
```
Columns:
- Month
- Unique Customers (COUNTIF unique Account Names)
- New Customers (First appearance in dataset)
- Repeat Customers (=Unique Customers - New Customers)
- Total Visits (COUNT)
- Repeat Visit Rate (=Repeat Customers/Unique Customers)
- Unique Pin Codes (COUNTIF unique)
- Unique Districts (COUNTIF unique)
- Retailer PB Count (COUNTIF)
- Influencer Count (COUNTIF)
- Retailer % (=Retailer PB Count/Total Visits)
- Influencer % (=Influencer Count/Total Visits)
- Top 3 Pin Code Visits (SUM of top 3)
- Concentration Index (=Top 3 Pin Code Visits/Total Visits)
```

### 3. MONTHLY COMPARISON ANALYSIS

**Track month-over-month trends for:**
- Volume metrics (visits, customers, coverage)
- Productivity metrics (visits/day, duration, field hours)
- Quality metrics (documentation completeness, outcome distribution)
- Territory expansion (new areas, new customer types)

**Excel Output - Sheet: "Monthly_Trends"**
```
Rows: Key metrics
Columns: Nov '25 | Dec '25 | Jan '26 | Feb '26 | Trend | % Change (Nov-Feb)

Metrics to include:
- Total Visits
- Avg Visits/Day
- Unique Customers
- Unique Pin Codes
- Avg Duration (min)
- Total Field Hours
- Documentation Score
- % Short Visits (<5 min)
- % Complete Documentation

Formulas:
- Trend: Use SLOPE or visual indicator (↑↓→)
- % Change: =(Feb Value - Nov Value)/Nov Value * 100
```

### 4. QUALITY ISSUE DETECTION

**A. Multiple Visits to Same Location/Customer**
- Identify customers visited 3+ times in the analysis period
- Flag visits to same coordinates (within 50m radius) on same day
- Calculate repeat visit efficiency: Does visit frequency correlate with outcomes?

**Excel Output - Sheet: "Repeat_Visits_Analysis"**
```
Columns:
- Customer Name
- Account SF Id
- Total Visits (COUNTIF)
- Visit Dates (concatenated list)
- Outcomes (concatenated list)
- Avg Duration (AVERAGEIF)
- Visit Frequency Flag (IF >2: "High", >5: "Excessive", else "Normal")
- Same Day Repeat (check coordinate proximity)
- Efficiency Score (based on outcomes)
```

**B. Short Duration Visits**
- Identify visits <5 minutes duration (potential check-in/out errors)
- Categorize duration bands: <5 min, 5-15 min, 15-30 min, 30+ min
- Correlate visit duration with outcome quality

**Excel Output - Sheet: "Short_Duration_Flags"**
```
Columns:
- Date
- Customer Name
- Check-in Time
- Check-out Time
- Duration (min) (calculated: =(Check-out - Check-in)*24*60)
- Duration Band (IF formula: <5, 5-15, 15-30, 30+)
- Meeting Outcome
- Comments Word Count (LEN formula approximation)
- Red Flag (IF Duration <5: "YES", else "NO")
- Quality Score (0-100 based on duration + comments)
```

**C. Documentation Quality**
- Comment completeness: Both check-in and check-out comments present
- Comment depth: Word count, presence of specifics (rates, competitor names, follow-up actions)
- Outcome clarity: Clear next steps documented

**Excel Output - Sheet: "Documentation_Quality"**
```
Columns:
- Date
- Customer Name
- Check-in Comment Present (IF NOT BLANK)
- Check-out Comment Present (IF NOT BLANK)
- Combined Word Count (LEN approximation)
- Contains Rate Info (SEARCH for "rate", "@", "₹")
- Contains Competitor (SEARCH for brand names)
- Contains Follow-up (SEARCH for "will", "inquiry", "connect")
- Completeness Score (0-100)
- Quality Rating (IF >80: "Excellent", >60: "Good", >40: "Fair", else "Poor")
```

### 5. CONSOLIDATED PERFORMANCE DASHBOARD

**Excel Output - Sheet: "Executive_Dashboard"**

Create a single executive summary table combining all critical metrics:

```
Layout:
Row Headers: Performance Dimension
Columns: Target/Benchmark | Nov '25 | Dec '25 | Jan '26 | Feb '26 | Overall | Rating

Performance Dimensions:
1. Visit Productivity
   - Visits per Day (Target: 6)
   - Formula: =AVERAGE(daily visit counts)
   
2. Territory Coverage  
   - Unique Customers per Month (Benchmark: 50+)
   - Formula: =COUNTIF(UNIQUE customer list)
   
3. Geographic Spread
   - Unique Pin Codes (Benchmark: 10+)
   - Formula: =COUNTIF(UNIQUE pin codes)
   
4. Visit Quality
   - Avg Duration (min) (Benchmark: 15+)
   - Formula: =AVERAGE(all durations)
   
5. Documentation Quality
   - Completeness Score (Benchmark: 70+)
   - Formula: =AVERAGE(completeness scores)
   
6. Quality Red Flags
   - % Short Visits (Threshold: <10%)
   - Formula: =COUNTIF(<5 min)/TOTAL VISITS
   - % Excessive Repeats (Threshold: <5%)
   - Formula: =COUNTIF(>5 visits)/UNIQUE CUSTOMERS

Rating Column Formula:
=IF(Actual>=Target*1.1,"🟢 Excellent",IF(Actual>=Target*0.9,"🟡 Good",IF(Actual>=Target*0.7,"🟠 Fair","🔴 Poor")))
```

### 6. DAILY VISIT LOG (Master Data Sheet)

**Excel Output - Sheet: "Daily_Visit_Log"**
```
This sheet contains all processed data with calculated fields:

Columns (in order):
1. Date (parsed from Check-in Date/Time)
2. Day of Week (TEXT formula: =TEXT(Date,"DDD"))
3. Month (TEXT formula: =TEXT(Date,"MMM 'YY"))
4. Check-in Time (parsed)
5. Check-out Time (parsed)
6. Duration (min) (=(Check-out - Check-in)*24*60)
7. Customer Name
8. Account SF Id
9. Pin Code
10. District
11. Taluka
12. Customer Type
13. Meeting Outcome
14. Check-in Comments
15. Check-out Comments
16. Combined Comments (=CONCATENATE)
17. Comments Word Count (approximation using LEN)
18. Latitude
19. Longitude
20. Visit Number for Customer (COUNTIF up to current row)
21. Short Visit Flag (=IF(Duration<5,"YES","NO"))
22. Documentation Complete (=IF(AND(Check-in<>"",Check-out<>""),"YES","NO"))
23. Time Slot (=IF(HOUR(Check-in)<12,"Morning",IF(HOUR(Check-in)<17,"Afternoon","Evening")))
24. Same Day Visits (COUNTIF same date)
```

## Methodology Requirements

### 1. Data Validation & Parsing
```
- Parse dates correctly from DD/MM/YYYY, HH:MM AM/PM format
- Convert to Excel date-time format for calculations
- Calculate durations programmatically: =(Check-out - Check-in)*24*60
- Standardize customer names (TRIM, PROPER functions)
- Validate coordinate pairs (must be numeric, within valid range)
- Handle blank/missing values explicitly (IFERROR wrapping)
```

### 2. Month Boundaries
```
- November 2025: >=01/11/2025 AND <01/12/2025
- December 2025: >=01/12/2025 AND <01/01/2026
- January 2026: >=01/01/2026 AND <01/02/2026
- February 2026: >=01/02/2026

Formula example: 
=COUNTIFS(Date_Column,">=01/11/2025",Date_Column,"<01/12/2025")
```

### 3. Working Days Calculation
```
- Exclude Sundays: Use WEEKDAY function
- Count only days with at least one visit as working days
- Formula: =SUMPRODUCT((WEEKDAY(Date_Range)<>1)*(Date_Range<>""))
- Clearly document any assumptions in a "Notes" sheet
```

### 4. Geographic Analysis
```
- Use Pin Code as primary territory identifier
- Use coordinates to detect exact location clustering
- Same location threshold: Coordinates within 0.001 degrees (≈100m)
- Formula: =SQRT((Lat1-Lat2)^2 + (Long1-Long2)^2) < 0.001
- Group by District for regional patterns using PIVOT tables
```

### 5. Quality Thresholds (Apply consistently)
```
- Short visit: <5 minutes
- Excessive repeat: 5+ visits to same customer in period
- Poor documentation: <20 words in combined comments OR missing check-out comment
- Same location: Coordinates within 0.001 degrees (≈100m)
- Low quality comment: No mention of rates, competitors, or follow-up actions
```

### 6. Formula Best Practices
```
- Use named ranges for better readability (e.g., "Visit_Data", "Customer_List")
- Use COUNTIFS/SUMIFS for multi-criteria calculations
- Use IFERROR to handle missing data gracefully
- Use conditional formatting for visual flags (red/yellow/green)
- Add data validation where appropriate
- Protect formula cells after setup
- Add comments explaining complex formulas
```

## Excel File Specifications

### Workbook Structure
```
Sheet 1: Executive_Dashboard (Summary view)
Sheet 2: Monthly_Trends (Time series comparison)
Sheet 3: Productivity_Monthly (Detailed productivity metrics)
Sheet 4: Territory_Coverage (Coverage analysis)
Sheet 5: Repeat_Visits_Analysis (Quality check: repeats)
Sheet 6: Short_Duration_Flags (Quality check: durations)
Sheet 7: Documentation_Quality (Quality check: comments)
Sheet 8: Daily_Visit_Log (Master data with calculations)
Sheet 9: Charts (Visual representations)
Sheet 10: Notes_Methodology (Documentation of calculations)
```

### Formatting Requirements
```
- Header row: Bold, background color (#4472C4), white text
- Freeze panes: Freeze top row and first column
- Column widths: Auto-fit with minimum 12 characters
- Number formats:
  * Percentages: 0.0%
  * Decimals: 0.00
  * Whole numbers: #,##0
  * Dates: DD/MM/YYYY
  * Times: HH:MM AM/PM
- Conditional formatting:
  * Green for above target
  * Yellow for 80-100% of target
  * Red for below 80% of target
- Borders: All cells with thin borders
- Alternating row colors for readability
```

### Chart Requirements (in "Charts" sheet)
```
1. Line Chart: Monthly trend of Visits/Day vs Target
2. Column Chart: Visit distribution by time slot
3. Pie Chart: Customer type distribution
4. Bar Chart: Top 10 most visited customers
5. Scatter Plot: Visit duration vs Documentation quality
6. Heat Map: Visits by day of week and time slot
```

## Expected Deliverables

### Primary Output: Single Excel File
**Filename**: `Saurav_Nath_Visit_Analysis_Nov25_Feb26.xlsx`

### Sheet-by-Sheet Deliverables:

1. **Executive_Dashboard** 
   - Single consolidated table with all KPIs
   - Traffic light ratings (🟢🟡🔴)
   - Month-over-month comparison
   - Clear target benchmarks shown

2. **Monthly_Trends**
   - All metrics in rows, months in columns
   - Trend indicators and % change calculations
   - Conditional formatting applied

3. **Productivity_Monthly**
   - Detailed productivity breakdown by month
   - Working days calculations
   - Target achievement tracking

4. **Territory_Coverage**
   - Geographic and customer segmentation metrics
   - Concentration analysis
   - New vs repeat customer tracking

5. **Repeat_Visits_Analysis**
   - Customers with 3+ visits flagged
   - Visit frequency patterns
   - Efficiency scoring

6. **Short_Duration_Flags**
   - All visits <5 minutes listed
   - Duration bands categorized
   - Quality correlation shown

7. **Documentation_Quality**
   - Comment completeness scoring
   - Content quality analysis
   - Rating distribution

8. **Daily_Visit_Log**
   - Complete master data
   - All calculated fields
   - Ready for filtering/pivot tables

9. **Charts**
   - 6 key visualizations
   - Professional formatting
   - Clear labels and legends

10. **Notes_Methodology**
    - Calculation explanations
    - Assumptions documented
    - Data quality notes
    - Limitation disclosures

### Secondary Output: Summary Document
Create a brief 1-page summary (can be in a text box on Dashboard sheet):

```
EXECUTIVE SUMMARY
Analysis Period: November 2025 - February 2026
Sales Manager: Saurav Nath
Target: 6 visits/day

KEY FINDINGS:
1. [Productivity insight with numbers]
2. [Territory coverage insight with numbers]
3. [Quality issue insight with numbers]

TOP 3 RED FLAGS:
1. [Specific issue with count/percentage]
2. [Specific issue with count/percentage]
3. [Specific issue with count/percentage]

RECOMMENDATIONS:
1. [Actionable recommendation]
2. [Actionable recommendation]
3. [Actionable recommendation]
```

## Critical Requirements - ZERO TOLERANCE

✅ **All counts must be programmatic** - Use formulas, never hard-coded numbers
✅ **Verify subcategories sum to totals** - Add validation checks
✅ **Document all methodology** - Notes sheet is mandatory
✅ **Handle missing data explicitly** - Use IFERROR, show N/A when appropriate
✅ **Test formulas with sample data** - Verify calculations before full application
✅ **Named ranges for key data** - Improve formula readability
✅ **Protect formula cells** - Lock after validation to prevent accidental changes
✅ **Add data validation** - Where user input might be needed
✅ **Professional formatting** - Clean, corporate appearance
✅ **Chart data sources documented** - Clear what feeds each chart

## Quality Assurance Checklist

Before delivering the Excel file, verify:

- [ ] All formulas calculate correctly (no #REF!, #VALUE!, #DIV/0! errors)
- [ ] Monthly totals add up to overall totals
- [ ] All sheets are properly named and organized
- [ ] Headers are frozen and formatted
- [ ] Conditional formatting is applied consistently
- [ ] Charts update dynamically with data
- [ ] Notes sheet documents all key calculations
- [ ] File size is reasonable (<5MB)
- [ ] All data sources are clearly indicated
- [ ] Executive Dashboard provides clear summary

## Usage Instructions for Claude Code

1. Load the visit data file (report1770198136609.xls)
2. Apply this prompt for comprehensive analysis
3. Generate the Excel output with all specified sheets and formulas
4. Validate all calculations programmatically
5. Apply formatting and conditional formatting
6. Create charts and visualizations
7. Document methodology in Notes sheet
8. Save as: `Saurav_Nath_Visit_Analysis_Nov25_Feb26.xlsx`

---

**End of Prompt**
