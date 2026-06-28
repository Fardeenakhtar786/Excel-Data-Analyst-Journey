# Day 4: Data Cleaning in Excel

## Why Data Cleaning Matters

In real-world data analyst jobs, raw data is almost never clean. It comes with:
- Extra spaces in text
- Duplicate rows
- Inconsistent formatting (e.g., "mumbai", "Mumbai", "MUMBAI")
- Missing values
- Numbers stored as text

A data analyst spends **60–70% of their time cleaning data** before any analysis begins. Mastering data cleaning in Excel is one of the most valuable skills you can have.

---

## Practice Dataset

Use this messy dataset throughout Day 4. It intentionally has errors to clean.

| | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | Client ID | Client Name | City | Reels Done | Income |
| 2 | 101 | Rahul | mumbai | 5 | 1000 |
| 3 | 102 | PRIYA | Delhi | 8 | 1600 |
| 4 | 103 | Aman | delhi | 3 | 600 |
| 5 | 101 | Rahul | Mumbai | 5 | 1000 |
| 6 | 104 |   Sneha   | Pune | 10 | 2000 |
| 7 | 105 | Vikas | | 6 | 1200 |
| 8 | 106 | sneha | pune | 10 | 2000 |

**Problems in this dataset:**
- Row 5 is a duplicate of Row 2 (same Client ID 101)
- City names are inconsistent (mumbai, Mumbai, MUMBAI)
- Row 6 has extra spaces around "Sneha"
- Row 7 has a missing City value
- Row 8 is a near-duplicate with inconsistent casing
- Name casing is inconsistent (rahul, PRIYA, sneha)

---

## 1. Sorting Data

Sorting organizes your data in ascending or descending order. It is the first step before any analysis.

### How to Sort (Single Column):
1. Click anywhere inside your data
2. Go to **Data tab → Sort & Filter → Sort A to Z** (ascending) or **Sort Z to A** (descending)

### How to Sort (Multiple Columns):
1. Go to **Data tab → Sort**
2. Click **Add Level**
3. Sort first by City (A to Z), then by Client Name (A to Z)

This is useful when you want to group data by city and then alphabetically sort names within each city.

### Formula-based Sort (Excel 365):
```
=SORT(A2:E8, 4, -1)
```
- `A2:E8` → range to sort
- `4` → sort by column 4 (Reels Done)
- `-1` → descending order (use 1 for ascending)

---

## 2. Filtering Data

Filtering lets you temporarily show only the rows that meet a specific condition, without deleting any data.

### How to Apply a Filter:
1. Click anywhere inside your data
2. Go to **Data tab → Filter** (or press **Ctrl + Shift + L**)
3. Dropdown arrows will appear on each column header
4. Click any dropdown → choose your filter condition

### Example Use Cases:
- Show only clients from Delhi
- Show only clients with Reels Done > 5
- Show only rows where City is blank (to find missing values)

### Advanced Filter — Filter by Multiple Conditions:
1. Go to **Data tab → Advanced**
2. Set the List Range (your data) and Criteria Range (your conditions)

### Formula-based Filter (Excel 365):
```
=FILTER(A2:E8, D2:D8>5, "No results")
```
- Returns only rows where Reels Done > 5
- "No results" is shown if nothing matches

---

## 3. Removing Duplicates

Duplicate rows are one of the most common data quality issues.

### How to Remove Duplicates:
1. Click anywhere inside your data
2. Go to **Data tab → Remove Duplicates**
3. Select which columns to check for duplicates
4. Click OK — Excel will tell you how many duplicates were removed

**In our dataset:** Rows 2 and 5 are duplicates (same Client ID 101, same name, same city). Removing duplicates will delete Row 5.

### Identifying Duplicates Without Deleting (using COUNTIF):

Add a helper column "Duplicate Check" (Column F):
```
=COUNTIF($A$2:$A$8, A2)
```
- If result is **1** → unique row
- If result is **2 or more** → duplicate exists

This lets you review duplicates before deleting them, which is safer in professional work.

---

## 4. Handling Missing Values

Missing data (blank cells) can cause errors in formulas and misleading analysis results.

### Step 1 — Find Missing Values:
1. Select your data range
2. Press **Ctrl + G** → Click **Special** → Select **Blanks** → Click OK
3. All blank cells will be highlighted

### Step 2 — Options for handling blanks:

**Option A — Fill with a default value:**
After selecting blanks (using Ctrl + G method above):
- Type "Unknown" or "0" and press **Ctrl + Enter**
- This fills all selected blank cells at once

**Option B — Use IF to flag missing values:**
```
=IF(C2="", "Missing", C2)
```
This creates a new column that shows "Missing" wherever City is blank, keeping original data intact.

**Option C — Use IFERROR to handle formula errors:**
```
=IFERROR(VLOOKUP(A2, D2:E8, 2, 0), "Not Found")
```
If VLOOKUP fails (because the value is missing), it shows "Not Found" instead of an error.

---

## 5. Text Cleaning Functions

These functions fix inconsistent text formatting — one of the most common real-world cleaning tasks.

### TRIM — Remove Extra Spaces
```
=TRIM(B6)
```
Converts `"   Sneha   "` → `"Sneha"`

Always apply TRIM when importing data from external sources like CSV files or databases.

### PROPER — Fix Name Casing
```
=PROPER(B2)
```
Converts `"PRIYA"` → `"Priya"`, `"sneha"` → `"Sneha"`, `"rahul"` → `"Rahul"`

### UPPER / LOWER — Force Case
```
=UPPER(C2)    → "MUMBAI"
=LOWER(C2)    → "mumbai"
```

Use LOWER when you want to standardize city names before comparing them.

### SUBSTITUTE — Replace Specific Text
```
=SUBSTITUTE(C2, "delhi", "Delhi")
```
Replaces all occurrences of "delhi" with "Delhi" in cell C2.

### CLEAN — Remove Non-Printable Characters
```
=CLEAN(B2)
```
Removes hidden characters that sometimes appear when data is imported from external systems or copied from websites.

### Combining TRIM + PROPER + CLEAN (Full Clean):
```
=PROPER(TRIM(CLEAN(B2)))
```
This single formula removes hidden characters, trims extra spaces, and fixes casing all at once. Use this on any column that contains imported text data.

---

## 6. Conditional Formatting

Conditional formatting automatically highlights cells based on rules you define. It is used for **visual data quality checks** — spotting outliers, duplicates, and missing values at a glance.

### How to Apply:
1. Select the column or range you want to format
2. Go to **Home tab → Conditional Formatting**
3. Choose a rule type

### Common Use Cases:

**Highlight duplicate values:**
Home → Conditional Formatting → Highlight Cell Rules → Duplicate Values
(Choose a color — Excel will highlight all duplicates in red/yellow)

**Highlight cells greater than a value:**
Home → Conditional Formatting → Highlight Cell Rules → Greater Than → enter 7
(Highlights all clients with more than 7 reels done)

**Highlight blank cells:**
Home → Conditional Formatting → New Rule → Format only cells that contain → Blanks
(Highlights all empty cells in your selected range)

**Color Scale (for numeric columns):**
Home → Conditional Formatting → Color Scales → choose a gradient
(Green = high value, Red = low value — useful for Reels Done or Income columns)

**Data Bars (mini bar charts inside cells):**
Home → Conditional Formatting → Data Bars
(Shows a horizontal bar proportional to the value — instant visual comparison)

---

## 7. Text to Columns

This feature splits a single column into multiple columns — very useful when importing CSV data.

### Example:
If a cell contains `"Rahul - Mumbai"` and you want to split it into two columns:
1. Select the column
2. Go to **Data tab → Text to Columns**
3. Choose **Delimited** → Click Next
4. Check **Other** and type ` - ` as the delimiter
5. Click Finish

**Result:**
- Column 1: `Rahul`
- Column 2: `Mumbai`

---

## 8. Find and Replace

A quick way to fix inconsistencies across the entire dataset.

**Shortcut:** `Ctrl + H`

### Example:
- Find: `delhi`
- Replace with: `Delhi`
- Click **Replace All**

This instantly standardizes all inconsistent city names without needing a formula.

---

## 9. Data Validation

Data validation prevents incorrect data from being entered in the first place — this is used when building data entry sheets.

### How to Apply:
1. Select the cells where you want to restrict input
2. Go to **Data tab → Data Validation**
3. Choose a rule

### Common Rules:

**Allow only whole numbers between 1 and 50:**
- Validation criteria: Whole Number → Between → 1 and 50

**Allow only values from a dropdown list:**
- Validation criteria: List → Source: `Mumbai, Delhi, Pune, Chennai`
- This creates a dropdown in each cell, preventing typos in City names

**Show a custom error message:**
- Go to Error Alert tab → Title: "Invalid Input" → Message: "Please select a valid city from the list"

---

## Complete Data Cleaning Workflow (Step-by-Step)

When you receive a raw dataset, follow this sequence:

```
Step 1 → Inspect the data visually (scroll through, spot obvious issues)
Step 2 → TRIM + PROPER + CLEAN on all text columns
Step 3 → Remove duplicates (Data → Remove Duplicates)
Step 4 → Find and handle blank cells (Ctrl + G → Special → Blanks)
Step 5 → Use Find & Replace for inconsistent values
Step 6 → Apply Conditional Formatting to verify cleaning results
Step 7 → Validate final dataset (no blanks, no duplicates, consistent casing)
```

---

## Practice Task — Day 4

Use the messy dataset provided at the top of this file:

1. Apply PROPER + TRIM to fix the Client Name column
2. Apply LOWER to standardize City names, then use Find & Replace to capitalize correctly
3. Use COUNTIF to identify duplicate Client IDs
4. Remove duplicate rows using Data → Remove Duplicates
5. Find all blank cells using Ctrl + G → Special → Blanks and fill with "Unknown"
6. Apply Conditional Formatting to highlight clients with Reels Done > 7
7. Apply a Color Scale to the Income column

---

## Key Functions and Features — Day 4 Recap

| Tool / Function | Purpose |
|---|---|
| Sort | Organize data in order |
| Filter | Show only matching rows |
| Remove Duplicates | Delete duplicate rows |
| COUNTIF | Identify duplicates without deleting |
| Ctrl + G → Blanks | Find all empty cells |
| TRIM | Remove extra spaces |
| PROPER | Fix name casing |
| UPPER / LOWER | Force uppercase or lowercase |
| SUBSTITUTE | Replace specific text |
| CLEAN | Remove hidden characters |
| IFERROR | Handle formula errors gracefully |
| Conditional Formatting | Visually highlight patterns |
| Text to Columns | Split one column into multiple |
| Find & Replace (Ctrl+H) | Fix inconsistencies across dataset |
| Data Validation | Restrict invalid data entry |

---

**Coming up — Day 5: Pivot Tables** (the most powerful data summarization tool in Excel, used in almost every data analyst job)
