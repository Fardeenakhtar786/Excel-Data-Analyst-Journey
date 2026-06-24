# Day 3: VLOOKUP, XLOOKUP, and INDEX-MATCH

## Why Learn These Functions?

One of the most common tasks for a data analyst is **looking up data from one table and matching it to another** — for example, finding a client's name using their ID, or retrieving a product price using a product code. These three functions handle exactly that.

---

## Practice Dataset

Build these two tables in Excel before starting:

**Table 1 — Client Data (cells A1:C6):**

| Client ID | Client Name | Reels Done |
|---|---|---|
| 101 | Rahul | 5 |
| 102 | Priya | 8 |
| 103 | Aman | 3 |
| 104 | Sneha | 10 |
| 105 | Vikas | 6 |

**Table 2 — Search Table (cells E1:F2):**

| Search ID | Client Name |
|---|---|
| 103 | ? |

---

## 1. VLOOKUP — Vertical Lookup

**"V" stands for Vertical — it searches top to bottom in a column.**

### Syntax
```
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

| Parameter | Meaning |
|---|---|
| `lookup_value` | What are you searching for? (e.g., Client ID) |
| `table_array` | Where should Excel search? (the full table range) |
| `col_index_num` | Which column's value do you want? (1, 2, 3...) |
| `range_lookup` | Always use **0** for exact match |

### Example
Type this formula in cell F2:
```
=VLOOKUP(E2, A2:C6, 2, 0)
```

**Breaking it down:**
- `E2` → Search for the value 103
- `A2:C6` → Search within this table
- `2` → Return the value from the 2nd column (Client Name)
- `0` → Exact match required

**Result:** Aman ✅

To get Reels Done instead, change `col_index_num` to 3:
```
=VLOOKUP(E2, A2:C6, 3, 0)
```
**Result:** 3 ✅

### Limitations of VLOOKUP

1. **Left to right only** — the lookup column must always be the first column of your table. It cannot search right to left.
2. **Fixed column number** — if you insert a new column in your table, you have to manually update the column number in the formula.

These limitations are why professionals prefer **XLOOKUP** or **INDEX-MATCH** in modern work.

---

## 2. XLOOKUP — The Upgraded VLOOKUP

> Available in Excel 2019, Excel 365, and newer versions.

### Syntax
```
=XLOOKUP(lookup_value, lookup_array, return_array)
```

| Parameter | Meaning |
|---|---|
| `lookup_value` | What are you searching for? |
| `lookup_array` | Which column to search in? |
| `return_array` | Which column to return the result from? |

### Example — Left to Right
```
=XLOOKUP(E2, A2:A6, B2:B6)
```

**Breaking it down:**
- `E2` → Search for 103
- `A2:A6` → Search in the Client ID column
- `B2:B6` → Return the value from the Client Name column

**Result:** Aman ✅

### Example — Right to Left (not possible in VLOOKUP)
Find the Client ID for "Aman":
```
=XLOOKUP("Aman", B2:B6, A2:A6)
```
**Result:** 103 ✅

XLOOKUP can search in any direction, which makes it far more flexible than VLOOKUP.

---

## 3. INDEX-MATCH — The Most Powerful Combination

This is a combination of two separate functions working together:
- **MATCH** → finds the row number of a value
- **INDEX** → retrieves the value from that row

### MATCH Syntax
```
=MATCH(lookup_value, lookup_array, 0)
```
Returns the **position** (row number) of the value within the range.

### INDEX Syntax
```
=INDEX(return_array, row_number)
```
Returns the value at a specific position in a range.

### Combined Formula
```
=INDEX(B2:B6, MATCH(E2, A2:A6, 0))
```

**Step-by-step logic:**
1. `MATCH(E2, A2:A6, 0)` → Where is 103 in A2:A6? → **Position 3**
2. `INDEX(B2:B6, 3)` → What is the 3rd value in B2:B6? → **"Aman"**

**Result:** Aman ✅

### Why Use INDEX-MATCH?

- Works in **any direction** — left, right, up, or down
- Faster than VLOOKUP on large datasets
- Does not break when columns are added or removed
- Works on all versions of Excel

---

## Comparison Table

| Feature | VLOOKUP | XLOOKUP | INDEX-MATCH |
|---|---|---|---|
| Easy to learn | ✅ Yes | ✅ Yes | ❌ Slightly complex |
| Right to Left lookup | ❌ No | ✅ Yes | ✅ Yes |
| Speed on large data | Slow | Fast | Fastest |
| Excel version | All versions | 2019+ only | All versions |
| Industry usage | Older teams | Modern teams | Advanced analysts |

---

## Practice Task — Day 3

1. Build the Client Data table in Excel
2. Use VLOOKUP to find the name of Client ID 104
3. Use VLOOKUP to find the reels done by Client ID 102
4. Use XLOOKUP to find the Client ID of "Vikas" (right to left test)
5. Use INDEX-MATCH to find the name of Client ID 101

---

## Key Takeaways

- Use **VLOOKUP** when you are just starting out and data is simple
- Use **XLOOKUP** when you have a newer version of Excel and need flexibility
- Use **INDEX-MATCH** when working with large or complex datasets — it is the industry standard for advanced analysts

---

**Coming up — Day 4: Data Cleaning** (sorting, filtering, removing duplicates, conditional formatting)
