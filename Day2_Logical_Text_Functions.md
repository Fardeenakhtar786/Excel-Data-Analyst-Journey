# Day 2: Logical Functions + Text Functions

Today's focus: how Excel makes decisions (IF function) and how to manipulate text data.

## 1. IF Function — Excel's "Decision Maker"

**Syntax:**
```
=IF(condition, value_if_true, value_if_false)
```

**Example:** Label clients who completed more than 5 reels as "Good Client", and the rest as "Normal Client".

Add a new column "Status" (Column E):
```
=IF(B2>5, "Good Client", "Normal Client")
```

| Client | Reels | Status |
|---|---|---|
| Rahul | 5 | Normal Client |
| Priya | 8 | Good Client |
| Aman | 3 | Normal Client |
| Sneha | 10 | Good Client |

**Common comparison operators:**
- `>` greater than
- `<` less than
- `>=` greater than or equal to
- `<=` less than or equal to
- `=` equal to
- `<>` not equal to

---

## 2. Nested IF — Multiple Conditions

To create three categories (e.g., Low, Medium, High), place one IF function inside another:

```
=IF(B2>=10, "High", IF(B2>=5, "Medium", "Low"))
```

**Step-by-step logic:**
1. First check: is B2 >= 10? If yes → "High"
2. If not, check: is B2 >= 5? If yes → "Medium"
3. If neither is true → "Low"

| Reels | Result |
|---|---|
| 10 | High |
| 5 | Medium |
| 3 | Low |

---

## 3. AND / OR Functions — Combining Conditions

### AND — All conditions must be true
```
=IF(AND(B2>5, C2=200), "Premium Client", "Regular Client")
```
Returns "Premium Client" only if reels > 5 **AND** rate = 200 — both conditions must hold.

### OR — At least one condition must be true
```
=IF(OR(B2>10, C2>250), "Special Case", "Normal")
```
Returns "Special Case" if reels > 10 **OR** rate > 250 — only one condition needs to be true.

**Memory tip:**
- AND = "both required"
- OR = "either one works"

---

## 4. Text Functions

These functions are useful for cleaning messy data or splitting/combining text.

### CONCATENATE / CONCAT — Joining text
```
=CONCATENATE(A2, " - ", E2)
```
Output: "Rahul - Normal Client"

(In newer versions of Excel, you can also use the `&` symbol: `=A2&" - "&E2`)

### UPPER / LOWER — Changing text case
```
=UPPER(A2)   → RAHUL
=LOWER(A2)   → rahul
```

### LEN — Counting text length
```
=LEN(A2)
```
For "Rahul", the result is 5 (5 characters).

### LEFT / RIGHT — Extracting part of a text string
```
=LEFT(A2, 3)   → "Rah" (first 3 characters from the left)
=RIGHT(A2, 3)  → "hul" (first 3 characters from the right)
```

**Real-world use case:** Suppose a phone number includes a country code (+919876543210), and you only need the last 10 digits:
```
=RIGHT(A2, 10)
```

### TRIM — Removing extra spaces
```
=TRIM(A2)
```
If text contains extra spaces (e.g., "  Rahul  "), TRIM cleans it up. Very useful when working with messy real-world data.

### MID — Extracting text from the middle
```
=MID(A2, 2, 3)
```
Extracts 3 characters starting from the 2nd character of A2.

---

## Practice Task — Day 2

1. Add a "Status" column to your Day 1 table using the IF function (more than 5 reels = "Good Client")
2. Create a "Category" column using Nested IF: High (10+), Medium (5–9), Low (<5)
3. Use AND to identify "Premium" clients (reels > 5 AND rate = 200)
4. Use CONCATENATE to combine Client Name and Status into one column
5. Use UPPER to convert all client names to CAPITAL letters

---

## Quick Recap — Day 1 + Day 2

| Function | Purpose |
|---|---|
| SUM | Calculate a total |
| AVERAGE | Calculate an average |
| COUNT/COUNTA | Count cells |
| MAX/MIN | Highest/lowest value |
| $ (Absolute reference) | Lock a cell reference |
| IF | Check a condition |
| Nested IF | Multiple conditions |
| AND/OR | Combine conditions |
| CONCATENATE/& | Join text |
| UPPER/LOWER | Change text case |
| LEFT/RIGHT/MID | Extract part of text |
| TRIM | Remove extra spaces |
| LEN | Get text length |

---

**Coming up — Day 3: VLOOKUP, XLOOKUP, INDEX-MATCH** (one of the most commonly asked topics in data analyst interviews)
