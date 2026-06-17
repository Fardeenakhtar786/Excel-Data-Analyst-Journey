# Day 1: Excel Basics + Core Formulas

## 1. Understanding the Excel Interface

When you open Excel, you'll see the following elements:

- **Ribbon** — the menu bar at the top (Home, Insert, Formulas, Data, etc.)
- **Cell** — a single box where you enter data (e.g., A1, B2)
- **Row** — horizontal line, identified by numbers (1, 2, 3...)
- **Column** — vertical line, identified by letters (A, B, C...)
- **Range** — a group of multiple cells (e.g., A1:A10)
- **Sheet** — a single page of data, shown as tabs at the bottom (Sheet1, Sheet2)

**Important rule:** Every formula in Excel starts with an `=` sign. Without it, Excel treats your input as plain text.

---

## 2. Data Types

Excel works with three main types of data:

| Type | Example | Notes |
|---|---|---|
| Text | "Rahul", "Mumbai" | Left-aligned by default |
| Number | 200, 5.5 | Right-aligned by default |
| Date | 17/06/2026 | Internally stored as a number |

---

## 3. Practice Dataset

This is the dataset used throughout Day 1 and Day 2 — based on a freelance video editing business.

| | A | B | C |
|---|---|---|---|
| 1 | Client Name | Reels Done | Rate per Reel |
| 2 | Rahul | 5 | 200 |
| 3 | Priya | 8 | 200 |
| 4 | Aman | 3 | 200 |
| 5 | Sneha | 10 | 200 |
| 6 | Vikas | 6 | 200 |

---

## 4. Basic Formulas

### SUM — Calculate a total
```
=SUM(B2:B6)
```
Adds up all reels completed across clients.

### AVERAGE — Calculate an average
```
=AVERAGE(B2:B6)
```
Average reels per client.

### COUNT — Counts only numeric cells
```
=COUNT(B2:B6)
```

### COUNTA — Counts cells containing any data (text or numbers)
```
=COUNTA(A2:A6)
```

### MAX / MIN — Find the highest / lowest value
```
=MAX(B2:B6)
=MIN(B2:B6)
```

### Multiplication (Calculating income)
Add a new column "Income" (Column D):
```
=B2*C2
```
This calculates Rahul's income: 5 × 200 = 1000

Drag this formula from D2 down to D6 (hover over the cell's bottom-right corner until a `+` cursor appears, then drag) — Excel will automatically adjust the formula for each row.

---

## 5. Cell Referencing — Relative vs Absolute

This is one of the most important concepts for beginners, and a common source of confusion.

**Relative Reference (default):** `=B2*C2`
When dragged, the row number changes automatically (B3*C3, B4*C4, etc.)

**Absolute Reference:** `=B2*$C$1`
The `$` symbol locks a cell reference so it does NOT change when dragged.

**Example scenario:** Suppose you have a fixed rate stored in a separate cell (F1 = 200), and you want every client's income calculated using that same fixed rate:

```
=B2*$F$1
```
When dragged down, B2 will change (B3, B4...) but $F$1 stays locked.

**Memory tip:** Think of `$` as a "lock" — place it before the row/column you want to keep fixed.

---

## 6. Basic Formatting

- **Currency format:** Select numbers → Home tab → Number format dropdown → Currency
- **Bold/Color:** Ctrl+B for bold, use the Font Color dropdown for color
- **Borders:** Home tab → Borders icon
- **Auto-fit column width:** Double-click the column border

---

## Practice Task — Day 1

1. Build the table above using real or sample client data
2. Create an Income column using the multiplication formula
3. Calculate total income using SUM
4. Calculate average reels per client
5. Find the client with the highest reels using MAX
6. Format the Income column as Currency
