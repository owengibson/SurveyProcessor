# Survey Input Formatting Cheatsheet

_Quick reference for formatting survey content in Excel_

## 📋 Basic Rules

-   ✅ **Single column** input (default: Column C)
-   ✅ **Sequential order**: Description → `consent` → Questions
-   ✅ **One item per row** (questions, options, etc.)

---

## 🔄 Processing Flow

```
1. Description Content     ← Everything before "consent"
2. consent                 ← Trigger word (switches mode)
3. Consent text           ← One row of consent info
4. Questions & Options    ← Auto-detected question types
```

---

## 📝 Question Types

### 🔤 Text Questions

```
What is your name?        ← Question with "?"
___________________       ← Underscores on next row
                          ← Empty row after
```

**Result**: TextQuestion, gets question ID

### ☑️ Tickboxes (Single Choice)

```
What is your age?         ← Question with "?"
Under 18                  ← Option 1
18-25                     ← Option 2
26-35                     ← Option 3
Over 35                   ← Option 4
                          ← Empty row after options
```

**Result**: Tickboxes, options get numbers only

### ✅ Multiple Choice

```
Please SELECT all social media you use?    ← Contains "select"
Facebook                                   ← Option 1
Instagram                                  ← Option 2
Twitter                                    ← Option 3
                                          ← Empty row after
```

**Result**: Tickboxes, options get individual IDs (a, b, c...)

### 📋 Instructions (Complex)

```
Rate each service 1-5?    ← Question with "?"
**Customer Service**      ← Bold sub-question 1
**Product Quality**       ← Bold sub-question 2
**Delivery Speed**        ← Bold sub-question 3
Excellent                 ← Shared option 1
Good                      ← Shared option 2
Average                   ← Shared option 3
Poor                      ← Shared option 4
                          ← Empty row after
```

**Result**: Instructions + multiple sub-questions

### 🏷️ Section Headings

```
DEMOGRAPHIC INFORMATION   ← Apply underline formatting in Excel
```

**Result**: Section Heading (uppercase, bold, underlined)

---

## 🎯 Key Formatting Tips

| Element             | Format Rule                  | Example                      |
| ------------------- | ---------------------------- | ---------------------------- |
| **Questions**       | Must contain `?`             | `What is your name?`         |
| **Text Input**      | Question + `___` + empty row | See above                    |
| **Multiple Choice** | Include word `select`        | `Please select all...?`      |
| **Sub-questions**   | Use **bold** formatting      | `**Customer Service**`       |
| **Section Headers** | Use underline formatting     | <u>DEMOGRAPHICS</u>          |
| **Consent Trigger** | Exact word `consent`         | `consent` (case insensitive) |

---

## ⚠️ Common Mistakes

| ❌ Wrong                    | ✅ Right                     | Issue                     |
| --------------------------- | ---------------------------- | ------------------------- |
| `What is your name`         | `What is your name?`         | Missing `?`               |
| Question + Options (no gap) | Question + `___` + empty row | Text question format      |
| Empty rows between options  | Options grouped together     | Breaks option detection   |
| `CONSENT:`                  | `consent`                    | Must be exact word        |
| Regular text for headers    | Underlined text              | Section heading detection |

---

## 📊 Output ID Patterns

| Question Type   | Catalogue ID   | Question ID   |
| --------------- | -------------- | ------------- |
| Text Question   | `PREFIX_01`    | `SHORT_01`    |
| Tickboxes       | `PREFIX_02`    | `SHORT_02`    |
| Multiple Choice | `PREFIX_03a`   | `SHORT_03a`   |
| Instructions    | `PREFIX_04_01` | `SHORT_04_01` |

---

## 🚀 Quick Start Template

```
Row | Content
----|--------
1   | This survey measures customer satisfaction.
2   | Please answer all questions honestly.
3   | consent
4   | By continuing, you agree to our privacy policy.
5   | What is your name?
6   | ___________________
7   |
8   | What is your age?
9   | Under 18
10  | 18-25
11  | 26-35
12  | Over 35
13  |
14  | Please select all that apply - which do you use?
15  | Email
16  | Phone
17  | Text
```

**💡 Pro Tip**: Start with this template and modify as needed!
