# Q3 vs Q4 Semantic Comparison Tool (WA & Medicaid)

This project provides an automated comparison system that detects **semantic differences** between Q3 and Q4 authorization matrices for **Washington (WA)** and **Medicaid** codes.

The script uses **Google Gemini Embeddings**, **Pandas**, and **OpenPyXL** to generate:

- A semantic difference report  
- Color-coded changes by severity  
- Two comparison sheets saved inside the same Excel file  
- A summary section for each comparison  
- Clean, NaN-free output  

---

## 🚀 Features

### ✔ Semantic Meaning Comparison  
Uses Gemini Embeddings (`text-embedding-004`) to compare **meaning**, not raw text.

### ✔ Detects Multiple Change Types
- 🔴 Severe Change  
- 🟡 Moderate Change  
- 🟢 Minor Wording Change  
- 🔵 New in Q4  
- ❌ Removed in Q4  
- ⚪ No Change  

### ✔ Supports 4 Input Sheets
- `WA Q3`
- `WA Q4`
- `Medicaid Q3`
- `Medicaid Q4`

### ✔ Output Sheets Created:
- `WA Q3 vs WA Q4`
- `Medicaid Q3 vs Medicaid Q4`

### ✔ Summary Section Automatically Added  
Example:

```
SUMMARY
Total in Q3: xx
No Change: xx
Modified: xx
Severe Change: xx
Moderate Change: xx
Minor Change: xx
New in Q4: xx
Removed in Q4: xx
```

### ✔ Conditional Formatting Color Codes

| Severity                | Color |
|------------------------|--------|
| Severe Change          | 🔴 Red |
| Moderate Change        | 🟡 Yellow |
| Minor Wording Change   | 🟢 Light Green |
| New Entry              | 🔵 Light Blue |
| No Change              | ⚪ White |

---

## 📦 Installation

Install required packages:

```bash
pip install pandas numpy openpyxl google-generativeai python-dotenv
```

---

## 🔑 Environment Setup

Create a `.env` file in the project directory:

```
GOOGLE_API_KEY=your_api_key_here
```

This is automatically loaded using:

```python
from dotenv import load_dotenv
load_dotenv()
```

---

## 📂 Expected Excel Structure

Your input Excel file must include these sheets:

| Sheet Name     | Required Columns                          |
|----------------|--------------------------------------------|
| WA Q3          | Code, Code Notes                           |
| WA Q4          | Code, Code Notes                           |
| Medicaid Q3    | Code, MHI Code Notes                       |
| Medicaid Q4    | Code, MHI Code Notes                       |

The script automatically normalizes column names to avoid hidden characters.

---

## 🧠 How Semantic Comparison Works

### 1️⃣ Convert Q3 and Q4 text → embeddings  
Using Gemini:

```
embed(text) → vector
```

### 2️⃣ Compute cosine similarity  
```
similarity = cosine(old_vector, new_vector)
```

### 3️⃣ Evaluate severity  
| Similarity | Severity |
|------------|----------|
| 0.00–0.55  | Severe Change |
| 0.55–0.80  | Moderate Change |
| 0.80–0.99  | Minor Wording Change |
| 1.0        | No Change |

### 4️⃣ Detect added/removed codes  
Compares Q3 vs Q4 code sets.

---

## 📊 Output Format

The output sheets include:

| Code | Status | Column | Q3 Value | Q4 Value | Similarity | Severity |
|------|--------|--------|----------|----------|------------|----------|

Example:

| 80305 | Modified | Code Notes | No PA 24 visits | No PA 12 visits | 0.46 | Severe Change |
| 97153 | Removed in Q4 | Code Notes | ... | | | Severe Change |
| G0481 | No Change | Code Notes | ... | ... | 1.0 | No Change |
| 15829 | New in Q4 | Code Notes | | carve-out text | | New Entry |

---

## 🎨 Conditional Formatting

Colors are applied using `openpyxl.styles.PatternFill`.

Rows are colored automatically based on severity and start after summary rows.

---

## ▶️ Running the Script

Update the filename and run:

```bash
python compare_script.py
```

or inside the file:

```python
process_file("Authorization Business Matrix 2025 Q3 - WA and Medicaid - Reference.xlsx")
```

After execution, the Excel file will contain:

- ✔ WA Q3 vs WA Q4  
- ✔ Medicaid Q3 vs Medicaid Q4  
- ✔ Summary at the top  
- ✔ Color-coded rows  

---

## 🧩 Error Handling

### Missing column names?
The script auto-detects using:

```
find_column()
```

If still not found, it raises:

```
❌ None of these columns found: [...]
```

### Missing `.env` key?
Gemini will fail authentication.

---

## 🧪 Example Folder Structure

```
project/
│── compare_script.py
│── Authorization Business Matrix.xlsx
│── .env
│── README.md
```

---

## 💡 Future Enhancements (optional)

- Export summary as PDF  
- Create charts (bar graphs of changes)  
- Streamlit web UI to upload Excel files  
- Support for multi-state comparison  

---

## 🎉 Conclusion

This tool allows analysts to **quickly identify meaningful changes** between quarterly authorization matrices using *AI-powered semantic comparison*, clean reporting, and a user-friendly Excel output.

If you need feature upgrades, integrations, or UI improvements, feel free to ask!
