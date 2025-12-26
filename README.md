# 🔗 Excel VBA – Open Multiple Hyperlinks

A simple **Excel VBA macro** that allows opening **multiple hyperlinks at once** from a selected range of cells.

The macro iterates through all hyperlinks in the chosen range and opens them one by one in the default web browser.

---

## 🎯 Purpose

This tool is useful when you need to:
- Open many URLs stored in Excel cells
- Quickly review multiple tickets, dashboards, documents, or web pages
- Avoid opening links manually one by one

Typical use cases include:
- Operations & support tracking
- Issue / ticket review
- Data validation workflows
- Analyst or operations daily routines

---

## 🧩 How It Works

1. The macro prompts the user to select a range of cells.
2. Each cell in the range is checked for hyperlinks.
3. All hyperlinks are opened sequentially using the default browser.

---

## 🛠 Macro Code

```vba
Sub OpenMultipleLinks()
    Dim MultiLink As Hyperlink
    Dim SelectedRng As Range
    On Error Resume Next

    WLTitleId = "WL Open Multiple Links"

    Set SelectedRng = Application.Selection
    Set SelectedRng = Application.InputBox("Range", WLTitleId, SelectedRng.Address, Type:=8)

    For Each MultiLink In SelectedRng.Hyperlinks
        MultiLink.Follow
    Next
End Sub
```

---

## 🚀 How to Use

1. Open Excel
2. Press **ALT + F11** to open the VBA Editor
3. Insert a new module
4. Paste the macro code
5. Close the VBA Editor
6. Select a range of cells containing hyperlinks
7. Run the macro:
   - `ALT + F8` → `OpenMultipleLinks` → **Run**

---

## ⚠️ Notes & Limitations

- Only cells containing valid hyperlinks will be opened
- Links open sequentially (browser popup blockers may apply)
- Large ranges may open many tabs at once
- Error handling is minimal by design (simple utility macro)

---

## 🧰 Requirements

- Microsoft Excel (Windows)
- VBA enabled (`.xlsm` file)
- Default web browser configured

---

## 📄 License

This project is released under the **MIT License**.  
You are free to use, modify, and distribute it.

---

## 👨‍💻 Author

**Paweł Goleń**  
Data Engineer | Automation & Productivity Tools
