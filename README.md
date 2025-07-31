# DOCX Font Extractor

A simple Python tool to extract all unique fonts used in a Microsoft Word `.docx` file.

## 🔍 Features

- Extracts fonts directly from `document.xml` inside `.docx` files
- Detects font declarations including `ascii`, `hAnsi`, and `cs`
- Automatically asks the user for the file path
- Deduplicates font entries
- Lightweight and dependency-free (uses only Python standard library)

---

## 📦 Requirements

- Python 3.6+
- A `.docx` file (not `.doc`) – `.doc` is not supported

---

## 🚀 How to Use

1. Clone the repository:
    ```bash
    git clone https://github.com/MohFathalla/docx-font-extractor.git
    cd docx-font-extractor
    ```

2. Run the script:
    ```bash
    python extract_fonts.py
    ```

3. Enter the full file path when prompted, for example:
    ```
    Enter the full path to the .docx file: d:\1\example.docx
    ```

4. See the output:
    ```
    Unique fonts found in the Word document:
    Calibri
    Times New Roman
    Arial
    ```

---

## 🧠 How It Works

- The script opens the `.docx` as a ZIP file
- Parses `word/document.xml` using `xml.etree.ElementTree`
- Looks for `<w:rFonts>` tags and extracts font attributes like:
  - `ascii`
  - `hAnsi`
  - `cs`
- Outputs a list of all unique fonts

---

## 🛑 Limitations

- Only works with `.docx` files (not `.doc`)
- Does not extract fonts from styles or theme parts

---

## 📄 License

This project is licensed under the MIT License.

---

## 🙋‍♂️ Author

**MohFathalla**  
GitHub: [@MohFathalla](https://github.com/MohFathalla)
