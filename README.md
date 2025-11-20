# pptx-font-extractor

**pptx-font-extractor** is a Python utility for automatically extracting and listing every font used in a PowerPoint file (`.pptx`), including those inside tables. It then tries to find matching font files in your Windows fonts directory, copies them to a `fonts` folder, and saves a report on which fonts were found, copied, or unavailable.

***

## Features

- **Detects all fonts used** (including those in tables/cells) in a PPTX file.
- **Matches font names** with Windows system font files (`.ttf` / `.otf`).
- **Copies found fonts** to a local folder for easy packaging/sharing.
- **Generates a detailed font usage & copy report**.

***

## How It Works

1. **Load your PPTX file** (`Presentation("yourfile.pptx")`).
2. **Scan all slides and shapes** for detected fonts (text frames & tables).
3. **Collect font names** found in text runs and table cells.
4. **Look for matches** in your `C:\Windows\Fonts` directory.
5. **Copy matched fonts** to a local `fonts` folder inside your repo.
6. **Log status** to `fonts/_used_fonts_.txt` (copied/unavailable).
7. **Print summary** to console after running.

***

## How to Use

1. **Install dependencies**  
   You need:  
   - Python (>=3.6)  
   - [python-pptx](https://python-pptx.readthedocs.io) (`pip install python-pptx` or `pip install -r requirements.txt`)

2. **Place your .pptx file**  
   Put your PowerPoint file in the repo folder (e.g., `My-Presentation.pptx`).

3. **Edit script filename if needed**  
   Change the PPTX filename in the script:
   ```python
   prs = Presentation("YOUR_PRESENTATION.pptx")
   ```

4. **Run the script**
   ```bash
   python pptx-font-extractor.py
   ```

5. **Check results**  
   - Find fonts in the new `fonts` folder.
   - Read the font usage & copy log in `fonts/_used_fonts_.txt`.

***

## Notes

- Only works on Windows, where `C:\Windows\Fonts` is available.
- Font filename matching is broad but may miss non-standard names.
- Script does not install fonts; it just copies the available files.

***

**License:** MIT
**Author:** [Ryan Heida](https://links.ryanheida.com)

Feel free to improve or contribute!