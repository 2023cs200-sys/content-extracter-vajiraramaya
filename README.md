# Vajirarama About-section collector

This script automatically visits the Siri Vajirarama Temple About page and follows nested links under `/about-us-2/`, including linked monk biographies. It creates `vajirarama_about_content.docx` with extracted headings, paragraphs, lists, source URLs, and image URLs plus available alt text/captions.

## Requirements
- Python 3.10 or newer
- Internet access
- Permission to crawl the target site; review its `robots.txt` and terms first

## Windows setup (PowerShell)
1. Extract this folder somewhere convenient, e.g. `C:\Users\YourName\Documents\vajirarama_collector`.
2. Open that folder in VS Code.
3. Open Terminal → New Terminal.
4. Create a virtual environment (recommended):

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

5. Install dependencies:

   ```powershell
   py -m pip install -r requirements.txt
   ```

6. Run the collector:

   ```powershell
   py collect_vajirarama.py
   ```

7. When it finishes, open `vajirarama_about_content.docx` in the same folder.

## Notes
- The crawl is restricted to `https://vajirarama.lk/about-us-2/` and its descendants. It does not crawl the Noticeboard or `bosat.vajirarama.lk` subdomain.
- It includes URLs for images rather than downloading/embedding the image files.
- Websites sometimes load content dynamically or block automated requests. Check the “Collection notes” section in the generated Word file for failures and manually review important pages against their source URLs.
- Use the collected content in accordance with the website's copyright and reuse terms.
