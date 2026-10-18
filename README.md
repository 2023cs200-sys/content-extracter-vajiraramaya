# Website content collector

This script starts at `START_URL`, follows navigation and content links on that exact host only, and creates `website_content.docx` with extracted headings, paragraphs, lists, tables, source URLs, and image URLs plus available alt text/captions. To collect another website, change `START_URL` in `collect_vajirarama.py`.

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

7. When it finishes, open `website_content.docx` in the same folder.

## Notes
- The default crawl is restricted to the exact `lib.vajirarama.lk` host. Links to `vajirarama.lk`, `www.vajirarama.lk`, `bosat.vajirarama.lk`, and all other hosts are excluded.
- The start URL is configured near the top of `collect_vajirarama.py`:

  ```python
  START_URL = "https://lib.vajirarama.lk/"
  ```

- Pages are limited to `MAX_PAGES = 500` by default. Increase it carefully for larger sites.
- It includes URLs for images rather than downloading/embedding the image files.
- Websites sometimes load content dynamically or block automated requests. Check the “Collection notes” section in the generated Word file for failures and manually review important pages against their source URLs.
- Use the collected content in accordance with the website's copyright and reuse terms.
