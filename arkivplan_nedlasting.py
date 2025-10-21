import asyncio
from playwright.async_api import async_playwright
import os

# Load SharePoint page URLs from a text file
with open("page_urls.txt", "r") as f:
    urls = [line.strip() for line in f if line.strip()]

# Create output folder for PDFs
os.makedirs("pdfs", exist_ok=True)

async def run():
    async with async_playwright() as p:
        # Launch browser with saved login session
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(storage_state="auth.json")

        for url in urls:
            page = await context.new_page()
            print(f"Henter: {url}")

            # Retry logic for page.goto
            for attempt in range(3):
                try:
                    await page.goto(url, wait_until="load", timeout=60000)
                    break
                except Exception as e:
                    print(f"⏳ Forsøker nummer {attempt+1} for {url} på grunn av feil: {e}")
                    await page.wait_for_timeout(5000)
            else:
                print(f"Klarte ikke å laste {url} etter tre forsøk.")
                await page.close()
                continue

            # Wait for collapsible headers to appear
            try:
                await page.wait_for_selector('[data-automation-id="CollapsibleLayer-Heading"]', timeout=1000)

                # Expand all collapsible sections
                await page.evaluate("""
                    () => {
                        document.querySelectorAll('[data-automation-id="CollapsibleLayer-Heading"]').forEach(header => {
                            const clickable = header.querySelector('a') || header;
                            const content = header.nextElementSibling;

                            // Check if the section is collapsed (content is hidden)
                            const isCollapsed = content && content.offsetParent === null;

                            if (clickable && isCollapsed) {
                                clickable.click();
                            }
                        });
                    }
                """)
                await page.wait_for_timeout(1000)
            except:
                print("Alle headere er åpne.")

            # Save page as PDF
            filename = url.split("/")[-1].replace(".aspx", ".pdf")
            await page.pdf(path=f"pdfs/{filename}", format="A4", print_background=True)
            print(f"✓ Lagret til pdfs/{filename}")
            await page.close()

        await browser.close()

asyncio.run(run())