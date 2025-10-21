import asyncio
import tkinter as tk
from tkinter import messagebox
from playwright.async_api import async_playwright

async def login_and_save_session():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        # Go to SharePoint login page
        await page.goto("https://tromkom.sharepoint.com/sites/O365-ProsjektervedArkiv-BK-plan")

        # Create a popup window using Tkinter
        def on_continue():
            root.destroy()

        root = tk.Tk()
        root.title("Logg inn")
        root.geometry("400x150")
        label = tk.Label(root, text="Logg inn i nettleservinduet. \nTrykk 'Fortsett' når du har logget inn.", wraplength=380, justify="center")
        label.pack(pady=20)
        continue_button = tk.Button(root, text="Fortsett", command=on_continue)
        continue_button.pack()
        root.mainloop()

        # Save session to auth.json
        await context.storage_state(path="auth.json")
        messagebox.showinfo("Innlogging lagret", "Innlogging lagret til auth.json")

        await browser.close()

# Use this instead of asyncio.run()
try:
    loop = asyncio.get_running_loop()
except RuntimeError:
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

loop.run_until_complete(login_and_save_session())

