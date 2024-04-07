# from playwright.sync_api import sync_playwright
from pyppeteer import launch
import asyncio
import os
import pandas as pd

async def fill_form(data):
    browser = await launch(executablePath="C:\Program Files\Google\Chrome\Application\chrome.exe", headless=False)
    page = await browser.newPage()
    await page.goto('https://help.instagram.com/contact/230197320740525')

    await page.click('input[name="continuereport"]')
    await page.waitFor(1000) # wait for 1 sec

    await page.click('input[value="I am reporting on behalf of my organization or client."]')
    await page.waitFor(1000) 

    # Fill in form fields
    await page.type('input[name="your_name"]', data['name'])
    await page.waitFor(1000)

    await page.type('textarea[name="Address"]', data['address'])
    await page.waitFor(1000)

    await page.type('input[name="email"]', data['email'])
    await page.waitFor(1000)

    await page.type('input[name="confirm_email"]', data['confirm_email'])
    await page.waitFor(1000)

    await page.type('input[name="reporter_name"]', data['reporter_name'])
    await page.waitFor(1000)

    await page.type('input[name="websiterightsholder"]', data['websiterightsholder'])
    await page.waitFor(1000)

    await page.type('input[name="what_is_your_trademark"]', data['what_is_your_trademark'])
    await page.waitFor(1000)

    await page.select('select[name="rights_owner_country_routing"]', 'India')
    await page.waitFor(1000)

    await page.type('textarea[name="TM_URL"]', data['TM_URL'])
    await page.waitFor(1000)

    # Upload attachment
    file_input = await page.querySelector('input[name="Attach1[]"]')
    await file_input.uploadFile(os.path.abspath('sample.pdf'))
    await page.waitFor(1000)

    await page.click('input[value="This photo, video, post or story uses rights owner\'s trademark."]')
    await page.waitFor(1000)

    await page.type('textarea[id="1622541521292980"]', data['content_urls'])
    await page.waitFor(1000)

    await page.type('textarea[name="additionalinfo"]', data['additionalinfo'])
    await page.waitFor(1000)

    file_input = await page.querySelector('input[name="content_attachment_no_urls[]"]')
    await file_input.uploadFile(os.path.abspath('sample.pdf'))
    await page.waitFor(1000)

    await page.type('input[name="signature"]', data['signature'])
    await page.waitFor(1000)

    # Hover over Send Button
    await page.hover('button[type="submit"]')
    await page.waitFor(1000)
    # Close the browser
    await browser.close()
    

def main():
    data = pd.read_excel('data.xlsx').iloc[0].to_dict()

    asyncio.get_event_loop().run_until_complete(fill_form(data))

if __name__ == "__main__":
    main()
