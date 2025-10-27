from playwright.sync_api import sync_playwright
from pathlib import Path
import re
import WarrantyTime
import requests



def scrape_serial(serial):
    serial = serial.strip().upper()

    if re.fullmatch(r"[A-Z0-9]{7}", serial): # DELL
        return scrape_dell_info(serial)

    elif re.fullmatch(r"[A-Z0-9]{10}", serial): #HP
        ...

    else:# Lenovo
        return scrape_lenovo_info(serial)

    
def scrape_dell_info(serial_number: str):

    """Scrape Dell warranty info."""
    url = f"https://www.dell.com/support/productsmfe/en-us/productdetails?selection={serial_number}&assettype=svctag&appname=warranty&inccomponents=false&isolated=false"
    profile_path = Path(__file__).parent / "puppet.Default"



    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(profile_path),
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.5993.90 Safari/537.36",
            headless=True
        )

        page = context.new_page()
        page.goto(url)

        page.wait_for_selector('#warrantystatusenddatetext')
        warranty_end_date = page.locator('#warrantystatusenddatetext').inner_text().strip()
        device_name = page.locator(
            '#mfe-productdetails h5'
        ).inner_text().strip()

        context.close()

        return {
            "brand": "Dell",
            "serial_number": serial_number,
            "device_name": device_name,
            "warranty_end_date": WarrantyTime.format_date(warranty_end_date)
        }


def scrape_lenovo_info(serial_number: str):

    """
    Scrape Lenovo warranty info.
    """

    url = "https://pcsupport.lenovo.com/us/en/warranty-lookup#/"

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
             args=[
                "--window-position=-32000,-32000",
                "--window-size=1024,768",
                "--start-minimized",
            ],
            )

        page = browser.new_page()
        page.goto(url)



        germany_button_selector = "button.button.blue-solid.modal-button-top40.btn_yes"
        if page.locator(germany_button_selector).count() > 0:

            page.click(germany_button_selector)

        input_selector = "input.button-placeholder__input"

        page.wait_for_selector(input_selector)

        page.fill(input_selector, serial_number)

        submit_selector = "button.basic-search__suffix-btn.btn.btn-primary"

        page.click(submit_selector)


        warranty_selector = ".cell-detail:nth-child(1) .detail-property:nth-child(5) > .property-value"

        device_selector = "div.prod-name > h4"


        page.wait_for_selector(warranty_selector)
        page.wait_for_selector(device_selector)

        warranty_end_date = page.locator(warranty_selector).inner_text().strip()
        device_name = page.locator(device_selector).inner_text().strip()

        browser.close()

        return {
            "brand": "Lenovo",
            "serial_number": serial_number,
            "device_name": device_name,
            "warranty_end_date": WarrantyTime.format_date(warranty_end_date)
        }


def scrape_hp_info(serial_number: str):

    """
    Scrapes HP (or Sigatics-supported) device info using an API endpoint.
    """

    url = f"https://warranty-check.sigatics.com/warranty/{serial_number}"

    response = requests.get(url)

    if response.status_code != 200:
        print(f"Failed to fetch warranty info for {serial_number}")
        return None

    data = response.json()

    return {
        "brand": "HP",
        "serial_number": serial_number,
        "device_name": data.get("Product Name", ""),
        "warranty_end_date": WarrantyTime.format_date(data.get("Warranty End", ""))
    }
