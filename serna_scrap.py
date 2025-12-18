from playwright.sync_api import sync_playwright, TimeoutError
import time
import re
import xlsxwriter
import requests
from io import BytesIO
import os

PROFILE_DIR = "./profile"
os.makedirs(PROFILE_DIR, exist_ok=True)


def extract_company_id(page):
    html = page.content()
    m = re.search(r"urn:li:fsd_company:(\d+)", html)
    return m.group(1) if m else None


def safe_text(locator):
    try:
        return locator.text_content().strip()
    except:
        return ""


def download_image_bytes(url):
    try:
        r = requests.get(url, timeout=15)
        if r.status_code == 200:
            return BytesIO(r.content)
    except:
        pass
    return None


def extract_contact_info(profile_page):
    phones = []
    emails = []

    try:
        contact_btn = profile_page.locator(
            "#top-card-text-details-contact-info"
        )

        if contact_btn.count() == 0:
            return "", ""

        # Click safely
        contact_btn.click(force=True)

        # 🔴 IMPORTANT: wait for REAL content, not modal shell
        profile_page.wait_for_selector(
            "section.pv-contact-info__contact-type",
            timeout=10000
        )

        sections = profile_page.locator(
            "section.pv-contact-info__contact-type"
        )

        for i in range(sections.count()):
            section = sections.nth(i)

            header = safe_text(
                section.locator("h3")
            )

            # PHONE
            if header.strip() == "Phone":
                phone_spans = section.locator(
                    "span.t-black.t-normal"
                )
                for j in range(phone_spans.count()):
                    value = phone_spans.nth(j).text_content().strip()
                    if value.startswith("+"):
                        phones.append(value)

            # EMAIL
            if header.strip() == "Email":
                email_links = section.locator(
                    "a[href^='mailto:']"
                )
                for j in range(email_links.count()):
                    value = email_links.nth(j).text_content().strip()
                    if "@" in value:
                        emails.append(value)

        # Close modal
        profile_page.keyboard.press("Escape")
        time.sleep(0.5)

    except:
        try:
            profile_page.keyboard.press("Escape")
        except:
            pass

    return ", ".join(phones), ", ".join(emails)
    print("DEBUG:", phone, email)



def main():
    company_url = input("[*] Enter LinkedIn company URL: ").strip()
    pages_input = input("[*] How many pages to scrape? (ENTER = ALL): ").strip()
    excel_name = input("[*] Enter Excel file name (without .xlsx): ").strip()

    if not excel_name:
        excel_name = "employees"

    excel_file = f"{excel_name}.xlsx"
    max_pages = int(pages_input) if pages_input else None

    rows = []

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=True,
            args=["--start-maximized"]
        )

        page = context.new_page()
        page.goto(company_url, timeout=60000)

        print("[*] Solve CAPTCHA if shown")
        input("[*] Press ENTER to continue...")

        company_id = extract_company_id(page)
        if not company_id:
            print("[FATAL] Company ID not found")
            return

        search_url = (
            "https://www.linkedin.com/search/results/people/"
            f"?currentCompany=%5B%22{company_id}%22%5D"
        )

        page.goto(search_url, timeout=60000)
        time.sleep(5)

        page_num = 0

        while True:
            page_num += 1
            print(f"\n[*] Processing page {page_num}")

            avatars = page.locator("img.EntityPhoto-circle-3")
            count = avatars.count()

            for i in range(count):
                profile_page = None

                try:
                    img = avatars.nth(i)
                    img.scroll_into_view_if_needed()
                    time.sleep(1)

                    ghost = img.locator(
                        "xpath=ancestor::div[contains(@class,'ghost')]"
                    )
                    if ghost.count() > 0:
                        continue

                    link = img.locator("xpath=ancestor::a[1]")
                    if link.count() == 0:
                        continue

                    href = link.get_attribute("href") or ""
                    if "headless" in href:
                        continue

                    with context.expect_page() as new_page_info:
                        link.click(button="middle")

                    profile_page = new_page_info.value

                    profile_page.wait_for_selector(
                        "img.pv-top-card-profile-picture__image--show",
                        timeout=20000
                    )

                    name = safe_text(
                        profile_page.locator("h1.break-words").first
                    )

                    if not name or "LinkedIn Member" in name:
                        profile_page.close()
                        continue

                    bio = safe_text(
                        profile_page.locator(
                            "div.text-body-medium.break-words"
                        ).first
                    )

                    profile_url = profile_page.url

                    image_data = None
                    photo = profile_page.locator(
                        "img.pv-top-card-profile-picture__image--show"
                    )
                    if photo.count() > 0:
                        src = photo.first.get_attribute("src")
                        if src:
                            image_data = download_image_bytes(src)

                    phone, email = extract_contact_info(profile_page)

                    print(f"    [+] {name}")

                    rows.append([
                        image_data,
                        name,
                        bio,
                        profile_url,
                        phone,
                        email
                    ])

                    profile_page.close()
                    time.sleep(1)

                except TimeoutError:
                    if profile_page:
                        profile_page.close()
                except Exception:
                    if profile_page:
                        profile_page.close()

            if max_pages and page_num >= max_pages:
                break

            next_btn = page.locator(
                "button.artdeco-pagination__button--next"
            )

            if next_btn.count() == 0:
                break

            btn_class = next_btn.get_attribute("class")
            if btn_class and "artdeco-button--disabled" in btn_class:
                break

            next_btn.scroll_into_view_if_needed()
            next_btn.click()
            time.sleep(5)

        context.close()

    # ===== SAVE TO EXCEL (FIXED LAYOUT) =====
    workbook = xlsxwriter.Workbook(excel_file)
    sheet = workbook.add_worksheet("Employees")

    wrap_format = workbook.add_format({
        "text_wrap": True,
        "valign": "top"
    })

    sheet.write_row(
        0, 0,
        ["Photo", "Name", "Bio", "Profile URL", "Phone", "Email"]
    )

    sheet.set_column(0, 0, 18)
    sheet.set_column(1, 1, 25)
    sheet.set_column(2, 2, 40)
    sheet.set_column(3, 3, 45)
    sheet.set_column(4, 5, 30)

    row_num = 1

    for row in rows:
        image_data, name, bio, profile_url, phone, email = row

        sheet.set_row(row_num, 120)

        sheet.write(row_num, 1, name, wrap_format)
        sheet.write(row_num, 2, bio, wrap_format)
        sheet.write(row_num, 3, profile_url, wrap_format)
        sheet.write(row_num, 4, phone, wrap_format)
        sheet.write(row_num, 5, email, wrap_format)

        if image_data:
            sheet.insert_image(
                row_num, 0, "photo.jpg",
                {
                    "image_data": image_data,
                    "x_scale": 0.45,
                    "y_scale": 0.45,
                    "object_position": 1
                }
            )

        row_num += 1

    workbook.close()
    print(f"\n[+] Saved {len(rows)} profiles to {excel_file}")


if __name__ == "__main__":
    main()
