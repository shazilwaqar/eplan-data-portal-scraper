from playwright_utils import launch_live_dynamic_engine , extract_safe_attribute , extract_safe_text , append_to_csv , convert_csv_to_styled_excel

url = "https://dataportal.eplan.com/parts/list"

item_selector = "li.list-item.ng-star-inserted"

page, browser, p = launch_live_dynamic_engine(url, item_selector)

if page:
    row_count = 0
    for page_no in range(1, 11):
        url = f"https://dataportal.eplan.com/parts/list?page={page_no}"
        try:
            page.goto(url, timeout=60000)
            page.wait_for_selector(item_selector, timeout=20000)
            items = page.locator(item_selector).all()

            for i in range(len(items)):
                try:
                    items = page.locator(item_selector).all()
                    current_item = items[i]
                
                    part_no_selector = current_item.locator("a[aria-label='part-details']").first

                    part_no = extract_safe_attribute(part_no_selector, "", "data-test")

                    manufacturer = extract_safe_text(current_item, "dt:text('Manufacturer') + dd.part-list-dd")

                    designation = extract_safe_text(current_item, "dt:text('Designation') + dd.part-list-dd")

                    detail_page = extract_safe_attribute(part_no_selector,"", "href")
                    link = f"https://dataportal.eplan.com{detail_page}"

                    page.goto(link, timeout=70000)
                    page.wait_for_selector("div.data-container table", timeout=40000)
                    table_locator = page.locator("div.data-container")

                    raw_decription = extract_safe_text(table_locator, "td:text('Description') + td.part-property-value", default="None ")
                    cleaned_description = raw_decription.replace("EN_US", " ").strip() if raw_decription != "None " else "None "
                    product_group = extract_safe_text(table_locator, "td:text(' Product group ') + td.part-property-value", default="None ")

                    data = {
                        "Part Number": part_no,
                        "Manufacturer": manufacturer,
                        "Designation": designation,
                        "Description": cleaned_description,
                        "Product Group": product_group,
                        "Link": link,
                    }
                
                    append_to_csv("eplan_data.csv", data)
                    row_count += 1
                except Exception as e:
                    print(f"ERROR on item {i} of page {page_no}: {e}")
                page.wait_for_timeout(3000)
                page.goto(url, timeout=60000)
                page.wait_for_selector(item_selector, timeout=10000)
        except Exception as e:
            print(f"ERROR on page {page_no} (URL: {url}): {e}")



    page.close()
    browser.close()
    p.stop()
    print(f"Scraping completed. Total rows scraped: {row_count}")

    convert_csv_to_styled_excel(csv_path="eplan_data.csv" , excel_path="eplan_data.xlsx")