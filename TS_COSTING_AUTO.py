from flask import Flask
from playwright.async_api import async_playwright
import nest_asyncio
import asyncio
from EXCEL_EXTRACT import get_style_no
from EXCEL_EXTRACT import get_target_ts_value
from EXCEL_EXTRACT import get_notes
from EXCEL_EXTRACT import get_labels
from EXCEL_EXTRACT import get_labour
from EXCEL_EXTRACT import get_wash
from EXCEL_EXTRACT import get_dox
from EXCEL_EXTRACT import get_finance
from EXCEL_EXTRACT import get_testing
from EXCEL_EXTRACT import get_markup
from EXCEL_EXTRACT import main_excel
import pandas as pd
nest_asyncio.apply()

app = Flask(__name__)


async def run_playwright():
    browser = None
    style_no = get_style_no()
    target_description =get_target_ts_value()
    notes =get_notes()
    labels = get_labels()
    labour = get_labour()
    wash = get_wash()
    docs =get_dox()
    finance =get_finance()
    testing = get_testing()
    markup =get_markup()
    df, style_excel, new_df, trim_df = main_excel()


    try:
        async with async_playwright() as p:

            # Launch Browser
            browser = await p.chromium.launch(headless=False, channel="chrome")
            context = await browser.new_context(
                accept_downloads=True
                
            )
            page = await context.new_page()

            # -------------------------
            # OPEN LOGIN PAGE
            # -------------------------
            for attempt in range(5):
                try:
                    await page.goto("https://urban.bamboorose.com/prod/", timeout=30000)
                    print("✅ Page loaded")
                    break
                except Exception as e:
                    print(f"[Attempt {attempt+1}] Failed: {e}")
                    await page.wait_for_timeout(4000)

            # -------------------------
            # LOGIN
            # -------------------------
            await page.fill('input[name="user_id_show"]', "VEN-97246-06")
            pwd = page.locator('input[type="password"].passwordInput')
            await pwd.click()
            await pwd.fill("Sahu12345678")
            await page.click('#formsubmit')
            await page.wait_for_timeout(4000)

            # -------------------------
            # DASHBOARD
            # -------------------------
            element = await page.query_selector('#B1')
            if element:
                await element.click()
                print("✅ Dashboard clicked")
            await page.wait_for_timeout(2000)

            # STYLE SEARCH
            await page.click('span[role="listbox"]')
            await page.wait_for_selector('#quickSearchDocument_listbox li')
            await page.click('li:has-text("Style")')
            await page.wait_for_timeout(2000)

            await page.fill('#quickSearchInput', style_no)
            await page.keyboard.press("Enter")
            await page.wait_for_timeout(6000)

            print("⏳ Waiting for first row...")

            # -------------------------
            # FIRST RESULT CLICK
            # -------------------------
            for attempt in range(12):
                try:
                    await page.wait_for_selector('td.clsBrdBtmRt a', timeout=4000)
                    first = page.locator('td.clsBrdBtmRt a').first
                    await first.scroll_into_view_if_needed()
                    await first.click()
                    print("✅ FIRST RESULT CLICKED")
                    break
                except Exception as e:
                    print("Retry:", e)
                    await page.wait_for_timeout(1000)
            else:
                raise Exception("❌ Could not click first result")

            # -------------------------
            # DETAILS TAB
            # -------------------------
            await page.wait_for_timeout(3000)
            details = page.locator('td:has-text("Details")').nth(2)
            await details.click(force=True)
            print("✅ DETAILS TAB CLICKED")

            # -------------------------
            # LOAD OFFER TABLE
            # -------------------------
            print("📥 Loading Offer Table...")
            await page.wait_for_selector('#detail_section_tabDetail', timeout=20000)
            await page.wait_for_timeout(3000)

            rows = await page.query_selector_all('#detail_section_tabDetail tr[class^="row"]')

            if not rows:
                raise Exception("❌ No offer rows found")

            table_data = []

            for row in rows:
                cells = await row.query_selector_all("td")
                row_data = [(await cell.inner_text()).replace("\n", " ").strip() for cell in cells]
                table_data.append(row_data)

            # PRINT TABLE
            print("\n======= ✅ FULL OFFER TABLE DATA =======\n")
            for i, row in enumerate(table_data):
                print(f"ROW {i+1}: {row}")

            # -------------------------
            # FIND HIGHEST OFFER #
            # -------------------------
            print("\n🔎 Finding highest Offer Number...")

            offer_numbers = []

            for row in rows:
                cells = await row.query_selector_all("td")

                if len(cells) >= 3:
                    link = await cells[2].query_selector("a")
                    if link:
                        txt = (await link.inner_text()).strip()
                        if txt.isdigit():
                            offer_numbers.append(int(txt))

            if not offer_numbers:
                raise Exception("❌ No numeric Offer numbers found")

            print("📊 Offer Numbers:", offer_numbers)

            highest_offer = max(offer_numbers)
            print("🏆 Highest Offer:", highest_offer)

            # -------------------------
            # ✅ FIXED SAFE CLICK SYSTEM
            # -------------------------
            print("🎯 Opening highest Offer...")
            opened = False

            for attempt in range(6):
                print(f"🔁 Attempt {attempt+1}")

                offer_link = page.locator(
                    f"#detail_section_tabDetail a:has-text('{highest_offer}')"
                ).first

                try:
                    await offer_link.wait_for(state="visible", timeout=8000)
                    await offer_link.scroll_into_view_if_needed()
                    await offer_link.click(force=True)

                    # ✅ FIX: detect successful navigation by DOM change (URL does NOT change here)
                    try:
                        await page.wait_for_selector("#detail_section_tabDetail", state="detached", timeout=8000)
                        print("✅ OFFER OPENED SUCCESSFULLY")
                        opened = True
                        break
                    except:
                        print("⚠️ Clicked but Offer list still visible")

                except Exception as e:
                    print("⏳ Retry error:", e)

            if not opened:
                raise Exception("❌ OFFER CLICKED BUT PAGE NEVER OPENED")

            print("✅ OFFER PAGE CONFIRMED OPEN")
            await page.wait_for_timeout(4000)

#             # -------------------------
#             # COPY OFFER
#             # -------------------------
#             xpath = '/html/body/form/div[1]/div/div[2]/div[4]/div/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/table[2]/tbody/tr/td/div/div[2]/a'

#             print("🎯 Clicking target element...")

#             try:
#                  element = page.locator(f'xpath={xpath}')
#                  await element.wait_for(state="visible", timeout=15000) 
#                  await element.click(force=True)
#                  print("✅ CLICKED SUCCESSFULLY")

#             except Exception as e:
#                 print("❌ XPATH CLICK FAILED:", e)
#             await page.wait_for_timeout(5000)

#             description_input = page.locator('input[id="3_@100_@63_@2_@0_@0"]')

#             await description_input.click()
#             await description_input.fill(target_description)
#             await page.wait_for_timeout(2000)
#             desc_qty = page.locator('input[id="3_@100_@25_@2_@0_@0"]')
#             await desc_qty.click()
#             await desc_qty.fill('3000')
#             await page.wait_for_timeout(1000)
#             desc_price = page.locator('input[id="3_@100_@121_@2_@0_@0"]')
#             await desc_price.click()
#             await desc_price.fill('1500')
#             await page.wait_for_timeout(1000)
#             save_btn = page.locator("#OFFER_SAVEbtnCtr")
#             await save_btn.click(force=True)
#             await page.wait_for_timeout(2000)

#             country = page.locator('input[id="1_@100_@23_@2_@0_@0_@desc"]')
#             await country.click()
#             await country.type('INDIA')
#             # Press Down arrow 3 times
#             for _ in range(4):
#                await page.keyboard.press("ArrowDown")
#                await page.wait_for_timeout(1000)

# # Press Enter
#                await page.keyboard.press("Enter")
#             await page.wait_for_timeout(1000)
#             notes_fill = page.locator('//textarea[@id="14_@10100_@16_@0_@0_@0"]')

#             await notes_fill.click()
#             await notes_fill.fill(notes)
#             await page.wait_for_timeout(2000)
#             save_btn = page.locator("#OFFER_SAVEbtnCtr")
#             await save_btn.click(force=True)
#             await page.wait_for_timeout(2000)
            cost_bom = page.locator('td:has-text("Cost BOM")').nth(2)
            await cost_bom.click(force=True)
            await page.wait_for_timeout(3000)

            # Label_T = page.locator("//select[@id='0_@16700_@2_@0_@0_@0']")
            # await Label_T.click()
            # await Label_T.select_option(label="LABELS & TICKETING")

            # Label_T_value  = page.locator('input[id="0_@16700_@3_@0_@0_@0"]')
            # await Label_T_value.click()
            # await Label_T_value.fill(labels)

            # await page.wait_for_timeout(500)

            # cutnmake = page.locator("//select[@id='0_@16700_@2_@0_@1_@1']")
            # await cutnmake.click()
            # await cutnmake.select_option(label="LABOR")

            # cutnmake_value = page.locator('input[id="0_@16700_@3_@0_@1_@1"]')
            # await cutnmake_value.click()
            # await cutnmake_value.fill(labour)

            # await page.wait_for_timeout(500)

            # wash_t = page.locator("//select[@id='0_@16700_@2_@0_@2_@2']")
            # await wash_t.click()
            # await wash_t.select_option(label="WASH")

            # wash_v = page.locator('input[id="0_@16700_@3_@0_@2_@2"]')
            # await wash_v.click()
            # await wash_v.fill(wash)
            # await page.wait_for_timeout(500)


            # dox= page.locator("//select[@id='0_@16700_@2_@0_@3_@3']")
            # await dox.click()
            # await dox.select_option(label="CUSTOM/DOC FEES")

            # dox_v = page.locator('input[id="0_@16700_@3_@0_@3_@3"]')
            # await dox_v.click()
            # await dox_v.fill(docs)

            # await page.wait_for_timeout(500)

            # fin= page.locator("//select[@id='0_@16700_@2_@0_@4_@4']")
            # await fin.click()
            # await fin.select_option(label="FINANCE & HANDLING")

            # fin_v = page.locator('input[id="0_@16700_@3_@0_@4_@4"]')
            # await fin_v.click()
            # await fin_v.fill(finance)

            # await page.wait_for_timeout(500)


            # test = page.locator('//select[@id="0_@16700_@2_@0_@5_@5"]')
            # await test.click()
            # await test.select_option(label="TESTING")

            # test_v = page.locator('//input[@id="0_@16700_@3_@0_@5_@5"]')
            # await test_v.click()
            # await test_v.fill(testing)
            # await page.wait_for_timeout(500)

            # mark_up = page.locator('//select[@id="0_@16700_@2_@0_@6_@6"]')
            # await mark_up.click()
            # await mark_up.select_option(label="MARK UP/OVERHEAD")

            # mark_v = page.locator('//input[@id="0_@16700_@3_@0_@6_@6"]')
            # await mark_v.click()
            # await mark_v.fill(markup)

            # f_dropdown = page.locator('//select[@id="0_@16700_@2_@0_@8_@8"]')
            # await f_dropdown.click()
            # await f_dropdown.select_option(label="MATERIAL FREIGHT")

            # t_dropdown = page.locator('//select[@id="0_@16700_@2_@0_@9_@9"]')
            # await t_dropdown.click()
            # await t_dropdown.select_option(label="PACKAGING")
            # await page.keyboard.press("Enter")

            await page.wait_for_timeout(2000)
            await page.mouse.wheel(0, 3000)
      
            bom_rows = await page.query_selector_all("//input[contains(@id,'_@4_@1_@') and contains(@id,'@desc')]")
            total_rows = len(bom_rows)

            print(f"📊 TOTAL BOM ROWS AVAILABLE: {total_rows}")

            for i, row in new_df.reset_index(drop=True).iterrows():

                if i >= total_rows:
                  print(f"⚠️ Excel has more rows than UI ({i+1} skipped)")
                  break

                code = str(row["Code"])
                desc = str(row["DESCRIPTION"])
                price = str(row["PRICE"])
                yy = str(row["YY"])

                print(f"🔹 Filling BOM row {i+1}")

    # ✅ CODE (CLICK ONLY — READONLY INPUT)
                code_input = page.locator(f"(//input[contains(@id,'_@4_@1_@') and contains(@id,'@desc')])[{i+1}]")
                await code_input.click(force=True)
                await code_input.evaluate("el => el.value = ''")
                await code_input.evaluate(f"el => el.value = `{code}`")   # JS SET
                await code_input.dispatch_event("change")

                await page.wait_for_timeout(1000)


    # ✅ DESCRIPTION
                desc_input = page.locator(f"(//input[contains(@id,'_@3_@1_@')])[{i+1}]")
                await desc_input.click(force=True)
                await desc_input.type(desc)
                print("DESC FILLED")
                await page.wait_for_timeout(1000)

    # ✅ YY
                yy_input = page.locator(f"(//input[contains(@id,'_@10_@1_@')])[{i+1}]")
                await yy_input.click(force=True)
                await yy_input.type(yy)
                print("FILLED YY")
                await page.wait_for_timeout(1000)

    # ✅ PRICE
                price_input = page.locator(f"(//input[contains(@id,'_@31_@1_@')])[{i+1}]")
                await price_input.click(force=True)
                await price_input.type(price) 
                print("✅ COMPONENT TABLE FILLED")
                await page.wait_for_timeout(4000)
                print("complete")

            await page.mouse.wheel(0, 3000)
            await page.wait_for_timeout(4000)

            await page.locator("td[onclick*='14001_30001']").nth(1).click()
            await page.locator("td[onclick*='14001_30002']").nth(1).click()
            print("HERE DONE")

            await page.wait_for_timeout(4000)

            # ============================================
            # TRIM TABLE FILLING
            # ============================================
            print("\n======================")
            print("🔎 TRIM TABLE FILLING")
            print("======================\n")

            trim_section = page.locator('[id="14001_30002"]')
            await trim_section.wait_for(state="visible")

            component_cells = trim_section.locator("//input[contains(@id,'_@4_@1_@') and contains(@id,'@desc')]")
            desc_cells = trim_section.locator("//input[contains(@id,'_@3_@1_@')]")

            ui_rows = await component_cells.count()
            excel_rows = len(trim_df)

            print("UI Rows    :", ui_rows)
            print("Excel Rows :", excel_rows)

            filled = 0
            skipped = 0
            excel_index = 0

            for i in range(ui_rows):

                # ✅ EXIT WHEN ALL EXCEL ROWS USED
                if excel_index >= excel_rows:
                    print("✅ All Excel rows inserted. Exiting loop.")
                    break

                comp = component_cells.nth(i)
                desc_input = desc_cells.nth(i)

                await comp.scroll_into_view_if_needed()
                comp_val = (await comp.input_value()).strip()

                # ✅ SKIP FILLED COMPONENTS
                if comp_val:
                    print(f"⏭ UI Row {i+1} skipped (Component exists): {comp_val}")
                    skipped += 1
                    continue

                # ✅ FETCH EXCEL ROW
                row = trim_df.iloc[excel_index]
                description = str(row["components"])
                price = str(row["Price"])
                yy = str(row["YY"])

                print(f"\n✅ UI ROW {i+1}  <=  EXCEL ROW {excel_index+1}")
                print("   Desc :", description)
                print("   YY   :", yy)
                print("   Price:", price)

                # ✅ DESCRIPTION
                await desc_input.click(force=True)
                await desc_input.press("Control+A")
                await desc_input.type(description)

                # ✅ YY
                yy_input = trim_section.locator("//input[contains(@id,'_@10_@1_@')]").nth(i)
                await yy_input.click(force=True)
                await yy_input.press("Control+A")
                await yy_input.type(yy)

                # ✅ PRICE
                price_input = trim_section.locator("//input[contains(@id,'_@31_@1_@')]").nth(i)
                await price_input.click(force=True)
                await price_input.press("Control+A")
                await price_input.type(price)

                print("✅ Row Filled")

                excel_index += 1
                filled += 1

                await page.wait_for_timeout(800)

            # ============================================
            # FINAL SUMMARY
            # ============================================
            print("\n===================================")
            print("✅ TRIM FILLING FINISHED")
            print("Excel Used :", excel_index)
            print("Filled     :", filled)
            print("Skipped    :", skipped)
            print("===================================\n")

            await page.wait_for_timeout(10000)
            
           
  
            

            


        










            


            




            
            



       









        

            

    except Exception as e:
        print("❌ Automation error:", e)

    finally:
        if browser:
            await browser.close()
            print("🧹 Browser closed")


if __name__ == "__main__":
    asyncio.run(run_playwright())
