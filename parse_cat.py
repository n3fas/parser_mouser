import asyncio
import random
import pandas as pd
from playwright.async_api import async_playwright
from bs4 import BeautifulSoup

async def scrape_mouser():
    async with async_playwright() as p:
        # Папка для профиля, чтобы сохранять куки и не ловить капчу каждый раз
        user_data_dir = "./mouser_profile"
        
        context = await p.chromium.launch_persistent_context(
            user_data_dir,
            headless=False, # Видим окно браузера
            args=["--disable-blink-features=AutomationControlled"]
        )
        
        page = await context.new_page()

        # РУЧНОЙ СТЕЛС (маскировка под человека)
        await page.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', { get: () => false });
            window.chrome = { runtime: {} };
            Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
        """)

        base_url = "https://eu.mouser.com"
        final_data = []

        try:
            print("--- ШАГ 1: Заходим на главную страницу категорий ---")
            await page.goto(f"{base_url}/electronic-components/", wait_until="domcontentloaded")
            
            print("Ожидаем появления категорий (lnkCategory_)...")
            # Ждем именно те ID, которые ты прислал (до 60 секунд)
            await page.wait_for_selector('a[id^="lnkCategory_"]', timeout=60000)
            
            # Парсим все основные категории
            soup = BeautifulSoup(await page.content(), 'html.parser')
            category_tags = soup.find_all('a', id=lambda x: x and x.startswith('lnkCategory_'))
            
            unique_main = {}
            for a in category_tags:
                href = a.get('href', '')
                # Очищаем текст от счетчика товаров (span)
                name = a.get_text(separator='|', strip=True).split('|')[0].strip()
                
                if href and name:
                    full_url = base_url + href if href.startswith('/') else href
                    clean_url = full_url.split('?')[0]
                    unique_main[clean_url] = name

            print(f"Найдено {len(unique_main)} основных категорий. Начинаем углубление...")

            # --- ШАГ 2: Проход по каждой категории для поиска ПОДкатегорий ---
            for i, (url, parent_name) in enumerate(unique_main.items()):
                print(f"[{i+1}/{len(unique_main)}] Ищем подкатегории в: {parent_name}")
                
                try:
                    # Рандомная задержка 4-8 секунд
                    await page.wait_for_timeout(random.randint(4000, 8000))
                    await page.goto(url, wait_until="domcontentloaded")
                    
                    inner_soup = BeautifulSoup(await page.content(), 'html.parser')
                    
                    # Ищем блок "Types of..." (подкатегории на внутренней странице)
                    # Обычно это ссылки в блоке .search-results или .category-browse
                    sub_elements = inner_soup.select('div.search-results a[href*="/c/"], div.category-browse a[href*="/c/"]')
                    
                    if sub_elements:
                        for sub in sub_elements:
                            sub_name = sub.get_text(separator='|', strip=True).split('|')[0].strip()
                            sub_href = sub.get('href', '')
                            sub_url = base_url + sub_href if sub_href.startswith('/') else sub_href
                            
                            if sub_name and sub_name != parent_name:
                                final_data.append({
                                    'Main Category': parent_name,
                                    'Sub Category': sub_name,
                                    'URL': sub_url.split('?')[0]
                                })
                        print(f"   Найдено подкатегорий: {len(sub_elements)}")
                    else:
                        # Если подкатегорий внутри нет
                        final_data.append({
                            'Main Category': parent_name,
                            'Sub Category': '--- Конечная категория ---',
                            'URL': url
                        })

                except Exception as e_inner:
                    print(f"   Ошибка на странице {url}: {e_inner}")
                    continue

                # Сохраняем промежуточный результат каждые 10 страниц
                if (i + 1) % 10 == 0:
                    pd.DataFrame(final_data).to_excel("mouser_temp.xlsx", index=False)

        except Exception as e:
            print(f"Критическая ошибка: {e}")
        
        finally:
            if final_data:
                df = pd.DataFrame(final_data)
                df.to_excel("mouser_categories_final.xlsx", index=False)
                print(f"\nВСЁ! Файл готов: mouser_categories_final.xlsx")
                print(f"Всего строк в таблице: {len(final_data)}")
            
            await context.close()

if __name__ == "__main__":
    asyncio.run(scrape_mouser())