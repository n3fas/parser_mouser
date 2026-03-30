import asyncio
import random
import pandas as pd
from playwright.async_api import async_playwright
from bs4 import BeautifulSoup

async def scrape_level_2():
    # 1. Загружаем твои 349 ссылок
    try:
        df_input = pd.read_excel("mouser_categories_final.xlsx")
        # Убедимся, что берем только уникальные URL, чтобы не делать лишнюю работу
        urls_to_check = df_input[['Main Category', 'URL']].drop_duplicates().values.tolist()
        print(f"Загружено {len(urls_to_check)} ссылок для анализа.")
    except Exception as e:
        print(f"Не удалось прочитать Excel: {e}")
        return

    async with async_playwright() as p:
        user_data_dir = "./mouser_profile_level2"
        context = await p.chromium.launch_persistent_context(
            user_data_dir,
            headless=False,
            args=["--disable-blink-features=AutomationControlled"]
        )
        
        page = await context.new_page()
        # Ручной стелс
        await page.add_init_script("Object.defineProperty(navigator, 'webdriver', { get: () => false });")

        base_url = "https://eu.mouser.com"
        level_2_results = []

        for i, (parent_name, url) in enumerate(urls_to_check):
            print(f"[{i+1}/{len(urls_to_check)}] Проверяем вложенность для: {parent_name}")
            
            try:
                # Пауза 3-6 секунд, чтобы не злить Akamai
                await page.wait_for_timeout(random.randint(3000, 6000))
                await page.goto(url, wait_until="domcontentloaded")
                
                # Ждем появления подкатегорий (lnkCategory_)
                # Если их нет в течение 5 секунд, значит это конечный уровень
                try:
                    await page.wait_for_selector('a[id^="lnkCategory_"]', timeout=5000)
                except:
                    pass

                soup = BeautifulSoup(await page.content(), 'html.parser')
                # Ищем ссылки lnkCategory_ на этой странице
                sub_tags = soup.find_all('a', id=lambda x: x and x.startswith('lnkCategory_'))

                if sub_tags:
                    for a in sub_tags:
                        sub_name = a.get_text(separator='|', strip=True).split('|')[0].strip()
                        sub_href = a.get('href', '')
                        if sub_href:
                            full_sub_url = (base_url + sub_href if sub_href.startswith('/') else sub_href).split('?')[0]
                            
                            level_2_results.append({
                                'Level 1 Category': parent_name,
                                'Level 1 URL': url,
                                'Level 2 Category': sub_name,
                                'Level 2 URL': full_sub_url
                            })
                    print(f"   Найдено подкатегорий: {len(sub_tags)}")
                else:
                    # Если вложенности нет, помечаем как финал
                    level_2_results.append({
                        'Level 1 Category': parent_name,
                        'Level 1 URL': url,
                        'Level 2 Category': '--- FINAL LEVEL ---',
                        'Level 2 URL': url
                    })

            except Exception as e:
                print(f"   Ошибка при переходе на {url}: {e}")
                continue

            # Промежуточное сохранение каждые 20 ссылок
            if (i + 1) % 20 == 0:
                pd.DataFrame(level_2_results).to_excel("mouser_level_2_temp.xlsx", index=False)

        # Финальное сохранение
        final_df = pd.DataFrame(level_2_results)
        final_df.to_excel("mouser_level_2_complete.xlsx", index=False)
        print(f"\nГотово! Результаты в файле: mouser_level_2_complete.xlsx")
        
        await context.close()

if __name__ == "__main__":
    asyncio.run(scrape_level_2())