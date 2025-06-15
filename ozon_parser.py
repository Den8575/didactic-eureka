import time
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd

def setup_driver():
    options = uc.ChromeOptions()
    options.binary_location = r"C:\Users\usaev\Downloads\GoogleChromePortable\GoogleChromePortable.exe"
    driver = uc.Chrome(options=options)
    return driver

def search_on_ozon(driver, query):
    driver.get('https://www.ozon.ru')
    time.sleep(5)
    search_box = driver.find_element(By.NAME, "text")
    search_box.send_keys(query + Keys.RETURN)
    time.sleep(5)

def scroll_to_end(driver):
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
        time.sleep(2)

def get_sku_and_sizes(driver):
    sku_list = []
    size_list = []
    found = set()
    try:
        sizes = driver.find_elements(By.CSS_SELECTOR, '[data-widget="webCharacteristics"] button, [data-widget="webCharacteristics"] label')
        for s in sizes:
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", s)
                s.click()
                time.sleep(1)
            except: pass
            try:
                sku = driver.find_element(By.CSS_SELECTOR, '[data-widget="webInfoPanel"] [data-sku]').get_attribute('data-sku')
                size = s.text.strip()
                if (sku, size) not in found:
                    sku_list.append(sku)
                    size_list.append(size)
                    found.add((sku, size))
            except:
                continue
    except:
        try:
            sku = driver.find_element(By.CSS_SELECTOR, '[data-widget="webInfoPanel"] [data-sku]').get_attribute('data-sku')
            if (sku, '') not in found:
                sku_list.append(sku)
                size_list.append('')
                found.add((sku, ''))
        except: pass
    if not sku_list:
        try:
            sku = driver.find_element(By.CSS_SELECTOR, '[data-widget="webInfoPanel"] [data-sku]').get_attribute('data-sku')
            if (sku, '') not in found:
                sku_list.append(sku)
                size_list.append('')
        except: pass
    return sku_list, size_list

def extract_card_info(driver, link, parent_link, card_type):
    driver.get(link)
    time.sleep(3)
    data = []
    try:
        title = driver.find_element(By.CSS_SELECTOR, 'h1 span').text
    except:
        title = ''
    try:
        price = driver.find_element(By.CSS_SELECTOR, '[data-widget="webPrice"] span').text
    except:
        price = ''
    try:
        seller = driver.find_element(By.CSS_SELECTOR, '[data-widget="webCurrentSeller"] a span').text
    except:
        seller = ''
    sku_list, size_list = get_sku_and_sizes(driver)
    for sku, size in zip(sku_list, size_list):
        data.append({
            'Название товара': title,
            'Цена': price,
            'Название продавца': seller,
            'SKU': sku,
            'Размер': size,
            'Ссылка на карточку': link,
            'Родительская карточка': parent_link,
            'Примечание': card_type
        })
    return data

def get_attached_links(driver):
    attached_links = []
    try:
        attached_cards = driver.find_elements(By.CSS_SELECTOR, '[data-widget="webAspects"] a')
        for acard in attached_cards:
            href = acard.get_attribute('href')
            if href:
                attached_links.append(href)
    except:
        pass
    return list(set(attached_links))

def main():
    driver = setup_driver()
    print("Браузер запущен.")
    search_on_ozon(driver, "Платье ONG Fashion")
    print("Поиск выполнен.")
    scroll_to_end(driver)
    print("Прокрутка завершена.")

    # Новый универсальный способ найти все ссылки на карточки
    links = []
    elems = driver.find_elements(By.XPATH, '//a[contains(@href, "/product/") and @href]')
    for elem in elems:
        href = elem.get_attribute('href')
        if '/product/' in href and href not in links:
            links.append(href)

    all_data = []
    visited = set()

    for link in links:
        if not link or link in visited:
            continue
        visited.add(link)
        print("Извлекаю карточку:", link)
        all_data += extract_card_info(driver, link, '', 'основная')
        # Парсим прикреплённые карточки
        attached = get_attached_links(driver)
        for alink in attached:
            if alink and alink not in visited:
                visited.add(alink)
                print("Извлекаю прикреплённую карточку:", alink)
                all_data += extract_card_info(driver, alink, link, 'прикреплённая')

    # Сохраняем в Excel
    df = pd.DataFrame(all_data)
    df.to_excel('ozon_dresses_portable.xlsx', index=False)
    print("Результаты сохранены в ozon_dresses_portable.xlsx")
    driver.quit()
    print("Работа завершена успешно.")

if __name__ == "__main__":
    main()
