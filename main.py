import certifi
import os
import cloudscraper
from bs4 import BeautifulSoup
import openpyxl

# Функция для запроса ссылки и получения HTML контента
def get_html(url):
    scraper = cloudscraper.create_scraper()
    response = scraper.get(url, verify=certifi.where())  # Использование пакета сертификатов
    if response.status_code == 200:
        response.encoding = 'utf-8'
        return response.text
    else:
        print(f"Ошибка при запросе страницы: {response.status_code}")
        return None

# Функция для парсинга HTML и извлечения нужных данных
def parse_html(html):
    soup = BeautifulSoup(html, 'html.parser')
    
    # Извлечение title
    title = soup.title.string if soup.title else 'Нет title'
    
    # Извлечение description
    description_tag = soup.find('meta', attrs={'name': 'description'})
    description = description_tag['content'] if description_tag else 'Нет description'
    
    # Извлечение h1
    h1 = soup.h1.string if soup.h1 else 'Нет h1'
    
    # Извлечение всех <p> тегов
    paragraphs = [p.get_text() for p in soup.find_all('p')]
    
    return title, description, h1, paragraphs

# Функция для записи данных в Excel файл
def write_to_excel(data, filename='output.xlsx'):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    
    # Запись заголовков столбцов
    sheet['A1'] = 'Title'
    sheet['B1'] = 'Description'
    sheet['C1'] = 'H1'
    sheet['D1'] = 'Content'
    
    # Запись данных
    sheet['A2'] = data[0]
    sheet['B2'] = data[1]
    sheet['C2'] = data[2]
    sheet['D2'] = '\n'.join(data[3])
    
    # Получаем текущую директорию и сохраняем файл там
    current_directory = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_directory, filename)
    workbook.save(file_path)
    
    print(f"Данные успешно записаны в {file_path}")

# Основная функция
def main():
    url = input("Введите ссылку на сайт: ")
    html = get_html(url)
    
    if html:
        data = parse_html(html)
        write_to_excel(data)

if __name__ == '__main__':
    main()
