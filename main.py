import os
import requests
from bs4 import BeautifulSoup as BeaS
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
from html4docx import HtmlToDocx


BOOKS_DIR = 'C:\\Users\\Mick\\Books'
HOST = 'http://readli.net'


def create_dir_for_book(name: ''):
    # Вказуємо шлях до каталогу
    full_path = os.path.join(BOOKS_DIR, name)  # Використовуємо os.path.join для кросплатформності
    if not os.path.exists(full_path):
        os.makedirs(full_path)      # Створюємо директорію, включаючи всі проміжні
        print(f"Directory {full_path} was created.")
    else:
        print(f"Directory {full_path} not empty.")
    return full_path


def file_upload(pathname, filename, info):
    try:
        # Зберігаємо в файл або txt або html
        file_path = os.path.join(pathname, filename)
        with open(file_path, "w", encoding='utf-8') as handle:
            handle.write(info)
        print(f"File '{filename}' saved in '{pathname}'.")
    except Exception as e:
        print(f"An error occurred: {e}")


def file_to_docx(pathname, filename, info):
    doc_handler = Document()
    # Завдаємо параметри основого тексту
    doc_handler.styles['Normal'].font.bold = False
    style = doc_handler.styles['Normal']
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0, 0, 0)
    # Завдаємо параметри заголовки
    style_h1 = doc_handler.styles['Heading 1']
    style_h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    style_h1.font.underline = WD_UNDERLINE.SINGLE
    font_h1 = style_h1.font
    font_h1.name = 'Arial'
    font_h1.size = Pt(18)
    font_h1.color.rgb = RGBColor(0, 0, 0)
    style_h2 = doc_handler.styles['Heading 2']
    style_h2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    # style_h2.font.underline = WD_UNDERLINE.SINGLE
    font_h2 = style_h2.font
    font_h2.name = 'Arial'
    font_h2.size = Pt(16)
    font_h2.color.rgb = RGBColor(0, 0, 0)
    doc_handler.styles['Heading 3'].font.bold = False
    style_h3 = doc_handler.styles['Heading 3']
    style_h3.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    # style_h3.font.underline = WD_UNDERLINE.DOTTED
    font_h3 = style_h3.font
    font_h3.name = 'Calibri'
    font_h3.size = Pt(12)
    font_h3.color.rgb = RGBColor(0, 0, 0)

    html_parser = HtmlToDocx()
    file_path = os.path.join(pathname, filename)
    try:
        html_parser.add_html_to_document(info, doc_handler)
        # for paragraph in doc_handler.paragraphs:
        #    for run in paragraph.runs:
        #        if run.font.bold:
        #            run.font.bold = False
    except Exception as e:
        print(f"An error with Docx occurred {e}")
    doc_handler.save(file_path)
    print(f"File '{filename}' saved in '{pathname}'.")


def book_downloader(link, author_name, book_name_):
    text = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>{book_name_}</title>
    </head>
    <body>
        <h2>{book_name_}</h2>
        <h3>{author_name}</h3>
        <p>"""

    text_end = """</p>
    </body>
    </html>
    """
    with requests.Session() as session:
        try:
            response = session.get(link)
            response.raise_for_status()         # Перевіряємо на статус 200
            soup = BeaS(response.text, 'html.parser')

            # Визначаємо номер останньої сторінки для сбору тексту
            link = soup.find('a', class_='page-nav-2__button button-next')['href']
            last_page_nummer = link.split('=')[-1]

            # Визначаємо номер наступнної сторінки для сбору тексту
            next_page_link = soup.find('a', class_='page-nav-1__button')['href']
            if not next_page_link.startswith('http'):
                next_page_link = HOST + next_page_link

            for i in range(1, int(last_page_nummer)+1):

                try:
                    response = session.get(next_page_link)
                    response.raise_for_status()      # Перевіряємо на статус 200
                    soup = BeaS(response.text, 'html.parser')

                    # Отримуємо вміст сторінки
                    reading_text_div = soup.find('div', class_='reading__text')
                    if reading_text_div:
                        # Видаляємо небажані теги, такі як <script>
                        for script in reading_text_div(['script', 'div']):
                            script.decompose()  # Видаляє теги з дерева
                        for br in reading_text_div.find_all('br'):
                            br.replace_with('[BR_TAG]')
                        for h1 in reading_text_div.find_all('h1'):
                            h1.insert(0, '[H1_TAG]')
                            h1.append('[/H1_TAG]')
                        for h2 in reading_text_div.find_all('h2'):
                            h2.insert(0, '[H2_TAG]')
                            h2.append('[/H2_TAG]')
                        for h3 in reading_text_div.find_all('h3'):
                            h3.insert(0, '[H3_TAG]')
                            h3.append('[/H3_TAG]')
                        text += reading_text_div.get_text(separator='\n', strip=True) + '\n'
                    else:
                        print("Teg <div clas='reading__text> was not find.")

                    # Визначаємо URL наступнної сторінки для сбору тексту
                    next_page_link = soup.find('a', class_='page-nav-1__button')['href'].split('&')[0]
                    if not next_page_link.startswith('http'):
                        next_page_link = HOST + next_page_link + f'&pg={i}'

                except Exception as e:
                    print(f'An error occurred: {e} with page {next_page_link}')

        except Exception as e:
            print(f'An error occurred: {e}')
        text = text.replace('[BR_TAG]', '<br>')
        text = text.replace('[P_TAG]', '<p>').replace('[/P_TAG]', '</p>')
        text = text.replace('[H1_TAG]', '<h1>').replace('[/H1_TAG]', '</h1>')
        text = text.replace('[H2_TAG]', '<h2>').replace('[/H2_TAG]', '</h2>')
        text = text.replace('[H3_TAG]', '<h3>').replace('[/H3_TAG]', '</h3>')

    return text+text_end


def step_one(root_url):
    author_name = ''
    title_text = ''
    try:
        session = requests.Session()
        response = session.get(root_url)
        soup = BeaS(response.text, 'html.parser')
        # Перехід на початкову сторінку та шукаємо кнопку для переходу до книги

        link = soup.find('a', class_='book-actions__button')
        next_url_page = link['href']

        author_link = soup.find('a', class_='main-info__link')
        if author_link:
            author_name = author_link.text.strip()

        new_author_name = author_name.split(' ')[-1] + ' ' + author_name.split(' ')[-2]
        # print(f"Author: {new_author_name}")

        if not next_url_page.startswith('http'):
            next_url_page = HOST + next_url_page

        title_tag = soup.find('title')
        if title_tag:
            title_text = title_tag.get_text(strip=True).replace(f'{new_author_name} ', '')
            title_text = title_text.replace(' скачать книгу fb2 txt бесплатно,', '')
            title_text = title_text.replace(' читать текст онлайн, отзывы', '')
            # print(f"Title: {title_text}")
        else:
            print("Title tag not found.")

        author_link = soup.find('a', class_='main-info__link')
        if author_link:
            author_name = author_link.text.strip()
        # Переходимо до вмісту книги для скачування її тексту, що буде
        # повернено після виконнаня функції.
        text = book_downloader(next_url_page, author_name, title_text)
        session.close()
    except Exception as e:
        print(f"An error occurred: {e}")
        return
    return author_name, text


if __name__ == '__main__':
    if len(sys.argv) > 1:
        book_url = sys.argv[1]
        book_name = book_url.strip('/').split('/')[-1] + '.html'
    else:
        print("Missing data (URL) in the command line.")
        exit(0)

    author, book = step_one(book_url)

    full_book_name = create_dir_for_book(author)
    file_upload(full_book_name, book_name, book)

    book_name = book_url.strip('/').split('/')[-1] + '.txt'
    file_upload(full_book_name, book_name, book)

    book_name = book_url.strip('/').split('/')[-1] + '.docx'
    file_to_docx(full_book_name, book_name, book)
