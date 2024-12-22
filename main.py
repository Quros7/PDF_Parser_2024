# Для считывания PDF
import re
import PyPDF2
# Для анализа структуры PDF и извлечения текста
from pdfminer.high_level import extract_pages, extract_text
from pdfminer.layout import LTTextContainer, LTChar, LTRect, LTFigure
# Для извлечения текста из таблиц в PDF
import pdfplumber
# Для извлечения изображений из PDF
from PIL import Image
from pdf2image import convert_from_path
# Для выполнения OCR, чтобы извлекать тексты из изображений
import pytesseract
# Для удаления дополнительно созданных файлов
import os
# Для удобной и безопасной работы с файлами
from pathlib import Path
# Для обработки аргументов командной строки
import argparse
import pickle
import datetime

# Create the dictionary to extract text from each image
text_per_page = {}

start_content_flag = False
image_flag = False
error_messages = []


# Create function to extract text
def text_extraction(element):
    # Extracting the text from the in line text element
    line_text = element.get_text()

    # Find the formats of the text
    # Initialize the list with all the formats appeared in the line of text
    line_formats = []
    for text_line in element:
        if isinstance(text_line, LTTextContainer):
            # Iterating through each character in the line of text
            for character in text_line:
                if isinstance(character, LTChar):
                    # Append the font name of the character
                    line_formats.append(character.fontname)
                    # Append the font size of the character
                    line_formats.append(character.size)
    # Find the unique font sizes and names in the line
    format_per_line = list(set(line_formats))

    # Return a tuple with the text in each line along with its format
    return line_text, format_per_line


# Extracting tables from the page

def extract_table(pdf_path, page_num, table_num):
    # Open the pdf file
    pdf = pdfplumber.open(pdf_path)
    # Find the examined page
    table_page = pdf.pages[page_num]
    # Extract the appropriate table
    table = table_page.extract_tables()[table_num]

    return table


# Convert table into appropriate format
def table_converter(table):
    table_string = ''
    # Iterate through each row of the table
    for row_num in range(len(table)):
        row = table[row_num]
        # Remove the line breaker from the wrapted texts
        cleaned_row = [
            item.replace('\n', ' ') if item is not None and '\n' in item else 'None' if item is None else item for item
            in row]
        # Convert the table into a string
        table_string += ('|' + '|'.join(cleaned_row) + '|' + '\n')
    # Removing the last line break
    table_string = table_string[:-1]
    return table_string


# Create a function to check if the element is in any tables present in the page
def is_element_inside_any_table(element, page, tables):
    x0, y0up, x1, y1up = element.bbox
    # Change the cordinates because the pdfminer counts from the botton to top of the page
    y0 = page.bbox[3] - y1up
    y1 = page.bbox[3] - y0up
    for table in tables:
        tx0, ty0, tx1, ty1 = table.bbox
        if tx0 <= x0 <= x1 <= tx1 and ty0 <= y0 <= y1 <= ty1:
            return True
    return False


# Function to find the table for a given element
def find_table_for_element(element, page, tables):
    x0, y0up, x1, y1up = element.bbox
    # Change the cordinates because the pdfminer counts from the botton to top of the page
    y0 = page.bbox[3] - y1up
    y1 = page.bbox[3] - y0up
    for i, table in enumerate(tables):
        tx0, ty0, tx1, ty1 = table.bbox
        if tx0 <= x0 <= x1 <= tx1 and ty0 <= y0 <= y1 <= ty1:
            return i  # Return the index of the table
    return None


# Create a function to crop the image elements from PDFs
def crop_image(element, pageObj):
    # Get the coordinates to crop the image from PDF
    [image_left, image_top, image_right, image_bottom] = [element.x0, element.y0, element.x1, element.y1]
    # Crop the page using coordinates (left, bottom, right, top)
    pageObj.mediabox.lower_left = (image_left, image_bottom)
    pageObj.mediabox.upper_right = (image_right, image_top)
    # Save the cropped page to a new PDF
    cropped_pdf_writer = PyPDF2.PdfWriter()
    cropped_pdf_writer.add_page(pageObj)
    # Save the cropped PDF to a new file
    with open('cropped_image.pdf', 'wb') as cropped_pdf_file:
        cropped_pdf_writer.write(cropped_pdf_file)


# Create a function to convert the PDF to images
def convert_to_images(input_file):
    images = convert_from_path(input_file)
    image = images[0]
    output_file = 'PDF_image.png'
    image.save(output_file, 'PNG')


# Create a function to read text from images
def image_to_text(image_path):
    # Read the image
    img = Image.open(image_path)
    # Extract the text from the image
    text = pytesseract.image_to_string(img, lang='rus')
    return text


def scan_pdf(input_path):
    global text_per_page
    global image_flag

    pdf_path = Path(input_path)

    # Create a pdf file object
    pdfFileObj = open(pdf_path, 'rb')
    # Create a pdf reader object
    pdfReaded = PyPDF2.PdfReader(pdfFileObj)

    # Create a boolean variable for image detection
    image_flag = False

    # We extract the pages from the PDF
    for pagenum, page in enumerate(extract_pages(pdf_path)):
        # Initialize the variables needed for the text extraction from the page
        pageObj = pdfReaded.pages[pagenum]
        page_text = []
        line_format = []
        text_from_images = []
        text_from_tables = []
        page_content = []
        # Initialize the number of the examined tables
        table_in_page = -1
        # Open the pdf file
        pdf = pdfplumber.open(pdf_path)
        # Find the examined page
        page_tables = pdf.pages[pagenum]
        # Find the number of tables in the page
        tables = page_tables.find_tables()
        if len(tables) != 0:
            table_in_page = 0

        # Extracting the tables of the page
        for table_num in range(len(tables)):
            # Extract the information of the table
            table = extract_table(pdf_path, pagenum, table_num)
            # Convert the table information in structured string format
            table_string = table_converter(table)
            # Append the table string into a list
            text_from_tables.append(table_string)

        # Find all the elements
        page_elements = [(element.y1, element) for element in page._objs]
        # Sort all the element as they appear in the page
        page_elements.sort(key=lambda a: a[0], reverse=True)

        # Find the elements that composed a page
        for i, component in enumerate(page_elements):
            # Extract the element of the page layout
            element = component[1]

            # Check the elements for tables
            if table_in_page == -1:
                pass
            else:
                if is_element_inside_any_table(element, page, tables):
                    table_found = find_table_for_element(element, page, tables)
                    if table_found == table_in_page and table_found is not None:
                        page_content.append(text_from_tables[table_in_page])
                        page_text.append('table')
                        line_format.append('table')
                        table_in_page += 1
                    # Pass this iteration because the content of this element was extracted from the tables
                    continue

            if not is_element_inside_any_table(element, page, tables):

                # Check if the element is text element
                if isinstance(element, LTTextContainer):
                    # Use the function to extract the text and format for each text element
                    (line_text, format_per_line) = text_extraction(element)
                    # Append the text of each line to the page text
                    page_text.append(line_text)
                    # Append the format for each line containing text
                    line_format.append(format_per_line)
                    page_content.append(line_text)

                # Check the elements for images
                if isinstance(element, LTFigure):
                    # Crop the image from PDF
                    crop_image(element, pageObj)
                    # Convert the croped pdf to image
                    convert_to_images('cropped_image.pdf')
                    # Extract the text from image
                    image_text = image_to_text('PDF_image.png')
                    text_from_images.append(image_text)
                    page_content.append(image_text)
                    # Add a placeholder in the text and format lists
                    page_text.append('image')
                    line_format.append('image')
                    # Update the flag for image detection
                    image_flag = True

        # Create the key of the dictionary
        dctkey = 'Page_' + str(pagenum)
        # Add the list of list as value of the page key
        text_per_page[dctkey] = [page_text, line_format, text_from_images, text_from_tables, page_content]

    # Close the pdf file object
    pdfFileObj.close()

    # Delete the additional files created if image is detected
    # if image_flag:
    #     os.remove('cropped_image.pdf')
    #     os.remove('PDF_image.png')


def parse_pdf(page_key, page_data):
    global start_content_flag
    global error_messages

    dictionary = {'authors_data': None, 'email': None, 'title': None}

    page_num = int(page_key.split("_")[1]) + 1
    page_full = ''.join(page_data[0])
    if not start_content_flag:
        matched = re.search("МАТЕРИАЛЫ ЛЕКЦИЙ ВЕДУЩИХ УЧЕНЫХ", page_full)
        if not matched:
            return
        start_content_flag = True
        return
    if not start_content_flag:
        return
    if '©' not in page_full:
        return

    footer_part = page_full.split('©')[1].strip()
    footer_authors = re.findall(r"\w+?\s+\w\.\w\.,*", "".join(
        footer_part))  # Ищем ФИО внизу первой страницы доклада (с учетом пропущенных запятых)
    footer_authors_part = re.findall(r"\D+\s\w\.\w\.,", "".join(footer_part))
    authors_count = len(footer_authors)

    page_authors = []
    universities = []
    uni_dict = {}
    title_parts = []

    authors_counter = 0
    uni_count = 0

    isNumbered = False
    authors_found = False
    uni_found = False
    title_found = False

    for ind, line in enumerate(page_data[0]):
        line = line.strip()
        if not authors_found:
            if re.search(r"\d", line):
                isNumbered = True
            if isNumbered:
                line_authors = re.findall(r"(\w\.\s\w\.\s\D+)(\d(?:,\s*\d)*)?(?:,\s*)?", line)
            else:
                line_authors = re.findall(r"(?:\w\.\s){2}\w+", line)
            k = len(line_authors)
            if k == 0:
                error_messages.append(
                    f"При обработке страницы ${str(page_num)} возникла ошибка. Проверьте корректность расположения её элеметов:\n"
                    f"- доклад должен начинаться с новой страницы\n"
                    f"- в конце страницы должен содержаться элемент '©' и сразу за ним должен идти список авторов.")
                return "Error!"
            page_authors.extend(a for a in line_authors)
            authors_counter += k
            if authors_counter >= authors_count:
                authors_found = True
                continue
        elif not uni_found:
            x = re.search(r"e-mail", line)
            if x:
                email = re.findall(r"e-mail:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+.[a-zA-Z]{2,})",
                                   line)  # сразу парсим email, если он найден на странице
                dictionary['email'] = email[0]
                # email_ind = ind
                uni_found = True
            if len(line) == 0:
                # title_start = ind + 1
                uni_found = True
                continue
            for font_data in page_data[1][ind]:
                if "Bold" in str(font_data):
                    # title_start = ind
                    uni_found = True
                    break
            if uni_found:
                continue
            line_uni = line.strip()
            universities.append(line_uni)
            uni_count += 1
        elif not title_found:
            if len(line) == 0 and len(title_parts) == 0:
                continue
            elif len(line) == 0:
                title_found = True
                break
            test1 = False
            for font_data in page_data[1][ind]:  # случай, когда нет отступов между title и текстом
                if "Bold" in str(font_data):
                    test1 = True
            if not test1:
                title_found = True
                break
            tmp_title = re.sub(r"\n", " ", line)
            title_parts.append(re.sub(r"\s+", " ", tmp_title).strip())
            continue
    title = " ".join(title_parts).strip()
    dictionary["title"] = title

    multi_institutions = True if not isNumbered and len(
        universities) > 1 else False  # не пронумерованы, и одному автору соответствуют несколько институтов

    if isNumbered:
        tmp_uni = re.findall(r"(\d)(\D+)", "".join(universities))
        uni_dict = {int(num): name.strip() for num, name in tmp_uni}
    else:
        uni_dict = {int(num): name for num, name in enumerate(universities)}

    result = []
    if isNumbered:
        for author, numbers in page_authors:
            numbers = numbers.split(",")
            author_universities = [uni_dict.get(int(num)) for num in numbers]
            result.append((author, author_universities))
    else:
        for author in page_authors:
            author_universities = [uni_dict.get(num) for num in range(uni_count)]
            result.append((author, author_universities))
    dictionary['authors_data'] = result

    return dictionary


def parse_main_title_and_date(text_data, font_data):
    start_ind = -1
    end_ind = -1
    date_ind = -1

    for i in range(len(text_data)):
        test = re.search(r"Материалы", text_data[i])
        if test:
            end_ind = i - 1
    for i in range(end_ind + 1, -1, -1):
        if len(font_data[i]) == 0:
            start_ind = i + 1
            break
    main_title = ""
    for i in range(start_ind, end_ind + 1):
        main_title += text_data[i]
    main_title = re.sub(r"\n", " ", main_title)
    main_title = re.sub(r"\s+", " ", main_title).strip()

    for i in range(end_ind, len(text_data)):
        test = re.search(r".+\sгода", text_data[i])
        if test:
            date_ind = i
    raw_date = text_data[date_ind]
    days_part = raw_date.split(' ')[0]
    dates = re.findall(r"\d{1,2}", days_part)
    dates = [x.strip() for x in dates]
    month_year_part = ' '.join(raw_date.split(' ')[1:])
    date_start = ((dates[0] if len(dates[0]) == 2 else ('0' + dates[0])) + ' ' + month_year_part).strip()
    date_end = ((dates[1] if len(dates[1]) == 2 else ('0' + dates[1])) + ' ' + month_year_part).strip()

    return [main_title, date_start.strip(), date_end.strip()]


def form_thesis_data(main_dict):
    main_title, date_start, date_end = parse_main_title_and_date(text_per_page['Page_0'][0], text_per_page['Page_0'][1])
    main_dict['pdf_title'] = main_title
    main_dict['date_start'] = date_start
    main_dict['date_end'] = date_end
    main_dict['pages'] = {}
    page_counter = 0
    thesis_counter = 0
    for page_key, data in text_per_page.items():
        page_dict = parse_pdf(page_key, data)

        if page_dict is not None:
            thesis_counter += 1
            if page_dict == "Error!":  # обработка ошибки в документе
                continue
            page_dict['page_start_number'] = page_counter + 1
            page_dict['thesis_title'] = page_dict.get('title')
            page_dict['authors_data'] = page_dict.get('authors_data')
            page_dict['email'] = page_dict.get('email')

            thesis_name = 'Thesis_' + str(thesis_counter)
            main_dict['pages'][thesis_name] = page_dict
        page_counter += 1
    main_dict["thesis_count"] = thesis_counter


def make_timestamp():
    now = datetime.datetime.now()
    numbers = re.findall(r"(\d\d)", str(now))
    label = numbers[4] + numbers[5] + "_" + numbers[3] + "_" + numbers[2] + "_" + numbers[0] + numbers[1]
    return label


if __name__ == '__main__':
    try:
        parser = argparse.ArgumentParser(description='Обработка PDF-файлов')
        parser.add_argument('pdf_path', help='Путь к PDF-файлу')
        args = parser.parse_args()

        print("Ожидайте. Примерное время выполнения скрипта = 30sec.\n")
        scan_pdf(args.pdf_path)

        main_dictionary = {}
        form_thesis_data(main_dictionary)
        main_dictionary["errors"] = error_messages

        label = make_timestamp() + ".pickle"
        # Сериализация
        with open(label, 'wb') as f:
            # Pickle the 'data' dictionary using the highest protocol available.
            pickle.dump(main_dictionary, f, pickle.HIGHEST_PROTOCOL)
        print("Имя сериализованного словаря:", label, "\n")

        if len(error_messages) == 0:
            print("Ошибок в процессе выполнения программы не возникло.")
        else:
            print("В процессе выполнения программы возникли ошибки:")
            for er in error_messages:
                print("Error:", er)
        accuracy = round((1 - (len(error_messages) / main_dictionary["thesis_count"])) * 100, 2)
        print(f"\nПриближённая оценка точности извлечения данных = {accuracy}%")
    except Exception as error:
        print("При выполнении программы возникла критическая ошибка:", error)
    finally:
        # Delete the additional files created if image is detected
        if image_flag:
            os.remove('cropped_image.pdf')
            os.remove('PDF_image.png')
