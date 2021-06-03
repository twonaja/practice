from shutil import copy
import yaml
import docx
import openpyxl


# функция которая возвращает список индексов у которых в названии
# содержится сигнатура заданая пользователем data_list -> список
# с данными, sign -> сигнатура для поиска
def find_match_indexes(data_list: list, sign: str):
    length_dl = len(data_list)
    index_list = []
    for i in range(0, length_dl):
        if sign in data_list[i]:
            index_list.append(i)
    return index_list


# функция копирования файлов, производит поиск файлов с расширением .docx и .xlsx
# с помощью функции find_match_indexes - возвращает индексы названия новых файлов
# и далее с помощью функции библиотеки shutil -> copy(имя_файла_шаблона, имя_нового_файла)
# производится копирование файлов
def copy_files(example_name: list, dt_list: list, indexes_docx: list, indexes_xlsx):
    # количество документов в списке с расширением .docx
    length_ind_dcx = len(indexes_docx)
    # количество документов в списке с расширением .xlsx
    length_ind_xlx = len(indexes_xlsx)
    for i in range(0, length_ind_dcx):
        # example_name[0] - должно быть название шаблона документа с расширением .docx
        copy(example_name[0], dt_list[indexes_docx[i]])
    for j in range(0, length_ind_xlx):
        # example_name[1] - должно быть название шаблона документа с расширением .xlsx
        copy(example_name[1], dt_list[indexes_xlsx[j]])


# функция используется для замены слов и
# зачистки .docx документов - от оставшихся сигнатур
def docx_replace(doc_name: str, signature: str, repl_data: str):
    word_docx = docx.Document(doc_name)
    # поиск сигнатур в обычном тексте
    sign_list_len = len(signature)
    for paragraph in word_docx.paragraphs:
        # $1 - сигнатура на замену || const1 - вставляемое слово
        paragraph.text = paragraph.text.replace(signature, repl_data)
        # поиск сигнатур в имеющихся таблицах, актуально к примеру для титульных листов
        for table in word_docx.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell.text = cell.text.replace(signature, repl_data)
    word_docx.save(doc_name)


# функция используется для замены слов и
# зачистки .xlsx документов - от оставшихся сигнатур
def excel_replace(doc_name: str, signature: str, repl_data: str):
    excel_doc = openpyxl.load_workbook(doc_name)
    # Лист1 (имя листа в шаблоне не забыть найти более универсальный метод)
    ws = excel_doc["Лист1"]  # открытие листа
    # поиск по всем доступным колонкам и стобцам
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            # получает строку и проверяет ее, если она не пустая пытается выполнить замену
            s = ws.cell(r, c).value
            if s != None and type(s) != int:
                ws.cell(r, c).value = s.replace(signature, repl_data)
    excel_doc.save(doc_name)


if __name__ == '__main__':
    tmpList = []
    print("Start program...\n\n")
    # часть кода для работы с yaml файлом
    with open("script.yaml", 'r', encoding='utf-8') as stream:
        try:
            load = yaml.safe_load(stream)
            # имена шаблонов - тип список, [0] - .docx; [1] - .xlsx.
            examples_name = load['examples_name']
            # сигнатуры
            list_of_signature = load['list_of_signature']
            # имена новых файлов и вставляемые данные - тип список
            names_and_data = load['names_and_data']
        except yaml.YAMLError as exc:
            print(exc)
        # конец работы с yaml файлом

    # indexes_docx - список содержащий индексы имен новых документов с расширением .docx
    ind_docx = find_match_indexes(names_and_data, '.docx')
    # indexes_xlsx - список содержащий индексы имен новых документов с расширением .xlsx
    ind_xlsx = find_match_indexes(names_and_data, '.xlsx')

    #список сигнатур
    ind_beaddoc = find_match_indexes(names_and_data, '0xBEADDOC')
    ind_beadxl = find_match_indexes(names_and_data, '0xBEADXL')
    # копируем файлы
    copy_files(examples_name, names_and_data, ind_docx, ind_xlsx)

    # заменяем сигнатуры
    ind_docx_len = len(ind_docx)
    ind_xlsx_len = len(ind_xlsx)
    for i in range(0, ind_docx_len):
        tmpList = names_and_data[ind_docx[i] + 1: ind_beaddoc[i]] # получаем временный список данных для замены
        num_of_repl_str = len(tmpList)
        for j in range(0, num_of_repl_str):
            # заменяем сигнатуру на нужную строку
            docx_replace(names_and_data[ind_docx[i]], list_of_signature[j], tmpList[j])

    for i in range(0, ind_xlsx_len):
        tmpList = names_and_data[ind_xlsx[i] + 1: ind_beadxl[i]] # получаем временный список данных для замены
        num_of_repl_str = len(tmpList)
        for j in range(0, num_of_repl_str):
            # заменяем сигнатуру на нужную строку
            excel_replace(names_and_data[ind_xlsx[i]], list_of_signature[j], tmpList[j])

    # отчистка файлов от оставшихся сигнатур
    for i in range(0, ind_docx_len):
        len_of_lst_sign = len(list_of_signature)
        for j in range(0, len_of_lst_sign):
            # заменяем сигнатуру на нужную строку
            docx_replace(names_and_data[ind_docx[i]], list_of_signature[j], ' ')
    # отчистка файлов от оставшихся сигнатур
    for i in range(0, ind_xlsx_len):
        len_of_lst_sign = len(list_of_signature)
        for j in range(0, len_of_lst_sign):
            # заменяем сигнатуру на нужную строку
            excel_replace(names_and_data[ind_xlsx[i]], list_of_signature[j], ' ')

    print("End program!\n\n")
    input("Press enter for exit!")
