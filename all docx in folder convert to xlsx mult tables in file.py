def convert_docx_tables(filename): #на вход функция принимает путь к docx файлу в строковом формате, ищет там таблицы, на выходе создаёт xlsx файл с тем же названием и таблицами из docx
    import docx #модуль для работы с docx
    import pandas as pd #модуль для работы с датафреймами
    doc = docx.Document(filename) #создаём документ docx
    list_of_columns = [[[cell.text for cell in column.cells] for column in table.columns] for table in doc.tables] #Загоняем таблицы в список таблиц, в них сохдаём списки столбцов, в них создаём списки ячеек столбцов По какой-то причине проходить по строкам медленнее в 500 раз, так что только по столбцам.
    x = [pd.DataFrame(i) for i in list_of_columns] #на случай если в файле несколько таблиц, создаём список для датафреймов
    x = [i.transpose() for i in x] #Транспортируем датафреймы в списке с датафреймами
    for i in range(len(x)): #в цикле проходим по всем датафреймам, для каждого создаём отдельный файл и присваиваем ему имя оригинала + номер таблицы
        x[i].to_excel(filename + '_table_'+ str(i + 1) + '.xlsx', header=False, index=False, sheet_name='table' + str(i))  # Экспортируем датафрейм без индексов и заголовков
        print(filename + '_table_'+ str(i + 1) + '.xlsx writed')
    print('Done, file' ,(filename + '.xlsx'), 'in folder.')

import os #модуль для работы с операционной системой
current_folder = os.getcwd() #присваивает имя текущей директории
all_in_directory = os.listdir(current_folder) #список со всеми именами папок и файлов, без лишнего
docx_files_in_current_folder = [i for i in all_in_directory if i[-4:] == 'docx'] #фильтруем из всего списка файлов и папок только docx файлы

#а теперь погнали конвертировать все файлы docx в директории.
for doc_file in docx_files_in_current_folder:
    convert_docx_tables(doc_file)
print('Files', *docx_files_in_current_folder, 'converted')

