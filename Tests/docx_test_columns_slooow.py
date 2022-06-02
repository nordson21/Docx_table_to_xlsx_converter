
import timeit
code_to_test = """
import docx
doc = docx.Document("test.docx") #открываем документ

list_of_columns = [[[cell.text for cell in column.cells] for column in table.columns] for table in doc.tables] #почему то быстро

counter_in_columns = 0

for i in list_of_columns:
    for j in i:
        for text in j:
            counter_in_columns += 1

print(counter_in_columns)

"""
elapsed_time = timeit.timeit(code_to_test, number=1)
print('time in sec:', elapsed_time)
