
import timeit
code_to_test = """
import docx
doc = docx.Document("test.docx") #открываем документ

list_of_rows = [[[cell.text for cell in row.cells] for row in table.rows] for table in doc.tables] #медленно капец

counter_in_rows = 0

for i in list_of_rows:
    for j in i:
        for text in j:
            counter_in_rows += 1

print(counter_in_rows)

"""
elapsed_time = timeit.timeit(code_to_test, number=1)
print('time in sec:', elapsed_time)
