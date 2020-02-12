import win32com.client #python -m pip install pypiwin32 - вбить в терминале что-бы установить библиотеку
Excel = win32com.client.Dispatch("Excel.Application")

wb = Excel.Workbooks.Open(u'C:\\Users\\USER\\Desktop\\оборудование 90.xlsx') # указать путь к файлу
sheet = wb.ActiveSheet # указать какой лист править

# val = sheet.Cells(3,5).value - в таком виде выглядит получиние данных из ячейки
nac_str = 3     # переменная указывающая начальную строку
con_str = 151   # переменная указывающая конечную строку
con_stl = 9     # переменная указывающая конечный столбец
vivod = 10      # переменная указывающая столбец в который выводим

while nac_str < con_str:                # цикл от начальной строки до конечной
    print(nac_str,"stroki!!!!1!!")

    k = 0                               # подсчет колличества да
    nac_stl = 5
    while nac_stl <= con_stl:           # цикл от начального столбца до конечного столбца
        print(nac_stl,"stolbci")
        if sheet.Cells(nac_str,nac_stl).value == 'да':      # цикл сравнения ячейки с наличием текста да
            print("zashli v uslovie")
            k = k + 1
        sheet.Cells(nac_str, vivod).value = k
        nac_stl = nac_stl + 1
    nac_str = nac_str + 1


# sheet.Cells(3,10).value = "govno"

wb.Save()       # сохранить документ

wb.Close()      # закрыть

Excel.Quit()    # выйти
