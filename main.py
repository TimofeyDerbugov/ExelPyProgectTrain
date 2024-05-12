import openpyxl

book = openpyxl.open("Project-Management-Sample-Data.xlsx", read_only=True)

sheet = book.active


a = [i for i in range(80, 100)] # числа от 80 до 99, тк 100 не включительное число

# #[<строчка>][<стоблец>]


for prog in range(6, sheet.max_row + 1):

    progres = sheet[prog][7].value
    progress = progres * 100# прогресс в int
    projectName = sheet[prog][1].value# имя проекта
    taskName = sheet[prog][2].value#имя задачи
    asigned = sheet[prog][3].value#имя выполняющего
    startDate = sheet[prog][4].value#начало
    required = sheet[prog][5].value#требуемые дни
    endDate = sheet[prog][6].value#конец



    if progress == 100:
        print("Проект", projectName, "с задачей", taskName, "с началом", startDate, ", который выполнен by", asigned, ", выполнен и сдан за", required,"дней", "в дату:", endDate ," с прогрессом ", progress, "%.")
    elif progress == a:
        print("Проект", projectName, "с задачей", taskName, "с началом", startDate, ", который не до конца выполнен by", asigned, ", должен быть выполнен и сдан за", required, "дней", "в дату:", endDate ,", с прогрессом ", progress, "%", "не успевает в сроки.")
    elif progress == 0:
        print("Проект", projectName, "с задачей", taskName, "с началом", startDate, ", который не  выполнен by", asigned, ", должен быть выполнен и сдан за", required, "дней", "в дату:", endDate ,", с прогрессом ", progress, "%", " даже не начал выполнение.")
    else:
        print("Проект", projectName, "с задачей", taskName, "с началом", startDate, ", который не до конца выполнен by", asigned, ", должен быть выполнен и сдан за", required, "дней", "в дату:", endDate ,", с прогрессом ", progress, "%", "сильно не успевает в сроки.")

