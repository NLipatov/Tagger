import xlrd
import xlwt
path = "C:\\Tagger\\Исходник.xls"
workbook = xlrd.open_workbook(path)
workbookw = xlwt.Workbook()
sheet = workbook.sheet_by_index(0)
sheet1 = workbookw.add_sheet("Sheet1", cell_overwrite_ok=True)
dictionary = {"Сумка": ["bag", "handbag", "pouche", "shoulder bag", "backpack"],"Одежда": ["blouse","hoodie","jacket", "coat", "tshirt", "t-shirt", "parka", "dress", "anorak", "blazer", "costume"],"Обувь":["sandals", "sneakers"],"Аксессуары": ["earrings",] }
BagValues = ["bag", "handbag", "pouche", "shoulder bag", "backpack"]
ClothingValues = ["blouse","hoodie","jacket", "coat", "tshirt", "t-shirt", "parka", "dress", "anorak", "blazer", "costume"]
ShoesValues = ["sandals", "sneakers"]
AccesoriesValues = ["earrings"]
rows = sheet.nrows
rowsInWork = 1
print("Описаний грузов в файле:" + str(rows - 1))
FailureAtTagAssignmentProcessCounter = 0
FailureAtTagAssignmentProcessList = []


while (rowsInWork - 1) != (rows - 1):
	dictionaryValues = []
	description = sheet.cell_value((rowsInWork), 2).replace(",", "").lower().split()
	descriptionLength = len(description)
	if descriptionLength == 0:
		print("Нет описания!" + str(rowsInWork))
	else:
		print("Рассматриваю груз.")	
	counter = 0
	tagCounter = 0 #Необходим для написания в ячейку "НЕ ОБНАРУЖЕНО КЛЮЧЕЙ!". see line 80
	KeyCounter = 0
	CounterOfIteratedKeys = 0
	def KeyCountDefiner(dictionary):
		return len(dictionary.keys())
	KeyCounter = KeyCountDefiner(dictionary)
	watchlist = []
	


	#Назначаем списку watchlist списки-значения словаря и выводим все списки значений словаря следующей компандой:
	while True:
		a = (list(dictionary.values())[(CounterOfIteratedKeys)])
		watchlist.append(a)
		CounterOfIteratedKeys += 1
		if KeyCounter == CounterOfIteratedKeys:
			break		
	#print(range(0, (CounterOfIteratedKeys - 1)))	OUTPUT: range(0, 1)	
	#rint(watchlist)	OUTPUT: [['bag', 'handbag', 'pouche', 'shoulder bag'], ['coat', 'tshirt', 't-shirt', 'parka', 'dress', 'anorak', 'blazer', 'costume']]
	#Смотрю, есть ли в списках из значений каждого ключа слово из описания(дескрипшн)

	#print(dictionary["Сумка"]) OUTPUT ['bag', 'handbag', 'pouche', 'shoulder bag']

	for i in BagValues:
		matching = [s for s in description if i in s]
		matchingLen = len(matching)
		if matchingLen != 0:
			print("Обнаружена сумка")
			sheet1.write(rowsInWork, 6, "сумка")
			tagCounter += 1
			break
			#Строчка на запись в соответствующую ячейку, что это одежда
	for i in ClothingValues:
		matching = [s for s in description if i in s]
		matchingLen = len(matching)
		if matchingLen != 0:
			print("Обнаружена одежда")
			sheet1.write(rowsInWork, 7, "одежда")
			tagCounter += 1
			break
			#Строчка на запись в соответствующую ячейку, что это одежда
	for i in ShoesValues:
		matching = [s for s in description if i in s]
		matchingLen = len(matching)
		if matchingLen != 0:
			print("Обнаружена обувь")
			sheet1.write(rowsInWork, 8, "обувь")
			tagCounter += 1
			break
			#Строчка на запись в соответствующую ячейку, что это одежда
	for i in AccesoriesValues:
		matching = [s for s in description if i in s]
		matchingLen = len(matching)
		if matchingLen != 0:
			print("Обнаружены аксессуары")
			sheet1.write(rowsInWork, 9, "аксессуары")
			tagCounter += 1
			break
			#Строчка на запись в соответствующую ячейку, что это одежда						
	if tagCounter == 0:
		sheet1.write(rowsInWork, 10, "НЕ ОБНАРУЖЕНО КЛЮЧЕЙ!")	
		FailureAtTagAssignmentProcessCounter += 1
		FailureAtTagAssignmentProcessList.append((rowsInWork))
	counter = ((rows - 1) - rowsInWork)
	rowsInWork += 1	

	print("Остаётся грузов: " + str(counter))	
	print("***************************************************************************************")
	tagCounter = 0
	if rowsInWork == (rows):
		print(("Готово!").center(60, "="))
		if FailureAtTagAssignmentProcessCounter == 0:
			print("Всем грузам присвоены соответствующие тэги.".center(60))
		if FailureAtTagAssignmentProcessCounter != 0:
			print("\n\n\n\n\n\nВнимание: найдены ошибки!")
			print("Необработанных описаний: " + str(FailureAtTagAssignmentProcessCounter) + " шт." + "Необходимо обработать вручную!")
			print("Грузы без тэга: ", end="")
			print(*FailureAtTagAssignmentProcessList, sep=", ")
	

nrows = sheet.nrows #Соответствует заполнению по оси Y
ncols = sheet.ncols #Соответствует заполнению по оси X
info = ""
xlist = list(range(nrows))
ylist = list(range(ncols))
for x in xlist:
	y = 0
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)
for x in xlist:
	y = 1
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)
for x in xlist:
	y = 2
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)

workbookw.save("C:\\Tagger\\Taggered.xls")
