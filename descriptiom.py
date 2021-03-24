import xlrd
import xlwt
import pymorphy2
morph = pymorphy2.MorphAnalyzer()
path = "C:\\Tagger\\manifest_awb_DO1.xls"
workbook = xlrd.open_workbook(path)
workbookw = xlwt.Workbook()
sheet = workbook.sheet_by_index(0)
sheet1 = workbookw.add_sheet("Sheet1", cell_overwrite_ok=True)
#смотреть комментарий по позициям из двух слов в словаре на линии 86!
dictionary = {"Сумка": ["рюкзак","tote","bag", "handbag", "pouch", "shoulder bag", "backpack"],
"Одежда": ["jumpsuit","jumper","шарф", "футболк", "mask", "jeans", "shorts", "leggings", "scarf", "gloves", "bodysuit", "skirt", "swimsuit", "pant", "sweater", "cardigan", "shirt", "bikini", "trousers", "blouse","hoodie","jacket", "coat", "parka", "dress", "anorak", "blazer", "costume"],
"Обувь":["yeezy500","sandals", "boots", "pumps", "loafers", "mules", "trainers", "footwear", "flipflops", "flip-flops", "flipflops","slipon","slip-on", "low-top","low-topsneakers", "sneakers"],
"Бижутерия": ["brooch","earrings", "necklace", ],
"Очки": ["glasses",],
"Literally": ["оптическиелинзы","парфюмерия","документация","частикомпьютера","бумаг","посуда","персональныйкомпьютер","диппочта", "одежда", "компактдиски","cdдиски","сумка", "ароматизатор", "обувь", "ремень"],
"Посуда": ["тарелк", "миск"],
"Comma_and": [", ", " и ", "&"]}
flip_iteration = 0
sneakers_iteration = 0
rows = sheet.nrows
rowsInWork = 1
print("Описаний грузов в файле:" + str(rows - 1))
FailureAtTagAssignmentProcessCounter = 0
FailureAtTagAssignmentProcessList = []
while (rowsInWork - 1) != (rows - 1):
	comma_counter = 0
	description = sheet.cell_value((rowsInWork), 13).replace(" ", "").lower().split(",")
	print(description)
	raw_str = sheet.cell_value((rowsInWork), 13)
	if "," in raw_str:
		comma_counter += 1
		print("вижу запятую")
	if " и " in raw_str:
		comma_counter += 1
		print("вижу и")
	if "&" in raw_str:
		comma_counter += 1	
		print("вижу энд")
	print("Запятых: " + str(comma_counter))		
	description_for_comma_counter = sheet.cell_value((rowsInWork), 13)
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

	Finded_categories = []
	Bag_counter = 0
	Clothes_counter =0 
	Shoes_counter = 0
	Accesories_counter = 0
	Glasses_counter = 0
	Dishes_counter = 0
	Literally_counter = 0
	StopAtItemCounter = 0

		

	for i in (dictionary["Сумка"]):
		matching = [s for s in description if i in s]
		matchingLen = len(matching)
		if matchingLen != 0:
			print("Обнаружена сумка")
			tagCounter += 1
			Bag_counter += 1
			if (comma_counter + 1) == Bag_counter:
				StopAtItemCounter += 1
				break
			#Строчка на запись в соответствующую ячейку, что это одежда
	for i in (dictionary["Одежда"]):
		matching = [s for s in description if i in s]
		matchingLen = len(matching)
		if matchingLen != 0:
			print("Обнаружена одежда")
			tagCounter += 1
			Clothes_counter += 1
			if (comma_counter + 1) == Clothes_counter:
				StopAtItemCounter += 1
				break
			#Строчка на запись в соответствующую ячейку, что это одежда
	for i in (dictionary["Обувь"]):
		matching = [s for s in description if i in s]
		matchingLen = len(matching)
		if matchingLen != 0:
			print("Обнаружена обувь")
			tagCounter += 1
			Shoes_counter += 1
			if (comma_counter + 1) == Shoes_counter:
				StopAtItemCounter += 1
				break
	for i in (dictionary["Очки"]):
		matching = [s for s in description if i in s]
		matchingLen = len(matching)
		if matchingLen != 0:
			print("Обнаружены очки")
			tagCounter += 1
			Glasses_counter += 1
			if (comma_counter + 1) == Glasses_counter:
				StopAtItemCounter += 1
				break		
			#Строчка на запись в соответствующую ячейку, что это одежда
	for i in (dictionary["Бижутерия"]):
		matching = [s for s in description if i in s]
		matchingLen = len(matching)
		if matchingLen != 0:
			print("Обнаружена бижутерия")
			tagCounter += 1
			Accesories_counter += 1
			if (comma_counter + 1) == Accesories_counter:
				StopAtItemCounter += 1
				break
	for i in (dictionary["Посуда"]):
		matching = [s for s in description if i in s]
		matchingLen = len(matching)
		if matchingLen != 0:
			print("Обнаружена посуда")
			tagCounter += 1
			Dishes_counter += 1
			if (comma_counter + 1) == Dishes_counter:
				break		
	for i in (dictionary["Literally"]):
		matching = [s for s in description if i in s]
		matchingLen = len(matching)
		if matchingLen != 0:
			print("Обнаружено совпадение по литерал")
			tagCounter += 1
			Literally_counter += 1
			if (comma_counter + 1) == Literally_counter:
				StopAtItemCounter += 1
				break							
	if Bag_counter == 1 and (tagCounter - Bag_counter) == 0:
		Finded_categories.append("сумка " + "1 шт.")
	elif Bag_counter == 1 and (tagCounter - Bag_counter) != 0:
		Finded_categories.append("сумка " + str(Bag_counter) + " шт." + ",")	
	elif Bag_counter > 1 and (tagCounter - Bag_counter) == 0:
		Finded_categories.append("сумка " + str(Bag_counter) + " шт.")
	elif Bag_counter > 1 and (tagCounter - Bag_counter) != 0:
		Finded_categories.append("сумка " + str(Bag_counter) + " шт." + ",")
	if Clothes_counter == 1 and (tagCounter - Bag_counter - Clothes_counter) == 0:
		Finded_categories.append("одежда " + "1 шт.")
	elif Clothes_counter == 1 and (tagCounter - Bag_counter - Clothes_counter) != 0:
		Finded_categories.append("одежда " + str(Clothes_counter) + " шт." + ",")
	elif Clothes_counter > 1 and (tagCounter - Bag_counter - Clothes_counter) == 0:
		Finded_categories.append("одежда " + str(Clothes_counter) + " шт.")
	elif Clothes_counter > 1 and (tagCounter - Bag_counter - Clothes_counter) != 0:
		Finded_categories.append("одежда " + str(Clothes_counter) + " шт." + ",")		
	if Shoes_counter == 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter) == 0:
		Finded_categories.append("обувь " + "1 шт.")
	elif Shoes_counter == 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter) != 0:
		Finded_categories.append("обувь " + str(Shoes_counter) + " шт." + ",")
	elif Shoes_counter > 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter) == 0:
		Finded_categories.append("обувь " + str(Shoes_counter) + " шт.")
	elif Shoes_counter > 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter) != 0:
		Finded_categories.append("обувь " + str(Shoes_counter) + " шт." + ",")	
	if Accesories_counter == 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter) == 0:
		Finded_categories.append("бижутерия " + "1 шт.")
	elif Accesories_counter == 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter) != 0:
		Finded_categories.append("бижутерия " + str(Accesories_counter) + " шт." + ",")
	elif Accesories_counter > 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter) == 0:
		Finded_categories.append("бижутерия " + str(Accesories_counter) + " шт.")
	elif Accesories_counter > 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter) != 0:
		Finded_categories.append("бижутерия " + str(Accesories_counter) + " шт." + ",")	
	if Literally_counter == 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter - Literally_counter) == 0:
		Finded_categories.append(sheet.cell_value((rowsInWork), 13))
	elif Literally_counter == 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter - Literally_counter) != 0:
		Finded_categories.append((sheet.cell_value((rowsInWork), 13)) + str(Literally_counter) + " шт." + ",")
	elif Literally_counter > 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter - Literally_counter) == 0:
		Finded_categories.append((sheet.cell_value((rowsInWork), 13)) + str(Literally_counter) + " шт.")
	elif Literally_counter > 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter - Literally_counter) != 0:
		Finded_categories.append((sheet.cell_value((rowsInWork), 13)) + str(Literally_counter) + " шт." + ",")			
#	if Glasses_counter == 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter - Glasses_counter) == 0:
#		Finded_categories.append(str(glasses))
#	elif Accesories_counter == 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter - Glasses_counter) != 0:
#		Finded_categories.append(str(glasses) + ",")
#	elif Accesories_counter > 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter - Glasses_counter) == 0:
#		Finded_categories.append(str(Glasses_counter) + " " + str(glasses))
#	elif Accesories_counter > 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter - Glasses_counter) != 0:
#		Finded_categories.append(str(Glasses_counter) + " " + str(glasses) + ",")		
	if Glasses_counter > 0:
		Finded_categories.append("очки " + str(Glasses_counter) + " шт.")
	if Dishes_counter > 0:
		Finded_categories.append("посуда " + str(Dishes_counter) + " шт.")
	print("Вижу запятые: " + str(comma_counter) + " шт." + "Ожидаю позиций: " + str(comma_counter + 1) + " шт.")

	if (comma_counter + 1) != tagCounter and tagCounter != 0:
		if StopAtItemCounter != 0:
			sheet1.write(rowsInWork, 17, "Внимание! Сценарий \"или-или\"!")
	Finded_categories_str = " ".join(Finded_categories)
	sheet1.write(rowsInWork, 15, str(Finded_categories_str))
	if tagCounter == 0:
		sheet1.write(rowsInWork, 15, "НЕ ОБНАРУЖЕНО КЛЮЧЕЙ!")
		FailureAtTagAssignmentProcessCounter += 1
		FailureAtTagAssignmentProcessList.append((rowsInWork))
	counter = ((rows - 1) - rowsInWork)
	rowsInWork += 1	
	print("Остаётся грузов: " + str(counter))
	print("************************************************")
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
for x in xlist:
	y = 3
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)
for x in xlist:
	y = 4
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)
for x in xlist:
	y = 5
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)
for x in xlist:
	y = 6
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)
for x in xlist:
	y = 7
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)
for x in xlist:
	y = 8
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)
for x in xlist:
	y = 9
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)
for x in xlist:
	y = 10
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)
for x in xlist:
	y = 11
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)
for x in xlist:
	y = 12
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)
for x in xlist:
	y = 13
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)
for x in xlist:
	y = 14
	info = sheet.cell_value(x, y)
	sheet1.write(x, y, info)

workbookw.save("C:\\Tagger\\Taggered__manifest_awb_DO1.xls")

EndingScenario = input()