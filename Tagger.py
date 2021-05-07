import xlrd
import xlwt
import pymorphy2
def Path_definer():
    global path
    GotAnswer = False
    while GotAnswer == False:
        IQ = input("Файл находится по стандартному пути и имеет стандартное название? \nY\\N\n")
        if IQ.upper() == "Y":
            GotAnswer = True
            path = r'C:\Users\Admin\Downloads\TGR.xls'
            break
        if IQ.upper() == "N":
            GotAnswer = True
            path = r'%s' % (input("Укажите полный путь к файлу:"))
            break
        else:
            print("Пожалуйста, ответьте утвердительно или отрицательно используя буквы \"Y\" или \"N\"!")
Path_definer()
morph = pymorphy2.MorphAnalyzer()
workbook = xlrd.open_workbook( path )
workbookw = xlwt.Workbook()
sheet = workbook.sheet_by_index( 0 )
sheet1 = workbookw.add_sheet( "Sheet1", cell_overwrite_ok=True )
# смотреть комментарий по позициям из двух слов в словаре на линии 86!
dictionary = {"Сумка": ["bag", "bags", "handbag", "pouche", "shoulder bag", "backpack"],
              "Одежда": ["shorts", "coat", "leggings", "scarf", "gloves", "bodysuit", "skirt", "swimsuit", "pants",
                         "pant", "sweater", "cardigan", "cardigans", "sweatshirt", "shirt", "bikini", "bikinis",
                         "trousers", "blouse", "hoodie", "jacket", "coat", "tshirt", "t-shirt", "t-shirts", "t-shirts", "parka", "dress",
                         "anorak", "blazer", "blazers", "costume"],
              "Обувь": ["sandals", "sneakers", "boots", "low-top", "pumps", "loafers", "mules", "trainers", "footwear"],
              "Бижутерия": ["earrings", "necklace" ],
              "Очки": ["sunglasses", "glasses", ]}
Sumka = morph.parse( 'сумка' )[0]
Odezhda = morph.parse( 'одежда' )[0]
Obuv = morph.parse( 'обувь' )[0]
Aksessuary = morph.parse( 'бижутерия' )[0]
Ochki = morph.parse( 'очки ' )[0]
rows = sheet.nrows
rowsInWork = 1
print( "Описаний грузов в файле:" + str( rows - 1 ) )
FailureAtTagAssignmentProcessCounter = 0
FailureAtTagAssignmentProcessList = []
while (rowsInWork - 1) != (rows - 1):
    dictionaryValues = []
    description = sheet.cell_value( (rowsInWork), 2 ).replace( ",", " " ).lower().split()
    print( description )
    descriptionLength = len( description )
    if descriptionLength == 0:
        print( "Нет описания!" + str( rowsInWork ) )
    else:
        print( "Рассматриваю груз." )
    counter = 0
    tagCounter = 0  # Необходим для написания в ячейку "НЕ ОБНАРУЖЕНО КЛЮЧЕЙ!". see line 80
    KeyCounter = 0
    CounterOfIteratedKeys = 0


    def KeyCountDefiner(dictionary):
        return len( dictionary.keys() )


    KeyCounter = KeyCountDefiner( dictionary )
    watchlist = []

    # Назначаем списку watchlist списки-значения словаря и выводим все списки значений словаря следующей компандой:
    while True:
        a = (list( dictionary.values() )[(CounterOfIteratedKeys)])
        watchlist.append( a )
        CounterOfIteratedKeys += 1
        if KeyCounter == CounterOfIteratedKeys:
            break
    # print(range(0, (CounterOfIteratedKeys - 1)))	OUTPUT: range(0, 1)
    # rint(watchlist)	OUTPUT: [['bag', 'handbag', 'pouche', 'shoulder bag'], ['coat', 'tshirt', 't-shirt', 'parka', 'dress', 'anorak', 'blazer', 'costume']]
    # Смотрю, есть ли в списках из значений каждого ключа слово из описания(дескрипшн)

    # print(dictionary["Сумка"]) OUTPUT ['bag', 'handbag', 'pouche', 'shoulder bag']
    Finded_categories = []
    Bag_counter = 0
    Clothes_counter = 0
    Shoes_counter = 0
    Accesories_counter = 0
    Glasses_counter = 0

    for i in description:
        if i in (dictionary["Сумка"]):
            print( "Совпадение. Это — сумка! \n(ключевое слово: " + str( i ) + ")" )
            # Строчка на запись в соответствующую ячейку, что это сумка
            Bag_counter += 1
            tagCounter += 1
    for i in description:
        if i in (dictionary["Одежда"]):
            print( "Совпадение. Это — одежда! \n(ключевое слово: " + str( i ) + ")" )
            Clothes_counter += 1
            tagCounter += 1
    for i in description:
        if i in (dictionary["Обувь"]):
            print( "Совпадение. Это — обувь! \n(ключевое слово: " + str( i ) + ")" )
            Shoes_counter += 1
            tagCounter += 1
    for i in description:
        if i in (dictionary["Бижутерия"]):
            print( "Совпадение. Это — бижутерия! —  \n(ключевое слово: " + str( i ) + ")" )
            Accesories_counter += 1
            tagCounter += 1
    for i in description:
        if i in (dictionary["Очки"]):
            print( "Совпадение. Это — очки! —  \n(ключевое слово: " + str( i ) + ")" )
            Glasses_counter += 1
            tagCounter += 1
    # Состоящим из двух слов ключам необходимо прописывать исключения, как это сделано для "flip flop". "flip flop" записывается в description как 2 слова - flips и flops
    if "flip" and "flops" in description:
        print( "Совпадение. Это — обувь! \n(ключевое слово: " + str( "flips flops" ) + ")" )
        Shoes_counter += 1
        tagCounter += 1
    if "flipflops" in description:
        print( "Совпадение. Это — обувь! \n(ключевое слово: " + str( "flips flops" ) + ")" )
        Shoes_counter += 1
        tagCounter += 1
    if "flip-flops" in description:
        print( "Совпадение. Это — обувь! \n(ключевое слово: " + str( "flips flops" ) + ")" )
        Shoes_counter += 1
        tagCounter += 1
    bag = str( Sumka.make_agree_with_number( Bag_counter ).word )
    clothes = str( Odezhda.make_agree_with_number( Clothes_counter ).word )
    shoes = str( Obuv.make_agree_with_number( Shoes_counter ).word )
    accesories = str( Aksessuary.make_agree_with_number( Accesories_counter ).word )
    glasses = str( Ochki.make_agree_with_number( Glasses_counter ).word )

    # Нужно написать логику определения проставления запятой после написания числа совпадений и наименования
    if Bag_counter == 1 and (tagCounter - Bag_counter) == 0:
        Finded_categories.append( str( bag ) )
    elif Bag_counter == 1 and (tagCounter - Bag_counter) != 0:
        Finded_categories.append( str( bag ) + "," )
    elif Bag_counter > 1 and (tagCounter - Bag_counter) == 0:
        Finded_categories.append( str( Bag_counter ) + " " + str( bag ) )
    elif Bag_counter > 1 and (tagCounter - Bag_counter) != 0:
        Finded_categories.append( str( Bag_counter ) + " " + str( bag ) + "," )
    if Clothes_counter == 1 and (tagCounter - Bag_counter - Clothes_counter) == 0:
        Finded_categories.append( str( clothes ) )
    elif Clothes_counter == 1 and (tagCounter - Bag_counter - Clothes_counter) != 0:
        Finded_categories.append( str( clothes ) + "," )
    elif Clothes_counter > 1 and (tagCounter - Bag_counter - Clothes_counter) == 0:
        Finded_categories.append( str( Clothes_counter ) + " " + str( clothes ) )
    elif Clothes_counter > 1 and (tagCounter - Bag_counter - Clothes_counter) != 0:
        Finded_categories.append( str( Clothes_counter ) + " " + str( clothes ) + "," )
    if Shoes_counter == 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter) == 0:
        Finded_categories.append( str( shoes ) )
    elif Shoes_counter == 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter) != 0:
        Finded_categories.append( str( shoes ) + "," )
    elif Shoes_counter > 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter) == 0:
        Finded_categories.append( str( Shoes_counter ) + " " + str( shoes ) )
    elif Shoes_counter > 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter) != 0:
        Finded_categories.append( str( Shoes_counter ) + " " + str( shoes ) + "," )
    if Accesories_counter == 1 and (
            tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter) == 0:
        Finded_categories.append( str( accesories ) )
    elif Accesories_counter == 1 and (
            tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter) != 0:
        Finded_categories.append( str( accesories ) + "," )
    elif Accesories_counter > 1 and (
            tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter) == 0:
        Finded_categories.append( str( Accesories_counter ) + " " + str( accesories ) )
    elif Accesories_counter > 1 and (
            tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter) != 0:
        Finded_categories.append( str( Accesories_counter ) + " " + str( accesories ) + "," )
    #	if Glasses_counter == 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter - Glasses_counter) == 0:
    #		Finded_categories.append(str(glasses))
    #	elif Accesories_counter == 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter - Glasses_counter) != 0:
    #		Finded_categories.append(str(glasses) + ",")
    #	elif Accesories_counter > 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter - Glasses_counter) == 0:
    #		Finded_categories.append(str(Glasses_counter) + " " + str(glasses))
    #	elif Accesories_counter > 1 and (tagCounter - Bag_counter - Clothes_counter - Shoes_counter - Accesories_counter - Glasses_counter) != 0:
    #		Finded_categories.append(str(Glasses_counter) + " " + str(glasses) + ",")
    if Glasses_counter > 0:
        Finded_categories.append( str( glasses ) + " " + str( Glasses_counter ) + " шт." )

    Finded_categories_str = " ".join( Finded_categories )
    sheet1.write( rowsInWork, 3, str( Finded_categories_str ) )
    if tagCounter == 0:
        sheet1.write( rowsInWork, 3, "НЕ ОБНАРУЖЕНО КЛЮЧЕЙ!" )
        FailureAtTagAssignmentProcessCounter += 1
        FailureAtTagAssignmentProcessList.append( (rowsInWork) )
    counter = ((rows - 1) - rowsInWork)
    rowsInWork += 1
    print( "Остаётся грузов: " + str( counter ) )
    print( "***************************************************************************************" )
    tagCounter = 0
    if rowsInWork == (rows):
        print( ("Готово!").center( 60, "=" ) )
        if FailureAtTagAssignmentProcessCounter == 0:
            print( "Всем грузам присвоены соответствующие тэги.".center( 60 ) )
        if FailureAtTagAssignmentProcessCounter != 0:
            print( "\n\n\n\n\n\nВнимание: найдены ошибки!" )
            print( "Необработанных описаний: " + str(
                FailureAtTagAssignmentProcessCounter ) + " шт." + "Необходимо обработать вручную!" )
            print( "Грузы без тэга: ", end="" )
            print( *FailureAtTagAssignmentProcessList, sep=", " )

nrows = sheet.nrows  # Соответствует заполнению по оси Y
ncols = sheet.ncols  # Соответствует заполнению по оси X
info = ""
xlist = list( range( nrows ) )
ylist = list( range( ncols ) )
for x in xlist:
    y = 0
    info = sheet.cell_value( x, y )
    sheet1.write( x, y, info )
for x in xlist:
    y = 1
    info = sheet.cell_value( x, y )
    sheet1.write( x, y, info )
for x in xlist:
    y = 2
    info = sheet.cell_value( x, y )
    sheet1.write( x, y, info )

workbookw.save(path)

EndingScenario = input("Нажмите Enter для выхода")
