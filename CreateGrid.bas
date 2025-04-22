' Made by https://heybro.ai
' Объявляем переменные
Dim DataFile As String
Dim GridFile As String
Dim DataRange As Range
Dim GridRange As Range

' Указываем путь к файлу с данными
DataFile = "C:\путь\к\вашему\файлу.csv"

' Импорт данных из файла
' Предположим, что файл - CSV с тремя столбцами: X, Y, Z
' В Surfer можно использовать команду Import, например:

Dim ImportSheet As Worksheet
Set ImportSheet = ThisWorkbook.Sheets.Add
ImportSheet.QueryTables.Add Connection:="TEXT;" & DataFile, Destination:=ImportSheet.Range("A1")
With ImportSheet.QueryTables(1)
    .TextFileParseType = xlDelimited
    .TextFileCommaDelimiter = True
    .Refresh
End With

' После импорта данные в диапазоне A1:Cn
Set DataRange = ImportSheet.Range("A1").CurrentRegion

' Создаем сетку (Grid) с помощью метода Kriging
' Настройка параметров интерполяции (может отличаться в зависимости от версии)

Dim Grid As Grid
Set Grid = CreateObject("Surfer.Grid")
Grid.SetSize 100, 100 ' Размер сетки
Grid.SetRange DataRange.Columns(1).Cells, DataRange.Columns(2).Cells ' Пределы по X и Y

' Выполняем интерполяцию методом Kriging
Dim Kriging As Object
Set Kriging = CreateObject("Surfer.Kriging")
Kriging.DataRange = DataRange
Kriging.InterpolationMethod = "Kriging"
Kriging.ZColumn = 3

' Создаем сетку
Dim GridObj As Object
Set GridObj = Kriging.CreateGrid(Grid)

' Сохраняем результат
GridObj.SaveAs "C:\путь\к\сохранению\grid.grd"

MsgBox "Готово! Сетка создана."