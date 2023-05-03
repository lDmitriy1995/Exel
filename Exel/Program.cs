using Excel = Microsoft.Office.Interop.Excel;

//Объявляем приложение
Excel.Application ex = new Excel.Application();

//Отобразить Excel
ex.Visible = true;

//Количество листов в рабочей книге
ex.SheetsInNewWorkbook = 1;

//Добавить рабочую книгу
Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);

//Отключить отображение окон с сообщениями
ex.DisplayAlerts = false;

//Получаем первый лист документа (счет начинается с 1)
Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);

//Название листа (вкладки снизу)
sheet.Name = "Number";

//Пример заполнения ячеек
for (int i = 1; i <= 3; i++)
{
    for (int j = 1; j <= 3; j++)
    {
        sheet.Cells[i, j] = string.Format("{1}", i, j);
        sheet.Cells[2, 1] = string.Format("{0}", 8);
        sheet.Cells[2, 2] = string.Format("{0}", 9);
        sheet.Cells[2, 3] = string.Format("{0}", 4);
        sheet.Cells[3, 1] = string.Format("{0}", 7);
        sheet.Cells[3, 2] = string.Format("{0}", 6);
        sheet.Cells[3, 3] = string.Format("{0}", 5);

    }
       
}
    
Console.WriteLine("Data uploaded");




