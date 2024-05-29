using Interopt;
using Microsoft.Office.Interop.Word;
using System.Data;

WorkWithInterrupt interrupt = new WorkWithInterrupt();

string? pathToExcel = interrupt.PathToProject + "In\\Task1.xlsx";
string sheet = "Лист1";

string? pathToDocx = interrupt.PathToProject + "In\\Task.docx";

var AppExcel = new Microsoft.Office.Interop.Excel.Application();
var WorkbookExcel = AppExcel.Workbooks.Open(pathToExcel);
var WorksheetExcel = (Microsoft.Office.Interop.Excel.Worksheet)WorkbookExcel.Sheets["Лист1"];
var RangeExcel = WorksheetExcel.Range[WorksheetExcel.Cells[1, 1],
    WorksheetExcel.Cells[WorksheetExcel.UsedRange.Row + WorksheetExcel.UsedRange.Rows.Count - 1,
    WorksheetExcel.UsedRange.Column + WorksheetExcel.UsedRange.Columns.Count]];

Microsoft.Office.Interop.Word.Table TableWord;

var dataTable = new System.Data.DataTable();

//Создание колонок
for (int i = 1; i <= RangeExcel.Columns.Count; i++)
{
    var column = new DataColumn();
    column.ColumnName = "Column" + i;
    dataTable.Columns.Add(column);
}

//Занесение данных в dataTable
for (int i = 1; i <= RangeExcel.Rows.Count; i++)
{
    var row = dataTable.NewRow();

    for (int j = 1; j <= RangeExcel.Columns.Count; j++)
    {
        if (RangeExcel.Cells[i, j] != null)
            row["Column" + j] = (RangeExcel.Cells[i, j] as Microsoft.Office.Interop.Excel.Range).Value;
    }

    dataTable.Rows.Add(row);
}

AppExcel.Workbooks.Close();

//Удаление пустых колонок
var list_index = dataTable.Columns.Cast<DataColumn>()
    .Select(dc => dataTable.Columns.IndexOf(dc)).Reverse();

list_index.Select(ind =>
{
    bool bool_isEmpty = dataTable.AsEnumerable().All(dr => string.IsNullOrEmpty(dr[ind].ToString()));
    if (bool_isEmpty == true) { dataTable.Columns.RemoveAt(ind); }

    return "";
}).ToList();

//Удаление пустых строк
var intCol = dataTable.Columns.Cast<DataColumn>().Select(dc => dc.ColumnName).ToList()
    .FindIndex(colName => dataTable.AsEnumerable().Any(row => $"{row[colName]}".Contains("Имя")));

var intRow = dataTable.AsEnumerable().ToList().FindIndex(row => $"{row.ItemArray[intCol]}".Contains("Имя"));

for (int i = intRow - 1; i >= 0; i--)
    dataTable.Rows.RemoveAt(i);

List<string> names = new List<string>();

var AppWord = new Microsoft.Office.Interop.Word.Application();
var Document = AppWord.Documents.Open(pathToDocx);

foreach (Microsoft.Office.Interop.Word.Paragraph p in Document.Paragraphs)
    names.Add(p.Range.Text.ToString().Trim());

Document.Close();
AppWord.Quit();

var dr_isHave = dataTable.AsEnumerable()
    .SelectMany(row => names,
        (row, n) => new { dt_name = row, list_name = n })
    .Where(obj => obj.list_name == $"{obj.dt_name[intCol]}")
    .Select(obj => obj.dt_name).Distinct().ToList();

AppWord = new Microsoft.Office.Interop.Word.Application();
Document = AppWord.Documents.Open(pathToDocx);

if(Document.Tables.Count > 0)
{
    foreach (Microsoft.Office.Interop.Word.Table table in Document.Tables)
        table.Delete();
}

Document.Save();
Document.Close();
AppWord.Quit();

System.Data.DataTable dt_isHave = null;

if (dr_isHave is not null)
    dt_isHave = dr_isHave.CopyToDataTable();

if (dt_isHave is not null)
{
    var word = new System.Data.DataTable();

    foreach (string data in dataTable.Rows[0].ItemArray)
        word.Columns.Add(data);

    word = dt_isHave.AsEnumerable().Select(dr =>
    {
        var row = word.NewRow();

        row["Имя"] = dr.ItemArray[0];
        row["Фамилия"] = dr.ItemArray[1];
        row["Пол"] = dr.ItemArray[2];
        row["Возраст"] = dr.ItemArray[3];
        row["Доход"] = dr.ItemArray[4];

        word.Rows.Add(row);

        return row;
    }).CopyToDataTable();

    AppWord = new Microsoft.Office.Interop.Word.Application();
    Document = AppWord.Documents.Open(pathToDocx);

    foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in Document.Paragraphs)
    {
        string name = paragraph.Range.Text.Replace("\r", "");

        if (word.Rows.Cast<DataRow>().Select(row => row["Имя"].ToString()).Any(str => str == name))
        {
            //диблируем абзац
            Microsoft.Office.Interop.Word.Range newRange = paragraph.Range.Duplicate;
            //сворачиваем обзац(стираем имеющщийся текс)
            newRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            //всавляем пустую таблицу размером Columns.Count(соответсвует количеству параметров) строк и 2 столбца(первый для имён полонок, второй для значений)
            Table wordTable = Document.Tables.Add(newRange, 5, 2);

            //задаём границы
            var table = AppWord.ActiveDocument.Tables[AppWord.ActiveDocument.Tables.Count];
            table.set_Style("Сетка таблицы");

            bool isNameColumn = true;
            int i = 0;

            var item = word.AsEnumerable()
                .Where(row => row[0].ToString() == name).FirstOrDefault();

            foreach (Microsoft.Office.Interop.Word.Row row in table.Rows)
            {
                isNameColumn = true;

                foreach (Microsoft.Office.Interop.Word.Cell cell in row.Cells)
                {
                    if (isNameColumn)
                    {
                        isNameColumn = false;
                        cell.Range.Text = word.Columns[i].ColumnName;
                    }
                    else
                        cell.Range.Text = item[i].ToString();
                }

                i++;

            }
        }

    }

    Document.Save();
    Document.Close();
    AppWord.Quit();
}

    /*foreach (Paragraph para in wordDoc.Paragraphs)
                    {
                        //Получаем текст
                        string paraText = para.Range.Text.Trim();
                        // для кадой строки 
                        foreach (Microsoft.Office.Interop.Excel.Range row in excelWorksheet.UsedRange.Rows)
                        { 
                            //получаем текст из ячейки ИМЯ
                            Microsoft.Office.Interop.Excel.Range cell = row.Cells[headColumn] as Microsoft.Office.Interop.Excel.Range;
                            // если не пусто, ячейка находится в диапазоне таблицы и имя из дока соответствует имени в ячейку
                            if (cell.Value != null && cell.Row > headRow && paraText.Contains(cell.Value.ToString()))
                            {
                                //проверяем что не было записи на это имя, если было, пропускаем
                                if (names.Contains(cell.Value.ToString()))
                                    continue;
                                //диблируем абзац
                                Microsoft.Office.Interop.Word.Range newRange = para.Range.Duplicate;
                                //сворачиваем обзац(стираем имеющщийся текс)
                                newRange.Collapse(WdCollapseDirection.wdCollapseStart);
                                //всавляем пустую таблицу размером Columns.Count(соответсвует количеству параметров) строк и 2 столбца(первый для имён полонок, второй для значений)
                                Table wordTable = wordDoc.Tables.Add(newRange, excelWorksheet.UsedRange.Columns.Count, 2);
                                //
                                //задаём границы
                                wordTable.Borders.Enable = 1;
                                //заполнение таблицы
                                for (int current_row = 1; current_row <= excelWorksheet.UsedRange.Columns.Count; current_row++)
                                {
                                    wordTable.Cell(current_row, 1).Range.ParagraphFormat.SpaceAfter = 0;
                                    wordTable.Cell(current_row, 1).Range.Text = (excelWorksheet.UsedRange.Cells[headRow, current_row] as Microsoft.Office.Interop.Excel.Range)?.Text.ToString() ?? " ";

                                    wordTable.Cell(current_row, 2).Range.ParagraphFormat.SpaceAfter = 0;
                                    wordTable.Cell(current_row, 2).Range.Text = (excelWorksheet.UsedRange.Cells[row.Row - 1, current_row] as Microsoft.Office.Interop.Excel.Range)?.Text.ToString() ?? " ";
                                }
                                newRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                                //добавляем в спискок имён имя которое занести 
                                names.Add(cell.Value.ToString());
                            }
                        }*/