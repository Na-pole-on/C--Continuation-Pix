using Microsoft.Office.Interop.Word;
using System.Data;
using System.Reflection.Metadata;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

var application = new Microsoft.Office.Interop.Excel.Application();
var workbook = application.Workbooks.Open(@"C:\Users\User\Desktop\Task1.xlsx");
var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["Лист1"];
var range = worksheet.Range[worksheet.Cells[1, 1],
    worksheet.Cells[worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1,
    worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count]];

List<string> names = new List<string>();

System.Data.DataTable dataTable = new();

//Создание колонок
for (int i = 1; i <= range.Columns.Count; i++)
{
    var column = new DataColumn();
    column.ColumnName = "Column" + i;
    dataTable.Columns.Add(column);
}

//Занесение данных в dataTable
for (int i = 1; i <= range.Rows.Count; i++)
{
    var row = dataTable.NewRow();

    for (int j = 1; j <= range.Columns.Count; j++)
    {
        if (range.Cells[i, j] != null)
            row["Column" + j] = (range.Cells[i, j] as Microsoft.Office.Interop.Excel.Range).Value;
    }

    dataTable.Rows.Add(row);
}

application.Workbooks.Close();

var app = new Microsoft.Office.Interop.Word.Application();
var doc = app.Documents.Open(@"C:\Users\User\Desktop\Task.docx");

foreach (Paragraph p in doc.Paragraphs)
    names.Add(p.Range.Text.ToString().Trim());

doc.Close();
app.Quit();


//Удаление пустых колонок
var list_index = dataTable.Columns.Cast<DataColumn>()
    .Select(dc => dataTable.Columns.IndexOf(dc)).Reverse();

list_index.Select(ind =>
{
    bool bool_isEmpty = dataTable.AsEnumerable().All(dr => string.IsNullOrEmpty(dr[ind].ToString()));
    if (bool_isEmpty == true) { dataTable.Columns.RemoveAt(ind); }

    return "";
}).ToList();

var intCol = dataTable.Columns.Cast<DataColumn>().Select(dc => dc.ColumnName).ToList()
    .FindIndex(colName => dataTable.AsEnumerable().Any(row => $"{row[colName]}".Contains("Имя")));

var intRow = dataTable.AsEnumerable().ToList().FindIndex(row => $"{row.ItemArray[intCol]}".Contains("Имя"));

for (int i = intRow - 1; i >= 0; i--)
    dataTable.Rows.RemoveAt(i);

System.Data.DataTable dt_isHave = null;
List<System.Data.DataRow> dr_isHave = dataTable.AsEnumerable()
    .SelectMany(row => names,
                (row, n) => new { dt_name = row, list_name = n })
    .Where(obj => obj.list_name == $"{obj.dt_name[intCol]}")
    .Select(obj => obj.dt_name).Distinct().ToList();

if(dr_isHave.Count > 0)
    dt_isHave = dr_isHave.CopyToDataTable();


if (dt_isHave is not null)
{
    var word = new System.Data.DataTable();

    foreach(string data in dataTable.Rows[0].ItemArray)
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

    app = new Microsoft.Office.Interop.Word.Application();
    doc = app.Documents.Open(@"C:\Users\User\Desktop\Task.docx");

    foreach(Paragraph paragraph in doc.Paragraphs)
    {
        if (word.Rows.Cast<DataRow>().Select(row => row["Имя"].ToString()).Any(str => str == paragraph.Range.Text.Replace("\r", "")))
        {
            Microsoft.Office.Interop.Word.Range newRange = paragraph.Range.Duplicate;

            newRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            Table wordTable = doc.Tables.Add(newRange, 5, 2);

            var table = app.ActiveDocument.Tables[app.ActiveDocument.Tables.Count];
            table.set_Style("Сетка таблицы");

            bool isNameColumn = true;
            int i = 0;

            foreach (Row row in table.Rows)
            {
                isNameColumn = true;

                foreach (Cell cell in row.Cells)
                {
                    if (isNameColumn)
                    {
                        isNameColumn = false;
                        cell.Range.Text = word.Columns[i].ColumnName;
                    }
                    else
                        cell.Range.Text = word.Rows[0][i].ToString();
                }

                i++;
            }
        }
    }

    doc.Save();
    doc.Close();
    app.Quit();

    Console.WriteLine();
}