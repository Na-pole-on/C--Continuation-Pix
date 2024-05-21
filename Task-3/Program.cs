using Interopt;
using System.Data;
using System.IO;
using System.Reflection;

DataTable dataTable = new DataTable();
List<string> names = new List<string>();
List<DataRow> dr_isHave = new List<DataRow>();

WorkWithInterrupt interrupt = new WorkWithInterrupt();

string? pathToExcel = interrupt.PathToProject + "In\\Task1.xlsx";
string sheet = "Лист1";

string? pathToDocx = interrupt.PathToProject + "In\\Task.docx";

try
{
    if (pathToExcel != null)
        dataTable = interrupt.GetTable(pathToExcel, sheet);
    else
        throw new Exception("Пути к эксель файлу пустой!");


    if (pathToDocx != null)
        names = interrupt.GetNamesFromDocx(pathToDocx);
    else
        throw new Exception("Путь к ворд файлу пустой");


    if (dataTable.Rows.Count != 0 || names.Count != 0)
        dr_isHave = interrupt.GetDataByName(dataTable, names);
    else
        throw new Exception("Имен в ворде или эксель файл пустые");

    interrupt.DeleteAllTable(pathToDocx, dr_isHave.Count);

    if (dr_isHave.Count > 0)
        interrupt.FillDocx(dr_isHave.CopyToDataTable(), dataTable, pathToDocx);

}
catch(Exception exp)
{
    Console.WriteLine(exp.Message);
}


