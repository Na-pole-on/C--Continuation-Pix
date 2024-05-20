using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Activities.Custom
{
    internal class DataTableExcel
    {
        private string Path { get; set; }
        private string Sheet { get; set; }

        public DataTableExcel(string path, string sheet)
            => (this.Path, this.Sheet) = (path, sheet);

        public DataTable GetDataTableRange(string range)
        {
            Excel.Application Application = new Excel.Application();
            Excel.Workbook Workbook = Application.Workbooks.Open(Path);
            Excel.Worksheet Worksheet = (Excel.Worksheet)Workbook.Sheets[this.Sheet];
            Excel.Range Range = null;

            int res = 0;

            if (range == "")
                Range = Worksheet.UsedRange;
            else if (int.TryParse(range, out res))
                Range = Worksheet.Range[Worksheet.Cells[res, 1],
                Worksheet.Cells[Worksheet.UsedRange.Row + Worksheet.UsedRange.Rows.Count - 1,
                                Worksheet.UsedRange.Column + Worksheet.UsedRange.Columns.Count - 1]];
            else
                Range = Worksheet.Range[range];


            DataTable dataTable = GetDataTable(Range);

            Application.Workbooks.Close();

            return dataTable;
        }

        private DataTable GetDataTable(Excel.Range range)
        {
            DataTable dataTable = new DataTable();

            for (int i = 1; i <= range.Columns.Count; i++)
            {
                DataColumn column = new DataColumn();
                column.ColumnName = "Column" + i;
                dataTable.Columns.Add(column);
            }

            for (int i = 1; i <= range.Rows.Count; i++)
            {
                DataRow row = dataTable.NewRow();

                for (int j = 1; j <= range.Columns.Count; j++)
                {
                    if (range.Cells[i, j] != null)
                    {
                        row["Column" + j] = (range.Cells[i, j] as Excel.Range).Value;
                    }
                }

                dataTable.Rows.Add(row);
            }

            return dataTable;
        }
    }
}
