using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections;

namespace Importer.Converters
{
    public class ExcelConverter : Converter
    {
        List<int> amountCols = new List<int>();
        List<int> sheetNumbs = new List<int>();
        public ExcelConverter(Price config)
        {
            this.config = config;
            LoadCoef();

            if (config.AmountCol != null && config.AmountCol.Length > 0)
            {
                var strAmountList = config.AmountCol.Split(',');
                foreach (var item in strAmountList)
                {
                    amountCols.Add(System.Convert.ToInt32(item));
                }
            }

            var strList = config.SheetNumber.Split(',');
            foreach (var item in strList)
            {
                sheetNumbs.Add(System.Convert.ToInt32(item));
            }
        }

        public override List<string[]> LoadData()
        {
            var data = new List<string[]>();
            DataTable table;
            bool is2003 = (Path.GetExtension(fileName) == ".xls");

            if (is2003)
            {
                var connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\""; ;
                var conn = new OleDbConnection(connectionString);
                conn.Open();
                using (conn)
                {

                    var sheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT * FROM [" + sheets.Rows[0]["TABLE_NAME"].ToString() + "] ";

                        var adapter = new OleDbDataAdapter(cmd);
                        var ds = new DataSet();
                        adapter.Fill(ds);
                        table = ds.Tables[System.Convert.ToInt32(config.SheetNumber) - 1];
                    }
                    FromTableToData(data, table);
                }
            }
            else
            {
                XSSFWorkbook hssfworkbook;
                using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    hssfworkbook = new XSSFWorkbook(file);
                }

                foreach (var sheetNumb in sheetNumbs)
                {
                    table = new DataTable();
                    ISheet sheet = hssfworkbook.GetSheetAt(sheetNumb - 1);
                    int beginWith = Convert.ToInt32(this.config.BeginWith);
                    IRow headerRow = sheet.GetRow(beginWith);
                    if (headerRow.Cells.Count == 0)
                    {
                        headerRow = sheet.GetRow(beginWith + 1);
                    }
                    IEnumerator rows = sheet.GetRowEnumerator();

                    int colCount = headerRow.LastCellNum;
                    int rowCount = sheet.LastRowNum;
                    for (int c = 0; c < colCount; c++)
                    {
                        var cell = headerRow.GetCell(c);
                        if (cell == null)
                        {
                            table.Columns.Add("Col_" + c);
                        }
                        else
                        {
                            table.Columns.Add(cell.ToString());
                        }
                    }

                    while (rows.MoveNext())
                    {
                        IRow row = (XSSFRow)rows.Current;
                        DataRow dr = table.NewRow();

                        for (int i = 0; i < colCount; i++)
                        {
                            ICell cell = row.GetCell(i);

                            if (cell != null)
                            {
                                dr[i] = cell.ToString();
                            }
                        }
                        table.Rows.Add(dr);
                    }
                    FromTableToData(data, table);
                }
            }
            return data;
        }

        private void FromTableToData(List<string[]> data, DataTable table)
        {
            int defaultAmount;
            int.TryParse(config.DefaultAmount, out defaultAmount);
            defaultAmount = (defaultAmount > 0) ? defaultAmount : 1;

            for (int i = System.Convert.ToInt32(config.BeginWith); i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];
                double amount = 0;
                if (amountCols.Count == 0)
                {
                    amount = defaultAmount;
                }
                else
                {
                    foreach (var item in amountCols)
                    {
                        Double _amount;
                        if (row[item - 1].ToString().Trim() == "есть")
                        {
                            amount = defaultAmount;
                            break;
                        }
                        var s = row[item - 1].ToString().Replace('-', ' ').Replace('+', ' ').Replace('>', ' ').Replace('.', ',').Trim();
                        if (s == "") continue;
                        Double.TryParse(s, out _amount);
                        amount += _amount;
                    }
                }
                if (amount > 0)
                {
                    data.Add(new string[] {
                    row[System.Convert.ToInt32(config.CodeCol)-1].ToString(),
                    row[System.Convert.ToInt32(config.NameCol)-1].ToString(),
                    row[System.Convert.ToInt32(config.VendorCol)-1].ToString(),
                    row[System.Convert.ToInt32(config.PriceCol)-1].ToString(),
                    amount.ToString()
                });
                }
            }
        }

    }
}
