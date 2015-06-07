﻿using System;
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

            var strList = config.AmountCol.Split(',');
            foreach (var item in strList)
            {
                amountCols.Add(System.Convert.ToInt32(item));
            }

            strList = config.SheetNumber.Split(',');
            foreach (var item in strList)
            {
                sheetNumbs.Add(System.Convert.ToInt32(item));
            }
        }

        public override void LoadData()
        {
            DataTable table;
            bool is2003 = (Path.GetExtension(fileName) == ".xls");

            if (is2003)
            {
                var connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\""; ;
                using (var conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    var sheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT * FROM [" + sheets.Rows[0]["TABLE_NAME"].ToString() + "] ";

                        var adapter = new OleDbDataAdapter(cmd);
                        var ds = new DataSet();
                        adapter.Fill(ds);
                        table = ds.Tables[System.Convert.ToInt32(config.SheetNumber) - 1];
                    }
                    FromTableToData(table);
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
                    IRow headerRow = sheet.GetRow(0);
                    IEnumerator rows = sheet.GetRowEnumerator();

                    int colCount = headerRow.LastCellNum;
                    int rowCount = sheet.LastRowNum;
                    for (int c = 0; c < colCount; c++)
                    {

                        table.Columns.Add(headerRow.GetCell(c).ToString());
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
                    FromTableToData(table);
                }

            }
        }

        private void FromTableToData(DataTable table)
        {
            for (int i = System.Convert.ToInt32(config.BeginWith); i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];
                double amount = 0;
                foreach (var item in amountCols)
                {
                    var s = row[item - 1].ToString().Replace('-', ' ').Replace('+', ' ').Replace('>', ' ').Replace('.', ',').Trim();
                    if (s == "") continue;
                    Double _amount;
                    Double.TryParse(s, out _amount);
                    amount += _amount;
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
