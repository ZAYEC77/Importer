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
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using System.Globalization;

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
                using (StreamReader input = new StreamReader(fileName))
                {
                    IWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(input.BaseStream));
                    if (null == workbook)
                    {
                        throw new Exception("Can\'t load excel file");
                    }

                    HSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);

                    Console.WriteLine();

                    foreach (var sheetNumb in sheetNumbs)
                    {
                        table = new DataTable();

                        ISheet sheet = workbook.GetSheetAt(sheetNumb - 1);
                        int beginWith = Convert.ToInt32(this.config.BeginWith);
                        beginWith = (beginWith == 0) ? 0 : (beginWith - 1);
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
                            table.Columns.Add("Col_" + c);
                        }

                        while (rows.MoveNext())
                        {
                            IRow row = (HSSFRow)rows.Current;
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
            }
            else
            {
                XSSFWorkbook hssfworkbook;
                using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    hssfworkbook = new XSSFWorkbook(file);
                    XSSFFormulaEvaluator.EvaluateAllFormulaCells(hssfworkbook);
                }

                foreach (var sheetNumb in sheetNumbs)
                {
                    table = new DataTable();
                    ISheet sheet = hssfworkbook.GetSheetAt(sheetNumb - 1);
                    int beginWith = Convert.ToInt32(this.config.BeginWith);
                    beginWith = (beginWith == 0) ? 0 : (beginWith - 1);
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
                        table.Columns.Add("Col_" + c);
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
                                if (cell.CellType == CellType.Formula)
                                {

                                    DataFormatter dataFormatter = new DataFormatter(CultureInfo.CurrentCulture);
                                    var val = dataFormatter.FormatCellValue(cell);
                                    dr[i] = cell.NumericCellValue.ToString();
                                }
                                else
                                {
                                    dr[i] = cell.ToString();
                                }
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

                if (!config.IsValidRow(row))
                {
                    continue;
                }
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
