using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Data;
using System.Collections;
using NPOI.HSSF.UserModel;

namespace Importer.Cross
{
    public class Converter
    {
        protected List<DataTable> sheetTables = new List<DataTable>();

        protected Dictionary<String, Object> crossReplace;

        public List<DataTable> SheetTables { get { return sheetTables; } }

        public String FileName { get; set; }

        public Converter()
        {
            crossReplace = App.Instance.CrossReplace.Rows.OfType<DataRow>().ToDictionary(d => d.Field<string>(0).ToUpper(), v => v.Field<object>(1));
        }

        public void LoadFile(String fileName)
        {
            XSSFWorkbook hssfworkbook;
            FileName = fileName;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new XSSFWorkbook(file);
            }

            var numberOfSheets = hssfworkbook.NumberOfSheets;

            for (int i = 0; i < numberOfSheets; i++)
            {
                LoadSheet(i, hssfworkbook);
            }
        }

        private void LoadSheet(int index, XSSFWorkbook hssfworkbook)
        {
            DataTable table;
            table = new DataTable();
            ISheet sheet = hssfworkbook.GetSheetAt(index);
            IRow headerRow = sheet.GetRow(0);
            IEnumerator rows = sheet.GetRowEnumerator();
            if (rows.MoveNext())
            {
                int colCount = headerRow.LastCellNum;
                for (int c = 0; c < colCount; c++)
                {
                    table.Columns.Add((c + 1).ToString());
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
            }

            sheetTables.Add(table);
        }

        public void Convert(Config config, String fileName)
        {
            StreamWriter csv = new StreamWriter(Stream.Null);

            using (FileStream stream = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                IWorkbook wb = new HSSFWorkbook();
                ISheet sheet = wb.CreateSheet("Sheet1");

                ICreationHelper cH = wb.GetCreationHelper();

                IRow firstRow = sheet.CreateRow(0);

                String[] firstLine = { "brand", "code", "brand_from", "code_from", "load_image", "load_characteristics", "load_cross", "load_applicability" };

                for (int i = 0; i < firstLine.Length; i++)
                {
                    ICell headCell = firstRow.CreateCell(i);
                    headCell.SetCellValue(firstLine[i]);
                }

                var table = SheetTables[config.SheetNumber];

                int k = 1;

                foreach (var row in table.Rows)
                {
                    Object[] arr = ((System.Data.DataRow)row).ItemArray;
                    string codeCol = arr[config.CodeCol - 1].ToString();
                    string brandCol = ReplaceBrand(arr[config.BrandCol - 1].ToString());
                    string destCodeCol = arr[config.DestCodeCol - 1].ToString();
                    string destBrandCol = ReplaceBrand(arr[config.DestBrandCol - 1].ToString());

                    if ((codeCol == "") || (brandCol == "") || (destCodeCol == "") || (destBrandCol == ""))
                    {
                        continue;
                    }

                    String[] line = { codeCol, brandCol, destCodeCol, destBrandCol, "1", "1", "1", "1" };

                    IRow excelRow = sheet.CreateRow(k++);

                    for (int j = 0; j <line.Length; j++)
                    {
                        ICell cell = excelRow.CreateCell(j);
                        cell.SetCellValue(line[j]);
                    }
                }

                wb.Write(stream);

            }


            
            csv.Close();
        }

        private string ReplaceBrand(string brand)
        {
            var source = brand.ToUpper();

            if (crossReplace.ContainsKey(source))
            {
                return (string)crossReplace[source];
            }
            return source;
        }
    }
}
