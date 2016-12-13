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
            var csvFile = new FileStream(fileName, FileMode.Create, FileAccess.Write);
            csv = new StreamWriter(csvFile, Encoding.GetEncoding("windows-1251"));

            var table = SheetTables[config.SheetNumber];

            foreach (var row in table.Rows)
            {
                Object[] arr = ((System.Data.DataRow)row).ItemArray;
                string codeCol = arr[config.CodeCol - 1].ToString();
                string brandCol = arr[config.BrandCol - 1].ToString();
                string nameCol = arr[config.NameCol - 1].ToString();
                string destCodeCol = arr[config.DestCodeCol - 1].ToString();
                string destBrandCol = ReplaceBrand(arr[config.DestBrandCol - 1].ToString());


                if ((codeCol == "") || (brandCol == "") || (nameCol == "") || (destCodeCol == "") || (destBrandCol == ""))
                {
                    continue;
                }

                String[] line = { codeCol, brandCol, nameCol, destCodeCol, destBrandCol };
                csv.WriteLine(String.Join(";", line));
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
