using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Net;

namespace Importer.Converters
{
    public abstract class Converter
    {
        protected List<Double[]> pricesCoef = new List<Double[]>();

        protected Price config;

        protected string fileName;
        public double Rate { get { return rate; } }

        protected double rate;

        protected void LoadCoef()
        {
            rate = System.Convert.ToDouble(config.Rate.Replace('.', ','));
            if (config.Coeficients != null)
            {
                var coefStrings = config.Coeficients.Split('\n');
                foreach (var item in coefStrings)
                {
                    var s = item.Split(' ');
                    if (s.Length != 3) continue;
                    var begin = System.Convert.ToDouble(s[0]);
                    var end = System.Convert.ToDouble(s[1]);
                    var percent = System.Convert.ToDouble(s[2]);
                    pricesCoef.Add(new Double[] { begin, end, percent });
                }
            }
        }

        public static Converter CreateConverter(Price config, string fileName)
        {
            switch (config.Format)
            {
                case PriceFormat.Excel:
                    return new ExcelConverter(config) { fileName = fileName };
                case PriceFormat.Csv:
                    return new CSVConverter(config) { fileName = fileName };
                case PriceFormat.DBF:
                    return new DBFConverter(config) { fileName = fileName };
                default:
                    break;
            }
            throw new Exception("В конфігурації прайсу не вказано формат вхідного файлу");
        }

        public void ProcessData(List<string[]> data)
        {
            List<string[]> newList = new List<string[]>();
            for (int i = 0; i < data.Count; i++)
            {
                var item = ProcessItem(data[i]);
                if (item == null)
                {
                    continue;
                }
                var newItem = new string[]{
                    item[0],
                    item[1],
                    item[2],
                    item[3],
                    item[4]
                };

                newList.Add(newItem);
            }
            data = newList;
        }

        public string[] ProcessItem(string[] data)
        {
            Double rate = System.Convert.ToDouble(config.Rate.Replace('.', ','));
            var item = data;
            decimal a = 0;
            Decimal.TryParse(item[4], out a);
            if (a == 0) return null;
            var s = item[3].ToString().Replace(".", ",").Replace("EUR","").Trim();
            if (s == "") return null;
            item[4] = Math.Round(a).ToString();
            double price;
            Double.TryParse(s, out price);
            price *= this.Rate;
            var coef = UseCoeficients ? config.GetExtraByBrand(item[2]) + getEf(price) : 0;
            price = price * ((100 + coef) / 100);
            if (price == 0) return null;
            if (config.RoundPrice)
            {
                item[3] = Math.Round(price).ToString();
            }
            else
            {
                item[3] = price.ToString();
            }

            if (config.ReplaceBrand)
            {
                item[2] = config.NewBrand;
            }

            item[0] = config.RemovePrefix(item[0]);

            return item;
        }


        public void SaveCsv(List<String[]> data, String path)
        {
            if (!File.Exists(path))
            {
                path = Path.Combine(path, config.FileName);
            }
            path = Path.ChangeExtension(path, ".txt");
            StreamWriter csv = new StreamWriter(Stream.Null);
            var csvFile = new FileStream(path, FileMode.Create, FileAccess.Write);
            csv = new StreamWriter(csvFile, Encoding.GetEncoding("windows-1251"));
            csv.WriteLine("Код;Назва;Виробник;Ціна;Кількість");

            for (int i = 1; i < data.Count; i++)
            {
                var item = ProcessItem(data[i]);
                if (item == null)
                {
                    continue;
                }

                csv.WriteLine(String.Join(";", item.Select(x => x.Trim()).ToList()));
            }

            csv.Close();
        }

        public void SaveExcel(List<String[]> data, String path)
        {
            if (!File.Exists(path))
            {
                path = Path.Combine(path, config.FileName);
            }
            var excelPath = Path.ChangeExtension(path, ".xlsx");
            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook wb = new XSSFWorkbook();
                ISheet sheet = wb.CreateSheet("Sheet1");

                ICreationHelper cH = wb.GetCreationHelper();

                IRow firstRow = sheet.CreateRow(0);
                ICell headCell = firstRow.CreateCell(0);
                headCell.SetCellValue("Код");
                headCell = firstRow.CreateCell(1);
                headCell.SetCellValue("Назва");
                headCell = firstRow.CreateCell(2);
                headCell.SetCellValue("Виробник");
                headCell = firstRow.CreateCell(3);
                headCell.SetCellValue("Ціна");
                headCell = firstRow.CreateCell(4);
                headCell.SetCellValue("Кількість");

                int k = 1;
                for (int i = 1; i < data.Count; i++)
                {
                    var item = ProcessItem(data[i]);
                    if (item == null)
                    {
                        continue;
                    }

                    IRow row = sheet.CreateRow(k++);

                    for (int j = 0; j < 5; j++)
                    {
                        ICell cell = row.CreateCell(j);
                        cell.SetCellValue(item[j]);
                    }
                }
                wb.Write(stream);

            }
        }
        private double getEf(Double p)
        {
            var coef = 0d;
            foreach (var pi in pricesCoef)
            {
                if (pi[0] <= p)
                {
                    coef = pi[2];
                }
            }
            return coef;
        }

        protected static void SetCellValue(IRow row, int index, string value)
        {
            ICell cell = row.CreateCell(index);
            cell.SetCellValue(value);
        }

        public abstract List<String[]> LoadData();

        public bool UseCoeficients { get; set; }
    }
}
