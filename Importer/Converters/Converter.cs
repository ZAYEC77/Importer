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
    public static class Extensions
    {
        public static string ToCSV(this DataTable table)
        {
            var result = new StringBuilder();
            for (int i = 0; i < table.Columns.Count; i++)
            {
                result.Append(table.Columns[i].ColumnName);
                result.Append(i == table.Columns.Count - 1 ? "\n" : ";");
            }

            foreach (DataRow row in table.Rows)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    result.Append(row[i].ToString());
                    result.Append(i == table.Columns.Count - 1 ? "\n" : ";");
                }
            }

            return result.ToString();
        }
    }

    public abstract class Converter
    {
        protected List<String[]> data = new List<String[]>();

        protected List<Double[]> pricesCoef = new List<Double[]>();

        protected Price config;

        protected string fileName;

        protected void LoadCoef()
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
        public static Converter CreateConverter(Price config, string fileName)
        {
            switch (config.Format)
            {
                case PriceFormat.Excel:
                    return new ExcelConverter(config) { fileName = fileName };
                case PriceFormat.Csv:
                    return new CSVConverter(config) { fileName = fileName };
                default:
                    break;
            }
            throw new Exception("В конфігурації прайсу не вказано формат вхідного файлу");
        }

        public void Convert()
        {
            LoadData();
            ChangePrice();
            SaveFile();
        }

        private void ChangePrice()
        {
            foreach (var item in data)
            {
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

        private void SaveFile()
        {
            var sd = new SaveFileDialog() { Filter = "Excel | *.xlsx" };
            Double rate = System.Convert.ToDouble(config.Rate);
            if (sd.ShowDialog() == DialogResult.OK)
            {
                var csvFile = new FileStream(Directory.GetCurrentDirectory() + "\\" + config.FileName, FileMode.Create, FileAccess.Write);
                var csv = new StreamWriter(csvFile);
                csv.WriteLine("Код;Назва;Виробник;Ціна;Кількість");
                using (FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write))
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
                        var item = data[i];
                        decimal a = 0;
                        Decimal.TryParse(item[4], out a);
                        if (a == 0) continue;
                        var s = item[3].ToString().Replace(".", ",");
                        if (s == "") continue;
                        item[4] = Math.Round(a).ToString();
                        var price = System.Convert.ToDouble(s);
                        price = price * rate;
                        var coef = getEf(price);
                        item[3] = Math.Round((price + price * coef / 100)).ToString();

                        item[0] = config.RemovePrefix(item[0]);
                        IRow row = sheet.CreateRow(k++);
                        csv.WriteLine(String.Join(";", item.Select(x => x.Trim()).ToList()));
                        for (int j = 0; j < 5; j++)
                        {
                            ICell cell = row.CreateCell(j);
                            cell.SetCellValue(item[j]);
                        }
                    }
                    csv.Close();
                    wb.Write(stream);
                }
            }

        }

        public abstract void LoadData();

        public void Upload()
        {
            var sourceFile = Directory.GetCurrentDirectory() + "\\" + config.FileName;
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(App.Instance.Host);
            request.Method = WebRequestMethods.Ftp.UploadFile;
            request.RenameTo = config.FileName;


            // This example assumes the FTP site uses anonymous logon.
            request.Credentials = new NetworkCredential(App.Instance.User, App.Instance.Pass);

            // Copy the contents of the file to the request stream.
            StreamReader sourceStream = new StreamReader(sourceFile);

            byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
            sourceStream.Close();
            request.ContentLength = fileContents.Length;

            Stream requestStream = request.GetRequestStream();
            requestStream.Write(fileContents, 0, fileContents.Length);
            requestStream.Close();

            FtpWebResponse response = (FtpWebResponse)request.GetResponse();

            response.Close();
        }
    }
}
