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
        public List<String[]> Data { get { return data; } set { data = value; } }

        protected List<String[]> data = new List<String[]>();

        protected List<Double[]> pricesCoef = new List<Double[]>();

        protected Price config;

        protected string fileName;

        protected void LoadCoef()
        {
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

        public void Convert()
        {
            LoadData();
            SaveFile();
        }

        public void ProcessData()
        {
            List<string[]> newList = new List<string[]>();
            Double rate = System.Convert.ToDouble(config.Rate.Replace('.', ','));
            for (int i = 0; i < data.Count; i++)
            {
                var item = data[i];
                decimal a = 0;
                Decimal.TryParse(item[4], out a);
                if (a == 0) continue;
                var s = item[3].ToString().Replace(".", ",");
                if (s == "") continue;
                item[4] = Math.Round(a).ToString();
                double price;
                Double.TryParse(s, out price);
                price = price * ((100 - config.getSubPercent()) / 100);
                price = price * rate;
                if (price == 0) continue;
                var coef = getEf(price);
                if (UseCoeficients == false)
                {
                    coef = 0;
                }

                price = price + price * coef / 100;
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
                var newItem = new string[]{
                    item[0],
                    item[1],
                    item[2],
                    item[3],
                    item[4],
                    config.ProviderCode
                };

                newList.Add(newItem);
            }
            data = newList;
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

        public static void SaveResultFile(List<string[]> data)
        {
            var sd = new SaveFileDialog() { Filter = "Excel | *.xlsx" };

            if (sd.ShowDialog() == DialogResult.OK)
            {
                using (FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write))
                {
                    IWorkbook wb = new XSSFWorkbook();
                    ISheet sheet = wb.CreateSheet("Sheet1");

                    ICreationHelper cH = wb.GetCreationHelper();

                    IRow firstRow = sheet.CreateRow(0);

                    SetCellValue(firstRow, 0, "ID Категория товара");
                    SetCellValue(firstRow, 1, "Название товара");
                    SetCellValue(firstRow, 2, "Краткое описание");
                    SetCellValue(firstRow, 3, "Полное описание");
                    SetCellValue(firstRow, 4, "Изображение 1 (название файла)");
                    SetCellValue(firstRow, 5, "Изображение 2 (название файла)");
                    SetCellValue(firstRow, 6, "Изображение 3 (название файла)");
                    SetCellValue(firstRow, 7, "Цена");
                    SetCellValue(firstRow, 8, "Валюта (USD,EUR,RUR,UNI)");
                    SetCellValue(firstRow, 9, "Сортировка");
                    SetCellValue(firstRow, 10, "Код поставщика: Привязка товара к прайсу цен");
                    SetCellValue(firstRow, 11, "Артикул привязки: Привязка товара к прайсу цен");
                    SetCellValue(firstRow, 12, "Бренд привязки: Привязка товара к прайсу цен");
                    SetCellValue(firstRow, 13, "SEO: Заголовок");
                    SetCellValue(firstRow, 14, "SEO: Ключевые слова");
                    SetCellValue(firstRow, 15, "SEO: Описание");
                    SetCellValue(firstRow, 16, "Артикул для поиска по артикулу");
                    SetCellValue(firstRow, 17, "Адрес редиректа");
                    SetCellValue(firstRow, 18, "2 - Id фильтра на которого идут характеристики");

                    int k = 1;
                    for (int i = 1; i < data.Count; i++)
                    {
                        var item = data[i];

                        IRow row = sheet.CreateRow(k++);

                        SetCellValue(row, 0, "1517963");
                        SetCellValue(row, 1, item[1]);
                        SetCellValue(row, 2, item[1]);
                        SetCellValue(row, 3, item[1]);
                        SetCellValue(row, 4, "");
                        SetCellValue(row, 5, "");
                        SetCellValue(row, 6, "");
                        SetCellValue(row, 7, item[3]);
                        SetCellValue(row, 8, "UAH");
                        SetCellValue(row, 9, "1");
                        SetCellValue(row, 10, item[5]);
                        SetCellValue(row, 11, item[0]);
                        SetCellValue(row, 12, item[2]);
                        SetCellValue(row, 13, item[1]);
                        SetCellValue(row, 14, "");
                        SetCellValue(row, 15, "");
                        SetCellValue(row, 16, item[0]);
                        SetCellValue(row, 17, "");
                        SetCellValue(row, 18, "");

                    }
                    wb.Write(stream);
                }
            }
        }

        protected static void SetCellValue(IRow row, int index, string value)
        {
            ICell cell = row.CreateCell(index);
            cell.SetCellValue(value);
        }

        private void SaveFile(Boolean onlyCsv = false)
        {

            var sd = new SaveFileDialog() { Filter = "Excel | *.xlsx" };
            Double rate = System.Convert.ToDouble(config.Rate.Replace('.', ','));
            if (sd.ShowDialog() == DialogResult.OK)
            {
                var createCsv = config.FileName != null;
                StreamWriter csv = new StreamWriter(Stream.Null);
                if (createCsv)
                {
                    var csvFile = new FileStream(Directory.GetCurrentDirectory() + "\\" + config.FileName, FileMode.Create, FileAccess.Write);
                    csv = new StreamWriter(csvFile, Encoding.GetEncoding("windows-1251"));
                    csv.WriteLine("Код;Назва;Виробник;Ціна;Кількість");
                }
                using (FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write))
                {
                    IWorkbook wb = new XSSFWorkbook();
                    ISheet sheet = wb.CreateSheet("Sheet1");

                    if (!onlyCsv)
                    {
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
                    }
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
                        double price;
                        Double.TryParse(s, out price);
                        price *= rate;
                        var coef = UseCoeficients ? config.GetExtraByBrand(item[2]) + getEf(price) : 0;
                        price = price * ((100 + coef) / 100);
                        if (price == 0) continue;
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

                        IRow row = sheet.CreateRow(k++);

                        if (createCsv)
                        {
                            csv.WriteLine(String.Join(";", item.Select(x => x.Trim()).ToList()));
                        }

                        if (onlyCsv)
                        {
                            continue;
                        }

                        for (int j = 0; j < 5; j++)
                        {
                            ICell cell = row.CreateCell(j);
                            cell.SetCellValue(item[j]);
                        }
                    }
                    if (createCsv)
                    {
                        csv.Close();
                    }
                    wb.Write(stream);
                }
            }

        }

        public abstract void LoadData();

        public void Upload()
        {
            var sourceFile = Directory.GetCurrentDirectory() + "\\" + config.FileName;
            var uri = App.Instance.Host + "//" + config.FileName;
            FtpWebRequest request;
            FtpWebResponse response;
            NetworkCredential credentials = new NetworkCredential(App.Instance.User, App.Instance.Pass);

            request = (FtpWebRequest)WebRequest.Create(uri);
            request.Credentials = credentials;
            request.Method = WebRequestMethods.Ftp.GetFileSize;

            response = (FtpWebResponse)request.GetResponse();
            if (response.StatusCode != FtpStatusCode.ActionNotTakenFileUnavailable)
            {
                // To delete file
                FtpWebRequest delRequest = (FtpWebRequest)WebRequest.Create(uri);
                delRequest.Credentials = credentials;
                delRequest.Method = WebRequestMethods.Ftp.DeleteFile;
                response = (FtpWebResponse)delRequest.GetResponse();
            }

            request = (FtpWebRequest)WebRequest.Create(uri);
            request.Method = WebRequestMethods.Ftp.UploadFile;

            // This example assumes the FTP site uses anonymous logon.
            request.Credentials = credentials;

            // Copy the contents of the file to the request stream.
            StreamReader sourceStream = new StreamReader(sourceFile);

            byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
            sourceStream.Close();
            request.ContentLength = fileContents.Length;

            Stream requestStream = request.GetRequestStream();
            requestStream.Write(fileContents, 0, fileContents.Length);
            requestStream.Close();

            response = (FtpWebResponse)request.GetResponse();

            response.Close();
        }

        public bool UseCoeficients { get; set; }
    }
}
