using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Importer.Converters
{
    public class CSVConverter : Converter
    {
        public CSVConverter(Price config)
        {
            this.config = config;
            LoadCoef();
        }

        public override List<string[]> LoadData()
        {
            var data = new List<string[]>();
            var strList = config.AmountCol.Split(',');
            var amountCols = new List<int>();
            foreach (var item in strList)
            {
                amountCols.Add(System.Convert.ToInt32(item));
            }

            var table = ConvertCSVtoDataTable(fileName);
           

            for (int i = 0; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];
                if (!config.IsValidRow(row))
                {
                    continue;
                }
                double amount = 0;
                foreach (var item in amountCols)
                {
                    var s = row[item - 1].ToString().Replace('(', ' ').Replace(')', ' ').Replace('>', ' ').Replace(".", ",").Trim();
                    if (s == "") continue;
                    var _amount = 0;
                    Int32.TryParse(s, out _amount);
                    amount += _amount;
                }
                data.Add(new string[] {
                    row[System.Convert.ToInt32(config.CodeCol)-1].ToString(),
                    row[System.Convert.ToInt32(config.NameCol)-1].ToString(),
                    row[System.Convert.ToInt32(config.VendorCol)-1].ToString(),
                    row[System.Convert.ToInt32(config.PriceCol)-1].ToString(),
                    amount.ToString()
                });
            }
            return data;
        }

        protected static Encoding utf8 = Encoding.GetEncoding("utf-8");
        protected static Encoding win1251 = Encoding.GetEncoding("windows-1251");
        private static string Win1251ToUTF8(string source)
        {
            byte[] utf8Bytes = win1251.GetBytes(source);
            byte[] win1251Bytes = Encoding.Convert(win1251, utf8, utf8Bytes);
            source = win1251.GetString(win1251Bytes);
            return source;

        }
        public DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            DataTable dt = new DataTable();
            using (StreamReader sr = new StreamReader(strFilePath, win1251))
            {
                char delimiter = '\t';
                int beginWith = Convert.ToInt32(this.config.BeginWith);
                string[] headers;
                headers = sr.ReadLine().Split(delimiter);
                for (int i = 0; i < beginWith-2; i++)
                {
                    headers = sr.ReadLine().Split(delimiter);
                }
                if (headers.Length == 1)
                {
                    delimiter = ';';
                    headers = headers[0].Split(delimiter);
                }
                foreach (string header in headers)
                {
                    var index = dt.Columns.IndexOf(header);
                    if (index > -1)
                    {
                        dt.Columns.Add(header + index);
                    }
                    else
                    {
                        dt.Columns.Add(header);
                    }
                }

                int length = headers.Length;
                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(delimiter);
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < length; i++)
                    {
                        if (rows.Length < length)
                        {
                            continue;
                        }
                        dr[i] = (rows[i].ToString().Trim('\"'));
                    }
                    dt.Rows.Add(dr);
                }

            }


            return dt;
        }
    }
}
