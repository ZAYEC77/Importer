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

        public override void LoadData()
        {
            var strList = config.AmountCol.Split(',');
            var amountCols = new List<int>();
            foreach (var item in strList)
            {
                amountCols.Add(System.Convert.ToInt32(item));
            }

            var table = ConvertCSVtoDataTable(fileName);

            for (int i = System.Convert.ToInt32(config.BeginWith); i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];
                double amount = 0;
                foreach (var item in amountCols)
                {
                    var s = row[item - 1].ToString().Replace('>', ' ').Replace(".", ",").Trim();
                    if (s == "") continue;
                    var _amount = System.Convert.ToDouble(s);
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
        public static DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            DataTable dt = new DataTable();
            using (StreamReader sr = new StreamReader(strFilePath,win1251))
            {
                char delimiter = '\t';
                string[] headers = sr.ReadLine().Split(delimiter);
                if (headers.Length == 1)
                {
                    delimiter = ';';
                    headers = sr.ReadLine().Split(delimiter);
                }
                foreach (string header in headers)
                {
                    dt.Columns.Add(header);
                }
                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(delimiter);
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        dr[i] = (rows[i].ToString().Trim('\"'));
                    }
                    dt.Rows.Add(dr);
                }

            }


            return dt;
        }
    }
}
