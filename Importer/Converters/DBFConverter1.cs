using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Odbc;
using System.Data;
using System.IO;

namespace Importer.Converters
{
    class DBFConverter : Converter
    {

        private OdbcConnection Conn = null;

        public override void LoadData()
        {
            var command = "SELECT * FROM tmp";

            var localFile = App.GetHomeDir() + "tmp.dbf";
            File.Copy(fileName, localFile);

            var dir = Path.GetDirectoryName(fileName);

            this.Conn = new System.Data.Odbc.OdbcConnection();
            Conn.ConnectionString = @"Driver={Microsoft dBase Driver (*.dbf)};datasource=" + dir;
            DataTable dt = null;
            if (Conn != null)
            {
                Conn.Open();
                dt = new DataTable();
                System.Data.Odbc.OdbcCommand oCmd = Conn.CreateCommand();
                oCmd.CommandText = command;
                dt.Load(oCmd.ExecuteReader());
                Conn.Close();
            }

            var strList = config.AmountCol.Split(',');
            var amountCols = new List<int>();
            foreach (var item in strList)
            {
                amountCols.Add(System.Convert.ToInt32(item));
            }

            for (int i = System.Convert.ToInt32(config.BeginWith); i < dt.Rows.Count; i++)
            {
                var row = dt.Rows[i];
                double amount = 0;
                foreach (var item in amountCols)
                {
                    var s = row[item - 1].ToString().Replace('<', ' ').Replace('>', ' ').Replace(".", ",").Trim();
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
            File.Delete(localFile);
        }

        public DBFConverter(Price config)
        {
            this.config = config;
            LoadCoef();
        }
    }
}
