using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Xml.Serialization;
using System.Data;

namespace Importer
{
    public class Extend
    {
        internal Dictionary<string, object> GetDict(DataTable dt)
        {
            return dt.AsEnumerable()
              .ToDictionary<DataRow, string, object>(row => row.Field<string>(0),
                                        row => row.Field<object>(1));
        }
    }

    [Serializable()]
    public class App
    {
        const string FileName = "AppConfig.xml";
        
        public DataTable CrossReplace;

        public BindingList<Price> Prices { get; protected set; }

        public string Host { get; set; }
        public string User { get; set; }
        public string Pass { get; set; }

        [NonSerialized()]
        public static readonly App Instance = new App();

        static App()
        {
            if (ConfigFileExist())
            {
                var app = LoadFromFile();
                Instance.Prices = app.Prices;
                Instance.Host = app.Host;
                Instance.User = app.User;
                Instance.Pass = app.Pass;
                Instance.CrossReplace = app.CrossReplace;
            }
        }

        public App()
        {
            Prices = new BindingList<Price>();
            CrossReplace = new DataTable();
            CrossReplace.Columns.Add("Src");
            CrossReplace.Columns.Add("Dest");
            CrossReplace.TableName = "CrossReplace";
        }

        private static App LoadFromFile()
        {
            XmlSerializer xmlDeserializer = new XmlSerializer(typeof(App));

            var file = File.Open(GetHomeDir() + FileName, FileMode.Open);
            var app = xmlDeserializer.Deserialize(file) as App;
            file.Close();
            try
            {

                foreach (var p in app.Prices)
                {
                    if (p.Coeficients != null)
                    {
                        p.Coeficients = p.Coeficients.Replace("\n", "\r\n");
                    }
                    if (p.Prefixes != null)
                    {
                        p.Prefixes = p.Prefixes.Replace("\n", "\r\n");
                    }
                    if (p.Suffixes != null)
                    {
                        p.Suffixes = p.Suffixes.Replace("\n", "\r\n");
                    }
                    if (p.BrandExtra != null)
                    {
                        p.BrandExtra = p.BrandExtra.Replace("\n", "\r\n").Trim();
                    }
                    p.Init();
                }

            }
            catch (Exception e)
            {
                var msg = e.Message;

            }
            return app;
        }

        private static bool ConfigFileExist()
        {
            return (File.Exists(GetHomeDir() + FileName));
        }

        public static string GetHomeDir()
        {
            return Directory.GetCurrentDirectory() + "\\";
        }

        public static void Save()
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(App));

            var file = File.Open(GetHomeDir() + FileName, FileMode.Create);
            xmlSerializer.Serialize(file, Instance);
            file.Close();
        }
    }
}
