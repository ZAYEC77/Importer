using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Xml.Serialization;

namespace Importer
{
    [Serializable()]
    public class App
    {
        const string filename = "AppConfig.xml";

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
            }
        }

        public App()
        {
            Prices = new BindingList<Price>();
        }

        private static App LoadFromFile()
        {
            XmlSerializer xmlDeserializer = new XmlSerializer(typeof(App));

            var file = File.Open(GetHomeDir() + filename, FileMode.Open);
            var app = xmlDeserializer.Deserialize(file) as App;
            file.Close();
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
                p.Init();
            }
            return app;
        }

        private static bool ConfigFileExist()
        {
            return (File.Exists(GetHomeDir() + filename));
        }

        public static string GetHomeDir()
        {
            return Directory.GetCurrentDirectory() + "\\";
        }

        public static void Save()
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(App));

            var file = File.Open(GetHomeDir() + filename, FileMode.Create);
            xmlSerializer.Serialize(file, Instance);
            file.Close();
        }
    }
}
