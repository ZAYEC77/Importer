using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Cache;
using System.Windows.Forms;

namespace Importer
{
    internal class Updater
    {
        public static void ReplaceUpdate()
        {
            var newFilePath = GetNewFilePath();
            if (File.Exists(newFilePath))
            {
                var currentAppPath = Path.Combine(System.Environment.CurrentDirectory, "Importer.exe");
                var oldPath = Path.Combine(System.Environment.CurrentDirectory, "Importer_old.exe");
                if (File.Exists(oldPath))
                {
                    File.Delete(oldPath);
                }
                System.IO.File.Move(currentAppPath, oldPath);
                System.IO.File.Move(newFilePath, currentAppPath);
                Process.Start(Application.ExecutablePath);
                Process.GetCurrentProcess().Kill();
            }
        }

        private static string GetNewFilePath()
        {
            return Path.Combine(System.Environment.CurrentDirectory, "Importer_new.exe");
        }

        internal void Update()
        {
            var versionFile = GetVersionFilePath();
            var remoteVersion = GetRemoteVersion();
            DownloadApp();
            File.WriteAllText(versionFile, remoteVersion.ToString());
            Updater.ReplaceUpdate();
        }

        internal static void CheckUpdate()
        {
            Updater updater = new Updater();
            if (updater.HasUpdates())
            {
                var result = MessageBox.Show("Доступне оновлення програми. Оновити?", "Оновлення програми", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    updater.Update();
                }
            }
        }

        protected void DownloadApp()
        {
            WebClient webClient = new WebClient();
            var localPath = GetNewFilePath();
            webClient.DownloadFile(@"https://raw.githubusercontent.com/ZAYEC77/Importer/master/Repository/Importer.exe", localPath);
        }

        internal bool HasUpdates()
        {
            try
            {
                int currentVersion = GetCurrentVersion();
                var remoteVersion = GetRemoteVersion();
                return remoteVersion > currentVersion;
            }
            catch (Exception)
            {
                return false;
            }

        }

        private int GetRemoteVersion()
        {
            RequestCachePolicy policy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
            var webRequest = WebRequest.Create(@"https://raw.githubusercontent.com/ZAYEC77/Importer/master/Repository/version.dat");
            WebRequest.DefaultCachePolicy = policy;
            var version = 0;
            using (var response = webRequest.GetResponse())
            using (var content = response.GetResponseStream())
            using (var reader = new StreamReader(content))
            {
                var text = reader.ReadToEnd();
                int.TryParse(text, out version);
            }

            return version;
        }

        private int GetCurrentVersion()
        {
            var versionFile = GetVersionFilePath();
            var version = 0;
            if (File.Exists(versionFile))
            {
                var text = File.ReadAllText(versionFile);
                int.TryParse(text, out version);
            }
            return version;
        }

        private string GetVersionFilePath()
        {
            return Path.Combine(System.Environment.CurrentDirectory, "version.dat");
        }
    }
}