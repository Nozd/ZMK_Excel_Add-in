using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Tools.Excel;

public abstract class AppConfig : IDisposable
{
    public static string GetCustomerAppConfigPath ()
    {
        string customerAppConfigFolder = string.Concat(@"C:\Program Files\TheFirstAddin\Application Files\TheFirstAddin_", Assembly.GetExecutingAssembly().GetName().Version.ToString().Replace(".", "_"));
        if (Directory.Exists(customerAppConfigFolder))
        {
            return string.Concat(customerAppConfigFolder, @"\TheFirstAddin.dll.config.deploy");
        }
        else
        {
            return String.Empty;
        }
    }
    public static AppConfig Change(string path)
    {
        return new ChangeAppConfig(path);
    }

    public abstract void Dispose();

    private class ChangeAppConfig : AppConfig
    {
        private readonly string oldConfig =
            AppDomain.CurrentDomain.GetData("APP_CONFIG_FILE").ToString();

        private bool disposedValue;

        public ChangeAppConfig(string path)
        {
            AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", path);
            ResetConfigMechanism();
        }

        public override void Dispose()
        {
            if (!disposedValue)
            {
                AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", oldConfig);
                ResetConfigMechanism();


                disposedValue = true;
            }
            GC.SuppressFinalize(this);
        }

        private static void ResetConfigMechanism()
        {
            typeof(ConfigurationManager)
                .GetField("s_initState", BindingFlags.NonPublic |
                                         BindingFlags.Static)
                .SetValue(null, 0);

            typeof(ConfigurationManager)
                .GetField("s_configSystem", BindingFlags.NonPublic |
                                            BindingFlags.Static)
                .SetValue(null, null);

            typeof(ConfigurationManager)
                .Assembly.GetTypes()
                .Where(x => x.FullName ==
                            "System.Configuration.ClientConfigPaths")
                .First()
                .GetField("s_current", BindingFlags.NonPublic |
                                       BindingFlags.Static)
                .SetValue(null, null);
        }
    }
}