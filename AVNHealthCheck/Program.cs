using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Xml;
using Microsoft.Win32;
using ServiceHostHelper;
using ServiceHostHelper.Helper;
using ServiceHostHelper.McAfee;

namespace AVNHealthCheck
{
    class Program
	{
		private static string logfilepath = Environment.ExpandEnvironmentVariables("%LOGFILEDIR%");

		private static Logdatei logfile;

		private static string logfileName;

		private static int TimeOut;

		private static List<string> alarmedAVN;

		private static void Main(string[] args)
		{
			logfileName = Assembly.GetExecutingAssembly().GetName().Name + ".log";
			if (string.IsNullOrEmpty(logfilepath))
			{
				logfilepath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			}
			logfilepath += "\\NetAppLog";
			if (!Directory.Exists(logfilepath))
			{
				Directory.CreateDirectory(logfilepath);
			}
			logfile = new Logdatei(logfilepath, logfileName);
			Console.WriteLine("Logpath: " + logfilepath);
			Console.WriteLine("LogName: " + Assembly.GetExecutingAssembly().GetName().Name);
			loadConfig();
			alarmedAVN = getAlarmed();
			RegistryKey registryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\DEVK\\McAfee\\VSES\\Server", writable: false);
			if (registryKey == null)
			{
				return;
			}
			string empty = string.Empty;
			bool flag = false;
			string[] Servers = registryKey.GetValueNames();
			foreach (string Server in Servers)
			{
				empty = RegistryAccess.GetRegistryValue("SOFTWARE\\DEVK\\McAfee\\VSES\\Server", Server);
				try
				{
					if (DateTime.Compare(DateTime.ParseExact(empty.Split(';')[3], "yyyy_MM_dd_HH_mm", CultureInfo.InvariantCulture), DateTime.Now.AddMinutes(TimeOut * -1)) < 0)
					{
						if (!alarmedAVN.Contains(Server) && alarmtime())
						{
							alarmedAVN.Add(Server);
							logfile.WriteMessage(Server + " has no activity -> sending Mail");
							sendMessage(Server, "NetApp-Scan läuft nicht", "Es wurde länger als " + TimeOut + " Minuten keine Aktivität festgestellt!");
							
							flag = true;
						}
					}
					else if (alarmedAVN.Contains(Server) && alarmtime())
					{
						alarmedAVN.Remove(Server);
						sendMessage(Server, "NetApp-Scan wieder aktiv", "Letzte Aktivität: " + empty.Split(';')[3]);
						logfile.WriteMessage(Server + " is back again");
						flag = true;
					}
				}
				catch (Exception ex)
				{
					sendMessage(Server, "Read-Error bei Aktivitätscheck", "Bei der Kovertierung des letzten Datums trat ein Fehler auf - " + ex.Message);
				}
			}
			try
			{
				if (!flag)
				{
					return;
				}
				string text2 = string.Empty;
				foreach (string item in alarmedAVN)
				{
					text2 = text2 + item + ",";
				}
				if (text2.Length > 1)
				{
					RegistryAccess.SetRegistryValue("SOFTWARE\\DEVK\\McAfee\\VSES", "Alarmed", text2.Substring(0, text2.Length - 1));
				}
				else
				{
					RegistryAccess.DeleteRegistryValue("SOFTWARE\\DEVK\\McAfee\\VSES", "Alarmed");
				}
			}
			catch (Exception)
			{
				logfile.WriteMessage("FEHLER bei Registryzugriff!");
			}
		}

		private static bool alarmtime()
		{
			DateTime now = DateTime.Now;
			if (now.DayOfWeek == DayOfWeek.Saturday || now.DayOfWeek == DayOfWeek.Sunday)
			{
				return false;
			}
			if (now.Hour >= 22 || now.Hour <= 5)
			{
				return false;
			}
			return true;
		}

		private static void sendMessage(string from, string Betreff, string body)
		{
			McAfeeClient mcAfeeClient = new McAfeeClient(new BaseConfig(),false);
			mcAfeeClient.config.ServiceType = "global";

			McAfeeCommand mcAfeeCommand = new McAfeeCommand() { CmdType = McAfeeCommandType.SendMail };
			mcAfeeCommand.CommandText = "meldung2#,#" + from.ToUpper() + "#,#" + body + "#,#" + Betreff;
			mcAfeeClient.sendCommand(mcAfeeCommand, true);
		}

		private static List<string> getAlarmed()
		{
			string registryValue = RegistryAccess.GetRegistryValue("SOFTWARE\\DEVK\\McAfee\\VSES", "Alarmed");
			List<string> list = new List<string>();
			if (!string.IsNullOrEmpty(registryValue))
			{
				string[] array = registryValue.Split(',');
				foreach (string item in array)
				{
					list.Add(item);
				}
			}
			return list;
		}

		private static void loadConfig()
		{
			if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings["Timeout"]))
			{
				try
				{
					TimeOut = Convert.ToInt32(ConfigurationManager.AppSettings["Timeout"]);
				}
				catch
				{
					TimeOut = 60;
				}
			}
		}
	}

}
