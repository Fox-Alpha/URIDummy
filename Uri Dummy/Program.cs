/*
 * Erstellt mit SharpDevelop.
 * Benutzer: buck
 * Datum: 13.09.2016
 * Zeit: 09:47
 * 
 */
using System;
using System.IO;
using Microsoft.Win32;
using System.Diagnostics;
using System.Threading;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Collections.Generic;
using System.Security.Principal;
using System.ComponentModel;

namespace Uri_Dummy
{
	class Program
	{
		[DllImport ("kernel32.dll")]
		static extern IntPtr GetConsoleWindow ();

		[DllImport ("user32.dll")]
		static extern bool ShowWindow (IntPtr hWnd, int nCmdShow);

		[DllImport("user32.dll")]
		private static extern bool SetForegroundWindow(IntPtr hWnd);
		
		[DllImport("user32.dll")]
		private static extern IntPtr GetForegroundWindow();		
		
		[DllImport("user32.dll", SetLastError=true)]
		static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

		[DllImport("user32.dll", SetLastError=true)]
		static extern bool SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
		
		[DllImport("user32.dll", SetLastError=true)]
		static extern IntPtr SetActiveWindow(IntPtr hWnd);
		
		const int SW_HIDE = 0;
		const int SW_SHOW = 5;

		static string RegPath2Uninstall32 = "wow6432Node\\";
        static string RegPath2Uninstall = "Software\\{0}Microsoft\\Windows\\CurrentVersion\\Uninstall\\";
		static string DameWareInstallPath = "";
		static string strTargetHost = "";
		static bool dwFound = false;
		static IntPtr przHandle;

		public static void Main(string[] args)
		{
			Console.WriteLine ("Suche nach DameWare Installation in der ...");
			if (Environment.Is64BitOperatingSystem) // && Environment.Is64BitProcess)
			{
				//  32 & 64 BIT Registry Bereiche lesen
				Debug.WriteLine ("... 32 Bit Registry");
				Debug.WriteLine ("-------------------------------------------");
				Console.WriteLine ("... 32 Bit Registry");
				dwFound = SearchDameWarePathInRegistry (false);
				Debug.WriteLine ("... 64 Bit Registry");
				Debug.WriteLine ("-------------------------------------------");
				Console.WriteLine ("... 64 Bit Registry");
				dwFound = SearchDameWarePathInRegistry (true);
			}
			else
			{
				//  nur 32 Bit vorhanden
				Console.WriteLine ("... 32 Bit Registry");
				dwFound = SearchDameWarePathInRegistry (true);
			}
			//var handle = GetConsoleWindow ();

			// Hide
			//ShowWindow (handle, SW_HIDE);

			// Show
			//ShowWindow (handle, SW_SHOW);

			string [] cmd;
			int i = 1;
			
			//	Aufruf ist mit mindestens einen Parameter erfolgt
			if ((cmd = Environment.GetCommandLineArgs()).Length > 1)
			{
#if (DEBUG)
				Console.WriteLine ("Anwendung wurde mit diesen Parametern Aufgerufen");
				foreach (var str in cmd) {
					Console.WriteLine("{0} = {1}", i,  str);
					i++;
				}
#endif

				//	Prüfen ob Registry Eintrag für URI vorhandenb ist
				if (!CheckUriInRegistry ())
				{
					SetUriInRegistry ();
					Console.WriteLine ("Press any key to continue . . . ");
					Console.ReadKey (true);
				}

				//	Wenn DameWare vorhanden ist und eine Hostadresse angegeben wurde
				//	Dann Dameware mit Parameter aufrufen
				if (dwFound && ExtractTargetHostAdressFromCommandLine ())
				{
					Process dw = new Process ();
					ProcessStartInfo dwpci = new ProcessStartInfo (DameWareInstallPath, string.Format("-x -c: -h: -m:{0}", strTargetHost));

					dw.StartInfo = dwpci;
					dw.Start ();
					
					// DameWare Fenster in den Vordergrund setzen
					if (dw.MainWindowHandle != null) 
					{	
						if(ActivateWindow(dw.MainWindowHandle, dw.Id, out przHandle))
						{
							Debug.WriteLine(string.Format("{0} / {1}", dw.MainWindowHandle, przHandle), "ActivateWindow()");
						}
						else
							Debug.WriteLine(string.Format("{0} / {1}", dw.MainWindowHandle, przHandle), "ActivateWindow()");
					}
				}

				if (!dwFound)
				{
					Console.WriteLine ("DameWare Installation wurde nicht in der Registry gefunden !");
					Console.WriteLine ("Press any key to continue . . . ");
					Console.ReadKey (true);
					return;
				}
			}
			//	Aufruf ohne Parameter, Prüfen der URI in der Registry
			else
			{
				if (!dwFound)
				{
					Console.WriteLine ("DameWare Installation wurde nicht in der Registry gefunden !");
					Console.WriteLine ("Press any key to continue . . . ");
					Console.ReadKey (true);
					return;
				}

				//	Prüfen ob Registry Eintrag für URI vorhandenb ist
				if (!CheckUriInRegistry ())
				{
					SetUriInRegistry ();
				}
			}
#if (DEBUG)
			Console.WriteLine ("Press any key to continue . . . ");
			Console.ReadKey (true);
#endif
		}

		private static void SetUriInRegistry ()
		{
			Console.WriteLine ("... Setzen der URI in Registry ");

			//if (!(new WindowsPrincipal (WindowsIdentity.GetCurrent ()).IsInRole (WindowsBuiltInRole.Administrator)))
			//{
			//	
			//	return;
			//}

			WindowsPrincipal principal = new WindowsPrincipal (WindowsIdentity.GetCurrent ());
			if (principal.IsInRole (WindowsBuiltInRole.Administrator) == false)
			{
				//Get the full qualified path to the own executable
				Console.WriteLine ("Zum anlegen des Schlüssels muss die Anwendung mit Administrationsrechten gestartet sein");
				Console.WriteLine ("Anwendung wird jetzt neugestartet");
				string x = Assembly.GetExecutingAssembly ().GetName ().CodeBase;

				//Start itself with adminrights
				if (RunElevated (x) == true)
				{
					//Close old programwindow
					Environment.Exit (0);
				}
			}

			RegistryKey key;
			FileInfo fi = new FileInfo (Assembly.GetExecutingAssembly ().Location);

			try
			{
				List<String> test = new List<string> ();
				test.AddRange (Registry.ClassesRoot.GetSubKeyNames ());
				if (!test.Contains ("dwrcc"))
				{
					Console.WriteLine ("Schlüssel nicht vorhanden");
					
					if((key = Registry.ClassesRoot.CreateSubKey ("dwrcc"))!= null)
						Console.WriteLine ("Schlüssel wurde erstellt");


					if ((key = Registry.ClassesRoot.OpenSubKey ("dwrcc", true)) != null)
					{
						key.SetValue ("", "URL: Protocol handled by DameWare");
						key.SetValue ("URL Protocol", "");

						//	Setzen des Icons auf das DameWare Icon
						key.CreateSubKey ("DefaultIcon").SetValue ("", "\"" + DameWareInstallPath + "\",1");

						//	Setzen des 'Command' zum Aufruf der Anwendung
						key.CreateSubKey ("shell\\open\\command").SetValue ("", fi.FullName + " %1");
					}
				}
				else
					Console.WriteLine ("Schlüssel vorhanden");
			}
			catch (Exception ex)
			{
				Debug.WriteLine (ex.Message);
				Debug.WriteLine (ex.Data);
				Console.WriteLine (ex.Message);
			}
		}

		private static bool RunElevated (string fileName)
		{
			ProcessStartInfo processInfo = new ProcessStartInfo ();
			processInfo.Verb = "runas";
			processInfo.FileName = fileName;
			try
			{
				Process.Start (processInfo);
				return true;
			}
			catch (Win32Exception)
			{
				//Do nothing. Probably the user canceled the UAC window
			}
			return false;
		}

		private static bool CheckUriInRegistry ()
		{
			Console.WriteLine ("... Prüfen ob URI in Registry vorhanden ist");
			Debug.WriteLine ("... Prüfen ob URI in Registry vorhanden ist");
			RegistryKey key = Registry.ClassesRoot.OpenSubKey ("dwrcc", false);
			FileInfo fi = new FileInfo (Assembly.GetExecutingAssembly ().Location);

			if (key == null)
				return false;

			//	Prüfen ob der Pfad zur Anwendung korrekt eingetragen ist
			Console.WriteLine ("... Öffnen des Command Schlüssels");
			Debug.WriteLine ("... Öffnen des Command Schlüssels");
			if (key.OpenSubKey ("shell\\open\\command", false) != null)
			{
				string [] reg = key.OpenSubKey ("shell\\open\\command", false).GetValueNames ();
				foreach (var str in reg)
				{
					//	Der Aufruf steht in dem Standardschlüssel
					if (str == "")
					{
						Console.WriteLine ("... Lesen des Pfades Schlüssels");
						
						string val = key.OpenSubKey ("shell\\open\\command", false).GetValue ("").ToString ();
						string valApp = "";
						string valReg = "";
						
						Debug.WriteLine ("Pfad in Registry: " + key.OpenSubKey ("shell\\open\\command", false).GetValue (""));
						RegistryKey regCommand = key.OpenSubKey ("shell\\open\\command", false);
						valReg = val;
						valApp = Path.GetFileName (val);
						val = Path.GetDirectoryName (val);

						//	Sichergehen das am Ende des Pfades ein Backslash steht
						if (!val.EndsWith (@"\", StringComparison.CurrentCulture))
						{
							val += @"\";
						}

						//	Vergleich des Pfads in der Registry mit dem der Anwendung
						if ((val == AppDomain.CurrentDomain.BaseDirectory) && (Regex.IsMatch(valApp, fi.Name)))
//						string input = Regex.Escape(valReg);
//						string pattern = Regex.Escape(fi.FullName);
//						
//						if (Regex.Match(input, pattern, RegexOptions.CultureInvariant).Success)
						{
							Console.WriteLine ("Pfad ist korrekt");
							return true;
						}
						else
						{
							//	Pfad inkorrekt, korrigieren
							Console.WriteLine ("Pfad ist NICHT korrekt");
							Console.WriteLine ("BaseDirectory: " + AppDomain.CurrentDomain.BaseDirectory);
							Console.WriteLine ("Registry Pfad: " + val);

							Debug.WriteLine ("Pfad ist NICHT korrekt");
							Debug.WriteLine ("BaseDirectory: " + AppDomain.CurrentDomain.BaseDirectory);
							Debug.WriteLine ("Registry Pfad: " + val);
							
							WindowsPrincipal principal = new WindowsPrincipal (WindowsIdentity.GetCurrent ());
							if (principal.IsInRole (WindowsBuiltInRole.Administrator) == false)
							{
								Console.WriteLine ("Für eine Korrektur, muss die Anwendung als Administrator ohne angabe von Parametern gestartet werden.");
							}
							else
							{
								regCommand.Close ();
								regCommand = key.OpenSubKey ("shell\\open\\command", true);
								regCommand.SetValue ("", fi.FullName + " %1");
								regCommand.Close ();
								Console.WriteLine ("Die Korrektur des Aufruf Pfades wurde vorgenommen.");
								return true;
							}
								
								//Console.WriteLine (AppDomain.CurrentDomain.BaseDirectory + System.Reflection.Assembly.GetExecutingAssembly ().GetName ().Name + ".exe");
							break;
						}
					}
					else
						continue;
				}
			}
			else
				Console.WriteLine ("... Öffnen des Command Schlüssels > Gescheitert");

			return false;
		}

		private static bool ExtractTargetHostAdressFromCommandLine ()
		{
			Match mc;
			string [] cmd;
			string match;

			if ((cmd = Environment.GetCommandLineArgs ()).Length > 0)
			{
				//	Protokoll Prefix und Suffix entfernen
				mc = Regex.Match(Environment.CommandLine, "dwrcc://(.*?)/", RegexOptions.IgnoreCase);
				if (mc.Success)
				{
					//	!!!Keine Prüfung der Hostadresse auf gültigkeit !!!
					match = mc.Groups [1].Value;
					Debug.WriteLine (match.ToString (), "Übergebene HostAdresse");
					Console.WriteLine (match.ToString (), "Übergebene HostAdresse");

					strTargetHost = match;

					return true;
				}
			}
			return false;
		}

		static bool SearchDameWarePathInRegistry(bool is64Bit)
        {
			string activePath = string.Empty;
			RegistryKey key;
			
			//	Wenn die Anwendung kein 64Bit Process ist, kann diese nur mit Umweg auf den 64Bit Bereich der Registry zugreifen
			if (is64Bit && !Environment.Is64BitProcess) 
			{
	            activePath = string.Format(RegPath2Uninstall, "");
	            key = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(activePath,false);				
			}
			else
			{
	            activePath = string.Format(RegPath2Uninstall, is64Bit ? "" : RegPath2Uninstall32);
	            key = Registry.LocalMachine.OpenSubKey(activePath,false);
			}
            
            string[] valueNames;
            string activeAppPath;
			//bool 
				dwFound = false;

            if (key != null && key.SubKeyCount > 0)
            {
                //  Alle Schlüssel der Installierten Anwendungen
                foreach (string  appKey in key.GetSubKeyNames())
                {
                	//	Jeder Schlüssel enspricht einer Anwendung
                	//	Hat der aktive Schlüssel unterschlüssel ?
                    if((valueNames = key.OpenSubKey(appKey, false).GetValueNames()).Length > 0)
                    {
                    	//	JA !
                    	//	Pfad merken
                    	activeAppPath = activePath + appKey;
                    	object tmp;
                    	foreach (var valName in valueNames)
						{
							// Den Hersteller prüfen
							if (valName == "Publisher")
							{
								if ((tmp = key.OpenSubKey (appKey, false).GetValue (valName)).ToString () == "SolarWinds")
								{
									Debug.WriteLine ("DameWare Installation gefunden ");
									//Debug.WriteLine (string.Format("{0,-30} {1}",key.OpenSubKey (appKey, false).GetValue (valName).ToString (), activeAppPath), "Publisher");
									Debug.WriteLine ("Publisher: {0,-30} \r\nRegKey: {1}\r\nFullRegPath: {2}", new[] { key.OpenSubKey (appKey, false).GetValue (valName).ToString (), appKey, activeAppPath});

									Console.WriteLine ("DameWare Installation gefunden ");
									Console.WriteLine ("Publisher: {0,-30} \r\nRegKey: {1}\r\nFullRegPath: {2}",key.OpenSubKey (appKey, false).GetValue (valName).ToString (), appKey, activeAppPath);

									// Prüfen der Version
									if (Regex.IsMatch ((tmp = key.OpenSubKey (appKey, false).GetValue ("DisplayVersion")).ToString (), "12.0.4010.3"))
									{
										//	Installationspfad merken
										DameWareInstallPath = key.OpenSubKey (appKey, false).GetValue ("InstallLocation").ToString ();
										if (!string.IsNullOrWhiteSpace (DameWareInstallPath)) // && Directory.Exists(DameWareInstallPath))
										{
											//	Exe Namen für Aufruf an Pfad anhängen
											dwFound = File.Exists (DameWareInstallPath + "dwrcc.exe");
											DameWareInstallPath += "dwrcc.exe";
											Console.WriteLine ("InstallationsPfad: {0}", DameWareInstallPath);
											Debug.WriteLine ("InstallationsPfad: {0}", new[] { DameWareInstallPath });
											
											Debug.WriteLine("Anwendungsname: {0}", new[] { key.OpenSubKey (appKey, false).GetValue ("DisplayName").ToString ()});
											Console.WriteLine("Anwendungsname: {0}", key.OpenSubKey (appKey, false).GetValue ("DisplayName").ToString ());
												
											return dwFound;
										}
									}

								}
								else
								{
									if (!Regex.IsMatch(tmp.ToString(), "Microsoft")) {
										Debug.WriteLine (string.Format("{0,-30} {1}",key.OpenSubKey (appKey, false).GetValue (valName).ToString (), activeAppPath), "Publisher");
									}
									
									continue;
								}
							}
                    	}
                    }
					if (dwFound)
					{
						break;
					}
                }
            }
			return false;
        }

		static bool ActivateWindow(IntPtr mainWindowHandle, int procID, out IntPtr newHandle)
		{
			uint iProcessID = 0;
			IntPtr handle = IntPtr.Zero;
			newHandle = IntPtr.Zero;
			IntPtr fgHandle = IntPtr.Zero;
			
		    //check if already has focus
		    if (mainWindowHandle == GetForegroundWindow()) 
		    	return true;
		
		    //check if window is minimized
		    if(SwitchToThisWindow(mainWindowHandle, true))
		    	Debug.WriteLine("SwitchToThisWindow erfolgreich", "ActivateWindow()");
		    else
		    	Debug.WriteLine("SwitchToThisWindow nicht erfolgreich", "ActivateWindow()");
		
		    fgHandle = SetActiveWindow(mainWindowHandle);
		    SetForegroundWindow(mainWindowHandle);
		    
		    Thread.Sleep(500);
		    
		    handle = GetForegroundWindow();
		    
		    GetWindowThreadProcessId(handle, out iProcessID);
		    
		    Debug.WriteLine(string.Format("{0} / {1} / {2}", mainWindowHandle, handle, fgHandle), "ActivateWindow()");
		    
		    if (iProcessID == procID)
		    {
		    	newHandle = handle;
		    	return true;
		    }
		    
		    return false;
		}
	}
}