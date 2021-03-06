using Microsoft.Win32;
using Synergy.Logging;
using Synergy.Logging.EventArgs;
using Synergy.Logging.Interfaces;
using System;
using System.Diagnostics;
using System.Management;
using System.Reflection;
using System.Security.Principal;
using System.ServiceProcess;
using System.Threading.Tasks;
using Enums = Synergy.Logging.Enums;

namespace DisableWindowsService {
	internal class Program {
		private static readonly ILogger Logger = new Logger(nameof(Program));

		/// <summary>
		/// Maximum number of parallel tasks we should run at a time.
		/// </summary>
		private const int MAX_PARALLEL_TASKS = 10;

		/// <summary>
		/// SuperFetch service name.
		/// </summary>
		private const string SUPERFETCH_SERVICE = "SysMain";

		/// <summary>
		/// Update service name.
		/// </summary>
		private const string UPDATE_SERVICE = "wuauserv";

		/// <summary>
		/// The services to disable.
		/// </summary>
		private static string[] ServiceNames;

		/// <summary>
		/// Boolean value indicating if the current user is the administrator of this system and the current instance is started as in Administrator mode.
		/// </summary>
		private static bool IsAdministrator => new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);

		private static async Task<int> Main(string[] args) {
			if (args != null && args.Length > 0) {
				ServiceNames = new string[args.Length + 2];

				for (int i = 0; i < ServiceNames.Length; i++) {
					if (args[i] == null) {
						continue;
					}

					ServiceNames[i] = args[i];
				}
			}
			else {
				ServiceNames = new string[2];
			}

			ServiceNames[^2] = SUPERFETCH_SERVICE;
			ServiceNames[^1] = UPDATE_SERVICE;

			Console.Title = $"{Assembly.GetExecutingAssembly().GetName().Name} v{Assembly.GetExecutingAssembly().GetName().Version}";
			global::Synergy.Logging.Logger.LogMessageReceived += Logger_LogMessageReceived;

			if (!IsAdministrator) {
				Logger.Error("Please run the program as administrator.");
				Logger.Warning("Press any key to exit...");
				Console.ReadKey(true);
				return -1;
			}

			if (!IsWindows10()) {
				Logger.Error("The running platform is not Windows 10.");
				Logger.Warning("Cannot continue... press any key to exit.");
				Console.ReadKey(true);
				return -1;
			}

			int successCount = 0;
			int completedCount = 0;
			Action[] actionsList = new Action[ServiceNames.Length];

			Stopwatch stopwatch = new Stopwatch();
			stopwatch.Start();
			for (int i = 0; i < ServiceNames.Length; i++) {
				string serviceName = ServiceNames[i];

				if (string.IsNullOrEmpty(serviceName)) {
					continue;
				}

				actionsList[i] = () => {
					try {
						bool isSuccess = StopService(serviceName) && DisableService(serviceName);

						if (isSuccess) {
							successCount++;
						}
					}
					finally {
						completedCount++;
					}
				};
			}

			Parallel.Invoke(new ParallelOptions() {
				MaxDegreeOfParallelism = MAX_PARALLEL_TASKS
			}, actionsList);

			while (completedCount != actionsList.Length) {
				await Task.Delay(1).ConfigureAwait(false);
			}

			stopwatch.Stop();
			Logger.WithColor($"It took {Math.Round(stopwatch.Elapsed.TotalSeconds, 3)} seconds to stop '{ServiceNames.Length}' services.", ConsoleColor.Green);

			if (successCount == ServiceNames.Length) {
				Logger.WithColor("Successfully disabled all services!", ConsoleColor.Green);
				Logger.Info("Existing in 10 seconds...");
				await Task.Delay(TimeSpan.FromSeconds(10)).ConfigureAwait(false);
				return 0;
			}

			Logger.WithColor($"'{successCount}' services succeeded out of '{ServiceNames.Length}' services.", ConsoleColor.Green);
			Logger.Info("Existing in 10 seconds...");
			await Task.Delay(TimeSpan.FromSeconds(10)).ConfigureAwait(false);
			return -1;
		}

		/// <summary>
		/// Tries to stop the specified service with the service name specified in parameter.
		/// </summary>
		/// <param name="serviceName">The service to stop</param>
		/// <returns>Boolean indicating success or fail result</returns>
		private static bool StopService(string serviceName) {
			if (string.IsNullOrEmpty(serviceName)) {
				return false;
			}

			try {
				using (ServiceController sc = new ServiceController(serviceName)) {
					Logger.Info($"'{serviceName}' service is currently in '{sc.Status}' state.");

					if (sc.Status == ServiceControllerStatus.Stopped) {
						Logger.WithColor($"Skipping '{serviceName}' service as its already stopped!", ConsoleColor.Green);
						return true;
					}

					Logger.Info($"Trying to stop '{serviceName}' service...");

					if (!sc.CanStop) {
						Logger.Warning($"'{serviceName}' service can't be stopped.");
						return false;
					}

					sc.Stop();
					sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(10));

					if (sc.Status == ServiceControllerStatus.Stopped) {
						Logger.WithColor($"'{serviceName}' service has been stopped!", ConsoleColor.Green);
						return true;
					}

					return false;
				}
			}
			catch (Exception e) {
				Logger.Exception(e);
				return false;
			}
		}

		/// <summary>
		/// Tries to disable the specified service with the service name specified in parameter
		/// </summary>
		/// <param name="serviceName">The service to stop</param>
		/// <returns>Boolean indicating success or fail result</returns>
		private static bool DisableService(string serviceName) {
			if (string.IsNullOrEmpty(serviceName)) {
				return false;
			}

			Logger.Info($"Trying to disable '{serviceName}' service...");

			try {
				using (ManagementObject mo = new ManagementObject(string.Format("Win32_Service.Name=\"{0}\"", serviceName))) {
					mo.InvokeMethod("ChangeStartMode", new object[] { "Disabled" });
				}

				Logger.WithColor($"'{serviceName}' service disabled successfully!", ConsoleColor.Green);
				return true;
			}
			catch (Exception e) {
				Logger.Exception(e);
				Logger.Error($"Failed to disable '{serviceName}' service.");
				return false;
			}
		}

		/// <summary>
		/// Check if the current running OS is windows 10
		/// </summary>
		/// <returns>Boolean indicating if OS is windows 10 or not</returns>
		private static bool IsWindows10() {
			try {
				using (RegistryKey reg = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion")) {
					return ((string) reg.GetValue("ProductName")).StartsWith("Windows 10");
				}
			}
			catch (Exception e) {
				Logger.Exception(e);
				return false;
			}
		}

		/// <summary>
		/// Handles all log messages in this application.
		/// </summary>
		/// <param name="sender">The sender</param>
		/// <param name="e">The log message object</param>
		private static void Logger_LogMessageReceived(object sender, LogMessageEventArgs e) {
			if (e == null) {
				return;
			}

			switch (e.LogLevel) {
				case Enums.LogLevels.Trace:
				case Enums.LogLevels.Debug:
				case Enums.LogLevels.Info:
					Console.ForegroundColor = ConsoleColor.White;
					Console.WriteLine($"{e.ReceivedTime} | {e.LogIdentifier} | {e.LogLevel} | {e.CallerMemberName}() {e.LogMessage}");
					break;
				case Enums.LogLevels.Warn:
					Console.ForegroundColor = ConsoleColor.Magenta;
					Console.WriteLine($"{e.ReceivedTime} | {e.LogIdentifier} | {e.LogLevel} | {e.CallerMemberName}() {e.LogMessage}");
					break;
				case Enums.LogLevels.Error:
					Console.ForegroundColor = ConsoleColor.Red;
					Console.WriteLine($"{e.ReceivedTime} | {e.LogIdentifier} | {e.LogLevel} | {e.CallerMemberName}() {e.LogMessage}");
					break;
				case Enums.LogLevels.Exception:
					Console.ForegroundColor = ConsoleColor.Red;
					Console.WriteLine($"{e.ReceivedTime} | {e.LogIdentifier} | {e.LogLevel} | {e.CallerMemberName}() {e.LogMessage}");
					break;
				case Enums.LogLevels.Fatal:
					Console.ForegroundColor = ConsoleColor.Yellow;
					Console.WriteLine($"{e.ReceivedTime} | {e.LogIdentifier} | {e.LogLevel} | {e.CallerMemberName}() {e.LogMessage}");
					break;
				case Enums.LogLevels.Green:
					Console.ForegroundColor = ConsoleColor.Green;
					Console.WriteLine($"{e.ReceivedTime} | {e.LogIdentifier} | {e.LogLevel} | {e.CallerMemberName}() {e.LogMessage}");
					break;
				case Enums.LogLevels.Red:
					Console.ForegroundColor = ConsoleColor.Red;
					Console.WriteLine($"{e.ReceivedTime} | {e.LogIdentifier} | {e.LogLevel} | {e.CallerMemberName}() {e.LogMessage}");
					break;
				case Enums.LogLevels.Blue:
					Console.ForegroundColor = ConsoleColor.Blue;
					Console.WriteLine($"{e.ReceivedTime} | {e.LogIdentifier} | {e.LogLevel} | {e.CallerMemberName}() {e.LogMessage}");
					break;
				case Enums.LogLevels.Cyan:
					Console.ForegroundColor = ConsoleColor.Cyan;
					Console.WriteLine($"{e.ReceivedTime} | {e.LogIdentifier} | {e.LogLevel} | {e.CallerMemberName}() {e.LogMessage}");
					break;
				case Enums.LogLevels.Magenta:
					Console.ForegroundColor = ConsoleColor.Magenta;
					Console.WriteLine($"{e.ReceivedTime} | {e.LogIdentifier} | {e.LogLevel} | {e.CallerMemberName}() {e.LogMessage}");
					break;
				case Enums.LogLevels.Input:
					Console.ForegroundColor = ConsoleColor.Gray;
					Console.WriteLine($"{e.ReceivedTime} | {e.LogIdentifier} | {e.LogLevel} | {e.CallerMemberName}() {e.LogMessage}");
					break;
				case Enums.LogLevels.Custom:
					Console.WriteLine($"{e.ReceivedTime} | {e.LogIdentifier} | {e.LogLevel} | {e.CallerMemberName}() {e.LogMessage}");
					break;
			}

			Console.ResetColor();
		}
	}
}
