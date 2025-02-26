using System;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using notifier.Properties;

namespace notifier {
	class Computer {

		#region #attributes

		/// <summary>
		/// Registration possibilities
		/// </summary>
		public enum Registration : uint {
			Off = 0,
			On = 1
		}

		/// <summary>
		/// Reference to the main interface
		/// </summary>
		private readonly Main UI;

		#endregion

		#region #methods

		/// <summary>
		/// Class constructor
		/// </summary>
		/// <param name="form">Reference to the application main window</param>
		public Computer(ref Main form) {
			UI = form;
		}

		/// <summary>
		/// Bind the "PowerModeChanged" event to automatically pause/resume the application synchronization
		/// </summary>
		public void BindPowerMode() {
			SystemEvents.PowerModeChanged += new PowerModeChangedEventHandler((object source, PowerModeChangedEventArgs target) => {
				if (target.Mode == PowerModes.Suspend) {

					// suspend the main timer
					//xtro: UI.timer.Enabled = false;
				}
				else if (target.Mode == PowerModes.Resume) UI.timer.Enabled = true;
			});
		}

		/// <summary>
		/// Bind the "SessionSwitch" event to automatically sync the inbox on session unlocking
		/// </summary>
		public void BindSessionSwitch() {
			SystemEvents.SessionSwitch += new SessionSwitchEventHandler(async (object source, SessionSwitchEventArgs target) => {

				// sync the inbox when the user is unlocking the Windows session
				if (target.Reason == SessionSwitchReason.SessionUnlock) {

					// do nothing if the timeout mode is set to infinite
					if (UI.NotificationService.Paused && UI.menuItemTimeoutIndefinitely.Checked) {
						return;
					}

					// synchronize the inbox and renew the token
					await UI.GmailService.Inbox.Sync();

					// enable the timer properly
					UI.timer.Enabled = true;
				} else if (target.Reason == SessionSwitchReason.SessionLock) {

					// suspend the main timer
					//xtro: UI.timer.Enabled = false;

					UI.GmailService.Inbox.ReconnectionAttempts = 0;
				}
			});
		}

		/// <summary>
		/// Asynchronous method to check the internet connectivity
		/// </summary>
		/// <returns>Indicate if the user is connected to the internet, false means that the request to the DNS_REGISTRY_IP and Google server has failed</returns>
		public static async Task<bool> IsInternetAvailable() {
			try {

				// send a ping to the DNS registry
				PingReply reply = await new Ping().SendPingAsync(Settings.Default.DNS_REGISTRY_IP, 1000, new byte[32]);

				if (reply.Status == IPStatus.Success) {
					return true;
				} else {

					// use Google secured homepage as alternative to the DNS ping
					using (WebClient client = new WebClient()) {
						using (Stream stream = await client.OpenReadTaskAsync("https://www.google.com")) {
							return true;
						}
					}
				}
			} catch (Exception) {
				return false;
			}
		}

		/// <summary>
		/// Regulate the start with Windows setting against the registry to prevent bad registry reflection
		/// </summary>
		public static void RegulatesRegistry() {
			using (RegistryKey key = Registry.CurrentUser.OpenSubKey(Settings.Default.REGISTRY_KEY, true)) {
				if (key.GetValue("Inbox Notifier") != null) {
					if (!Settings.Default.RunAtWindowsStartup) {
						Settings.Default.RunAtWindowsStartup = true;
					}
				} else {
					if (Settings.Default.RunAtWindowsStartup) {
						Settings.Default.RunAtWindowsStartup = false;
					}
				}
			}
		}

		/// <summary>
		/// Register or unregister the application from Windows startup program list
		/// </summary>
		/// <param name="mode">The registration mode for the application, Off means that the application will no longer be started at Windows startup</param>
		public static void SetApplicationStartup(Registration mode) {
			using (RegistryKey key = Registry.CurrentUser.OpenSubKey(Settings.Default.REGISTRY_KEY, true)) {
				if (mode == Registration.On) {
					key.SetValue("Inbox Notifier", $"{Application.ExecutablePath}");
				} else {
					key.DeleteValue("Inbox Notifier", false);
				}
			}
		}

		#endregion

		#region #accessors

		#endregion
	}
}