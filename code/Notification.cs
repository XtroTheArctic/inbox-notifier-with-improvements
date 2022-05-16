using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;
using notifier.Languages;
using notifier.Properties;

namespace notifier {
	class Notification {

		#region #attributes

		/// <summary>
		/// Privacy possibilities
		/// </summary>
		public enum Privacy : uint {
			None = 0,
			Short = 1,
			All = 2
		}

		/// <summary>
		/// Type possibilities
		/// </summary>
		public enum Type : uint {
			None = 0,
			Info = 1,
			Warning = 2,
			Error = 3
		}

		/// <summary>
		/// Behavior possibilities
		/// </summary>
		public enum Behavior : uint {
			DoNothing = 0,
			Open = 1,
			MarkAsRead = 2,
			OpenSimplifiedHTML = 3
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
		public Notification(ref Main form) {
			UI = form;
		}

		/// <summary>
		/// Display a notification balloon tip
		/// </summary>
		/// <param name="title">Title of the tip</param>
		/// <param name="text">Text contained in the balloon tip</param>
		/// <param name="icon">Icon of the tip</param>
		/// <param name="duration">How long the notification is displayed</param>
		public void Tip(string title, string text, Type icon = Type.Info, int duration = 450) {
			UI.notifyIcon.ShowBalloonTip(duration, title, text, (ToolTipIcon)icon);
		}

		/// <summary>
		/// Do the gmail specified action (inbox/message/spam) in a browser
		/// </summary>
		/// <param name="balloon">Define if the interaction is provided by the balloon tip</param>
		public async Task Interaction(bool balloon = false) {
				Process.Start($"{GetBaseURL()}/#inbox");
		}

		/// <summary>
		/// Pause the notifications during a certain time
		/// </summary>
		/// <param name="item">Selected item in the timeout menu</param>
		/// <param name="delay">Delay until the next inbox sync, 0 means "infinite" timeout</param>
		public void Pause(MenuItem item, int delay) {

			// exit if the selected menu item is already checked
			if (item.Checked) {
				return;
			}

			// disable notifications
			Paused = true;

			// clean the tag
			Tag = null;

			// uncheck all menu items
			foreach (MenuItem i in UI.menuItemTimout.MenuItems) {
				i.Checked = false;
			}

			// display the user choice
			item.Checked = true;

			// infinite variable
			bool infinite = delay == 0;

			// disable the timer if the delay is set to "infinite"
			UI.timer.Enabled = !infinite;

			// applie "1" if the delay is set to "infinite" because the timer interval attribute does not support "0"
			UI.timer.Interval = infinite ? 1 : delay;

			// update the systray icon and text
			UI.notifyIcon.Icon = Resources.timeout;
			UI.notifyIcon.Text = $"{Translation.timeout} - {(infinite ? "∞" : DateTime.Now.AddMilliseconds(delay).ToShortTimeString())}";

			// disable some menu items
			UI.menuItemMarkAsRead.Enabled = false;

			// update some text items
			UI.menuItemMarkAsRead.Text = Translation.markAsRead;
		}

		/// <summary>
		/// Resume notifications
		/// </summary>
		public async Task Resume() {

			// exit if the selected menu item is already checked
			if (UI.menuItemTimeoutDisabled.Checked) {
				return;
			}

			// enable notifications
			Paused = false;

			// uncheck all menu items
			foreach (MenuItem i in UI.menuItemTimout.MenuItems) {
				i.Checked = false;
			}

			// display the user choice
			UI.menuItemTimeoutDisabled.Checked = true;

			// restore the timer interval
			UI.timer.Interval = Settings.Default.TimerInterval;

			// enable the timer: this will automatically trigger the inbox synchronization in the timer tick
			UI.timer.Enabled = true;

			// synchronize the inbox
			await UI.GmailService.Inbox.Sync();
		}

		/// <summary>
		/// Return the Gmail base URL depending on the notification behavior
		/// </summary>
		/// <returns>URL to access Gmail simplified or full web interface</returns>
		public static string GetBaseURL() {
			return Settings.Default.NotificationBehavior == (uint)Behavior.OpenSimplifiedHTML ? $"{Settings.Default.GMAIL_BASEURL}/h" : Settings.Default.GMAIL_BASEURL;
		}

		#endregion

		#region #accessors

		/// <summary>
		/// Notification tag to allow the user to directly display the specified view (inbox/message/spam) in a browser
		/// </summary>
		public string Tag {
			get; set;
		}

		/// <summary>
		/// Timeout mode that indicates if the notification service is paused
		/// </summary>
		public bool Paused {
			get; set;
		}

		#endregion
	}
}