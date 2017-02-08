﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Media;
using System.Net;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.Win32;
using notifier.Languages;
using notifier.Properties;

namespace notifier {
	public partial class Main : Form {

		// privacy possibilities
		private enum Privacy : int {
			None = 0,
			Short = 1,
			All = 2
		}

		// update period possibilities
		private enum Period : int {
			Startup = 0,
			Day = 1,
			Week = 2,
			Month = 3
		}

		// gmail api service
		private GmailService service;

		// user credential for the gmail authentication
		private UserCredential credential;

		// inbox label
		private Google.Apis.Gmail.v1.Data.Label inbox;

		// unread threads
		private int? unreadthreads = 0;

		// number of automatic reconnection
		private int reconnect = 0;

		// local application data folder name
		private string appdata = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/Gmail Notifier";

		// number of maximum automatic reconnection
		private const int MAX_AUTO_RECONNECT = 3;

		// number in seconds between reconnections (in seconds)
		private const int INTERVAL_RECONNECT = 10;

		// github repository root link
		private const string GITHUB_REPOSITORY = "https://github.com/xavierfoucrier/gmail-notifier";

		// gmail base root link
		private const string GMAIL_BASEURL = "https://mail.google.com/mail/u/0";

		/// <summary>
		/// Initializes the class
		/// </summary>
		public Main() {
			InitializeComponent();
		}

		/// <summary>
		/// Loads the form
		/// </summary>
		private void Main_Load(object sender, EventArgs e) {

			// hides the form by default
			Visible = false;

			// displays a systray notification on first load
			if (Settings.Default.FirstLoad && !Directory.Exists(appdata)) {
				notifyIcon.ShowBalloonTip(7000, translation.welcome, translation.firstLoad, ToolTipIcon.Info);

				// switchs the first load state
				Settings.Default.FirstLoad = false;
				Settings.Default.Save();

				// waits for 7 seconds to complete the thread
				System.Threading.Thread.Sleep(7000);
			}

			// configures the help provider
			HelpProvider help = new HelpProvider();
			help.SetHelpString(fieldRunAtWindowsStartup, translation.helpRunAtWindowsStartup);
			help.SetHelpString(fieldAskonExit, translation.helpAskonExit);
			help.SetHelpString(fieldLanguage, translation.helpLanguage);
			help.SetHelpString(labelEmailAddress, translation.helpEmailAddress);
			help.SetHelpString(labelTokenDelivery, translation.helpTokenDelivery);
			help.SetHelpString(buttonGmailDisconnect, translation.helpGmailDisconnect);
			help.SetHelpString(fieldMessageNotification, translation.helpMessageNotification);
			help.SetHelpString(fieldAudioNotification, translation.helpAudioNotification);
			help.SetHelpString(fieldSpamNotification, translation.helpSpamNotification);
			help.SetHelpString(fieldNumericDelay, translation.helpNumericDelay);
			help.SetHelpString(fieldStepDelay, translation.helpStepDelay);
			help.SetHelpString(fieldNetworkConnectivityNotification, translation.helpNetworkConnectivityNotification);
			help.SetHelpString(fieldAttemptToReconnectNotification, translation.helpAttemptToReconnectNotification);
			help.SetHelpString(fieldPrivacyNotificationNone, translation.helpPrivacyNotificationNone);
			help.SetHelpString(fieldPrivacyNotificationShort, translation.helpPrivacyNotificationShort);
			help.SetHelpString(fieldPrivacyNotificationAll, translation.helpPrivacyNotificationAll);

			// authenticates the user
			this.AsyncAuthentication();

			// attaches the context menu to the systray icon
			notifyIcon.ContextMenu = contextMenu;

			// binds the "PropertyChanged" event of the settings to automatically save the user settings and display the setting label
			Settings.Default.PropertyChanged += new PropertyChangedEventHandler((object o, PropertyChangedEventArgs target) => {
				Settings.Default.Save();
				labelSettingsSaved.Visible = true;
			});

			// binds the "NetworkAvailabilityChanged" event to automatically display a notification about network connectivity, depending on the user settings
			NetworkChange.NetworkAvailabilityChanged += new NetworkAvailabilityChangedEventHandler((object o, NetworkAvailabilityEventArgs target) => {

				// checks for available networks
				if (Settings.Default.NetworkConnectivityNotification && !NetworkInterface.GetIsNetworkAvailable()) {
					notifyIcon.Icon = Resources.warning;
					notifyIcon.Text = translation.networkLost;
					notifyIcon.ShowBalloonTip(450, translation.networkLost, translation.networkConnectivityLost, ToolTipIcon.Warning);

					return;
				}

				// loops through all network interface to check network connectivity
				foreach (NetworkInterface network in NetworkInterface.GetAllNetworkInterfaces()) {

					// discards "non-up" status, modem, serial, loopback and tunnel
					if (network.OperationalStatus != OperationalStatus.Up || network.Speed < 0 || network.NetworkInterfaceType == NetworkInterfaceType.Loopback || network.NetworkInterfaceType == NetworkInterfaceType.Tunnel) {
						continue;
					}

					// discards virtual cards (like virtual box, virtual pc, etc.) and microsoft loopback adapter (showing as ethernet card)
					if (network.Name.ToLower().Contains("virtual") || network.Description.ToLower().Contains("virtual") || network.Description.ToLower() == ("microsoft loopback adapter")) {
						continue;
					}

					// syncs the inbox when a network interface is available
					this.SyncInbox();
					break;
				}

				// displays a notification to indicate that the network connectivity has been restored
				if (Settings.Default.NetworkConnectivityNotification) {
					notifyIcon.ShowBalloonTip(450, translation.networkRestored, translation.networkConnectivityRestored, ToolTipIcon.Info);
				}
			});

			// binds the "PowerModeChanged" event to automatically pause/resume the application synchronization
			SystemEvents.PowerModeChanged += new PowerModeChangedEventHandler((object o, PowerModeChangedEventArgs target) => {
				if (target.Mode == PowerModes.Suspend) {
					timer.Enabled = false;
				} else if (target.Mode == PowerModes.Resume) {
					timer.Enabled = true;
				}
			});

			// displays the step delay setting
			fieldStepDelay.SelectedIndex = Settings.Default.StepDelay;

			// displays the privacy notification setting
			switch (Settings.Default.PrivacyNotification) {
				case (int)Privacy.None:
					fieldPrivacyNotificationNone.Checked = true;
					pictureBoxPrivacyPreview.Image = Resources.privacy_none;
					break;
				default:
				case (int)Privacy.Short:
					fieldPrivacyNotificationShort.Checked = true;
					pictureBoxPrivacyPreview.Image = Resources.privacy_short;
					break;
				case (int)Privacy.All:
					fieldPrivacyNotificationAll.Checked = true;
					pictureBoxPrivacyPreview.Image = Resources.privacy_all;
					break;
			}

			// displays the update period setting
			fieldUpdatePeriod.SelectedIndex = Settings.Default.UpdatePeriod;

			// displays the update control setting
			labelUpdateControl.Text = Settings.Default.UpdateControl.ToString();

			// displays the product version
			string[] version = Application.ProductVersion.Split('.');
			linkVersion.Text = version[0] + "." + version[1] + "-" + (version[2] == "0" ? "alpha" : version[2] == "1" ? "beta" : version[2] == "2" ? "rc" : version[2] == "3" ? "release" : "") + (version[3] != "0" ? "." + version[3] : "");

			// positioning the check for update link
			linkCheckForUpdate.Left = linkVersion.Right + 2;

			// displays a tooltip for the product version
			ToolTip tipTag = new ToolTip();
			tipTag.SetToolTip(linkVersion, GITHUB_REPOSITORY + "/releases/tag/v" + linkVersion.Text.Replace(" ", "-"));
			tipTag.ToolTipTitle = translation.tipReleaseNotes;
			tipTag.ToolTipIcon = ToolTipIcon.Info;
			tipTag.IsBalloon = false;

			// displays a tooltip for the product version
			ToolTip tipCheckForUpdate = new ToolTip();
			tipCheckForUpdate.SetToolTip(linkCheckForUpdate, translation.checkForUpdate);
			tipCheckForUpdate.IsBalloon = false;

			// displays a tooltip for the website link
			ToolTip tipWebsiteYusuke = new ToolTip();
			tipWebsiteYusuke.SetToolTip(linkWebsiteYusuke, "http://p.yusukekamiyamane.com");
			tipWebsiteYusuke.IsBalloon = false;

			// displays a tooltip for the website link
			ToolTip tipWebsiteXavier = new ToolTip();
			tipWebsiteXavier.SetToolTip(linkWebsiteXavier, "http://www.xavierfoucrier.fr");
			tipWebsiteXavier.IsBalloon = false;

			// displays a tooltip for the license link
			ToolTip tipLicense = new ToolTip();
			tipLicense.SetToolTip(linkLicense, GITHUB_REPOSITORY + "/blob/master/LICENSE.md");
			tipLicense.IsBalloon = false;

			// check for update, depending on the user settings
			if (Settings.Default.UpdateService && Settings.Default.UpdatePeriod == (int)Period.Startup) {
				checkForUpdate(false);
			}
		}

		/// <summary>
		/// Prompts the user before closing the form
		/// </summary>
		private void Main_FormClosing(object sender, FormClosingEventArgs e) {

			// asks the user for exit, depending on the application settings
			if (e.CloseReason != CloseReason.ApplicationExitCall && e.CloseReason != CloseReason.WindowsShutDown && Settings.Default.AskonExit) {
				DialogResult dialog = MessageBox.Show(translation.applicationExitQuestion, translation.applicationExit, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

				if (dialog == DialogResult.No) {
					e.Cancel = true;
				}
			}

			// disposes the gmail api service
			if (this.service != null) {
				this.service.Dispose();
			}
		}

		/// <summary>
		/// Pings the 8.8.8.8 server to checks the internet connectivity
		/// </summary>
		/// <returns>Indicates if the user is connected to the internet, false means the ping to 8.8.8.8 server has failed</returns>
		private bool IsInternetAvailable() {
			try {
				Ping ping = new Ping();
				PingReply reply = ping.Send("8.8.8.8", 1000, new byte[32]);

				return reply.Status == IPStatus.Success;
			} catch (Exception) {
				return false;
			}
		}

		/// <summary>
		/// Asynchronous method used to get user credential
		/// </summary>
		private async void AsyncAuthentication() {
			try {
				
				// waits for the user authorization
				this.credential = await AsyncAuthorizationBroker();

				// creates the gmail api service
				this.service = new GmailService(new BaseClientService.Initializer() {
					HttpClientInitializer = this.credential,
					ApplicationName = "Gmail notifier for Windows"
				});

				// displays the user email address
				labelEmailAddress.Text = this.service.Users.GetProfile("me").Execute().EmailAddress;
				labelTokenDelivery.Text = this.credential.Token.Issued.ToString();
			} catch(Exception) {

				// exits the application if the google api token file doesn't exists
				if (!Directory.Exists(appdata) || !Directory.EnumerateFiles(appdata).Any()) {
					MessageBox.Show(translation.authenticationWithGmailRefused, translation.authenticationFailed, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					Application.Exit();
				}
			} finally {

				// synchronizes the user mailbox
				this.SyncInbox();
			}
		}

		/// <summary>
		/// Asynchronous task used to get the user authorization
		/// </summary>
		/// <returns>OAuth 2.0 user credential</returns>
		private async Task<UserCredential> AsyncAuthorizationBroker() {

			// uses the client secret file for the context
			using (var stream = new FileStream(Path.GetDirectoryName(Application.ExecutablePath) + "/client_secret.json", FileMode.Open, FileAccess.Read)) {

				// waits for the user validation, only if the user has not already authorized the application
				return await GoogleWebAuthorizationBroker.AuthorizeAsync(
					GoogleClientSecrets.Load(stream).Secrets,
					new string[] { GmailService.Scope.GmailModify },
					"user",
					CancellationToken.None,
					new FileDataStore(appdata, true)
				);
			}
		}

		/// <summary>
		/// Synchronizes the user inbox
		/// </summary>
		/// <param name="timertick">Indicates if the synchronization come's from the timer tick or has been manually triggered</param>
		private void SyncInbox(bool timertick = false) {

			// if internet is down, attempts to reconnect the user mailbox
			if (!this.IsInternetAvailable()) {
				timerReconnect.Enabled = true;
				timer.Enabled = false;

				return;
			}

			// activates the necessary menu items
			menuItemSynchronize.Enabled = true;
			menuItemTimout.Enabled = true;
			menuItemSettings.Enabled = true;

			// disables the timeout when the user do a manual synchronization
			if (timer.Interval != Settings.Default.TimerInterval) {
				Timeout(menuItemTimeoutDisabled, Settings.Default.TimerInterval);

				// exits the method because the timeout function automatically restarts a synchronization once it has been disabled
				return;
			}

			// displays the sync icon, but not on every tick of the timer
			if (!timertick) {
				notifyIcon.Icon = Resources.sync;
				notifyIcon.Text = translation.sync;
			}

			// check for update, depending on the user settings
			if (Settings.Default.UpdateService) {
				switch (Settings.Default.UpdatePeriod) {
					case (int)Period.Day:
						if (DateTime.Now >= Settings.Default.UpdateControl.AddDays(1)) {
							checkForUpdate(false);
						}

						break;
					case (int)Period.Week:
						if (DateTime.Now >= Settings.Default.UpdateControl.AddDays(7)) {
							checkForUpdate(false);
						}

						break;
					case (int)Period.Month:
						if (DateTime.Now >= Settings.Default.UpdateControl.AddMonths(1)) {
							checkForUpdate(false);
						}

						break;
				}
			}

			try {

				// manages the spam notification
				if (Settings.Default.SpamNotification) {

					// gets the "spam" label
					Google.Apis.Gmail.v1.Data.Label spam = this.service.Users.Labels.Get("me", "SPAM").Execute();

					// manages unread spams
					if (spam.ThreadsUnread > 0) {

						// sets the notification icon and text
						notifyIcon.Icon = Resources.spam;

						// plays a sound on unread spams
						if (Settings.Default.AudioNotification) {
							SystemSounds.Exclamation.Play();
						}

						// displays a balloon tip in the systray with the total of unread threads
						notifyIcon.ShowBalloonTip(450, spam.ThreadsUnread.ToString() + " " + (spam.ThreadsUnread > 1 ? translation.unreadSpams : translation.unreadSpam), translation.newUnreadSpam, ToolTipIcon.Error);
						notifyIcon.Text = spam.ThreadsUnread.ToString() + " " + (spam.ThreadsUnread > 1 ? translation.unreadSpams : translation.unreadSpam);
						notifyIcon.Tag = "#spam";

						return;
					}
				}

				// gets the "inbox" label
				this.inbox = this.service.Users.Labels.Get("me", "INBOX").Execute();

				// exits the sync if the number of unread threads is the same as before
				if (timertick && (this.inbox.ThreadsUnread == this.unreadthreads)) {
					return;
				}

				// manages unread threads
				if (this.inbox.ThreadsUnread > 0) {

					// sets the notification icon and text
					notifyIcon.Icon = this.inbox.ThreadsUnread <= 9 ? (Icon)Resources.ResourceManager.GetObject("mail_" + this.inbox.ThreadsUnread.ToString()) : Resources.mail_more;

					// manages message notification
					if (Settings.Default.MessageNotification) {

						// plays a sound on unread threads
						if (Settings.Default.AudioNotification) {
							SystemSounds.Asterisk.Play();
						}

						// gets the message details
						UsersResource.MessagesResource.ListRequest messages = this.service.Users.Messages.List("me");
						messages.LabelIds = "UNREAD";
						messages.MaxResults = 1;
						Google.Apis.Gmail.v1.Data.Message message = this.service.Users.Messages.Get("me", messages.Execute().Messages.First().Id).Execute();

						//  displays a balloon tip in the systray with the total of unread threads and message details, depending on the user privacy setting
						if (this.inbox.ThreadsUnread == 1 && Settings.Default.PrivacyNotification != (int)Privacy.All) {
							string subject = "";
							string from = "";

							foreach (MessagePartHeader header in message.Payload.Headers) {
								if (header.Name == "Subject") {
									subject = header.Value != "" ? header.Value : translation.newUnreadMessage;
								} else if (header.Name == "From") {
									Match match = Regex.Match(header.Value, ".* <");

									if (match.Length != 0) {
										from = match.Captures[0].Value.Replace(" <", "").Replace("\"", "");
									} else {
										match = Regex.Match(header.Value, "<?.*>?");
										from = match.Length != 0 ? match.Value.ToLower().Replace("<", "").Replace(">", "") : header.Value.Replace(match.Value, this.inbox.ThreadsUnread.ToString() + " " + translation.unreadMessage);
									}
								}
							}

							if (Settings.Default.PrivacyNotification == (int)Privacy.None) {
								notifyIcon.ShowBalloonTip(450, from, message.Snippet != "" ? WebUtility.HtmlDecode(message.Snippet) : translation.newUnreadMessage, ToolTipIcon.Info);
							} else if (Settings.Default.PrivacyNotification == (int)Privacy.Short) {
								notifyIcon.ShowBalloonTip(450, from, subject, ToolTipIcon.Info);
							}
						} else {
							notifyIcon.ShowBalloonTip(450, this.inbox.ThreadsUnread.ToString() + " " + (this.inbox.ThreadsUnread > 1 ? translation.unreadMessages : translation.unreadMessage), translation.newUnreadMessage, ToolTipIcon.Info);
						}

						// manages the balloon tip click event handler: we store the "notification tag" to allow the user to directly display the specified view (inbox/message/spam) in a browser
						notifyIcon.Tag = "#inbox" + (this.inbox.ThreadsUnread == 1 ? "/" + message.Id : "");
					}

					// displays the notification text
					notifyIcon.Text = this.inbox.ThreadsUnread.ToString() + " " + (this.inbox.ThreadsUnread > 1 ? translation.unreadMessages : translation.unreadMessage);

					// enables the mark as read menu item
					menuItemMarkAsRead.Enabled = true;
				} else {

					// restores the default systray icon and text
					notifyIcon.Icon = Resources.normal;
					notifyIcon.Text = translation.noMessage;

					// disables the mark as read menu item
					menuItemMarkAsRead.Enabled = false;
				}

				// saves the number of unread threads
				this.unreadthreads = this.inbox.ThreadsUnread;
			} catch(Exception exception) {

				// displays a balloon tip in the systray with the detailed error message
				notifyIcon.Icon = Resources.warning;
				notifyIcon.Text = translation.syncError;
				notifyIcon.ShowBalloonTip(1500, translation.error, translation.syncErrorOccured + exception.Message, ToolTipIcon.Warning);
			}
		}

		/// <summary>
		/// Manages the RunAtWindowsStartup user setting
		/// </summary>
		private void fieldRunAtWindowsStartup_CheckedChanged(object sender, EventArgs e) {
			if (fieldRunAtWindowsStartup.Checked) {
				using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true)) {
					key.SetValue("Gmail notifier", '"' + Application.ExecutablePath + '"');
				}
			} else {
				using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true)) {
					key.DeleteValue("Gmail notifier", false);
				}
			}
		}

		/// <summary>
		/// Manages the Language user setting
		/// </summary>
		private void fieldLanguage_SelectionChangeCommitted(object sender, EventArgs e) {
			Settings.Default.Language = fieldLanguage.Text;

			// indicates to the user that to apply the new language on the interface, the application must be restarted
			labelRestartToApply.Visible = true;
			linkRestartToApply.Visible = true;
		}

		/// <summary>
		/// Manages the SpamNotification user setting
		/// </summary>
		private void fieldSpamNotification_Click(object sender, EventArgs e) {
			this.SyncInbox();
		}

		/// <summary>
		/// Manages the NumericDelay user setting
		/// </summary>
		private void fieldNumericDelay_ValueChanged(object sender, EventArgs e) {
			Settings.Default.TimerInterval = 1000 * (fieldStepDelay.SelectedIndex == 0 ? 60 : 3600) * Convert.ToInt32(fieldNumericDelay.Value);
			Settings.Default.NumericDelay = fieldNumericDelay.Value;
		}

		/// <summary>
		/// Manages the StepDelay user setting
		/// </summary>
		private void fieldStepDelay_SelectionChangeCommitted(object sender, EventArgs e) {
			Settings.Default.TimerInterval = 1000 * (fieldStepDelay.SelectedIndex == 0 ? 60 : 3600) * Convert.ToInt32(fieldNumericDelay.Value);
			Settings.Default.StepDelay = fieldStepDelay.SelectedIndex;
		}

		/// <summary>
		/// Manages the PrivacyNotificationNone user setting
		/// </summary>
		private void fieldPrivacyNotificationNone_CheckedChanged(object sender, EventArgs e) {
			Settings.Default.PrivacyNotification = (int)Privacy.None;
			pictureBoxPrivacyPreview.Image = Resources.privacy_none;
		}

		/// <summary>
		/// Manages the PrivacyNotificationShort user setting
		/// </summary>
		private void fieldPrivacyNotificationShort_CheckedChanged(object sender, EventArgs e) {
			Settings.Default.PrivacyNotification = (int)Privacy.Short;
			pictureBoxPrivacyPreview.Image = Resources.privacy_short;
		}

		/// <summary>
		/// Manages the PrivacyNotificationAll user setting
		/// </summary>
		private void fieldPrivacyNotificationAll_CheckedChanged(object sender, EventArgs e) {
			Settings.Default.PrivacyNotification = (int)Privacy.All;
			pictureBoxPrivacyPreview.Image = Resources.privacy_all;
		}

		/// <summary>
		/// Manages the UpdatePeriod user setting
		/// </summary>
		private void fieldUpdatePeriod_SelectedIndexChanged(object sender, EventArgs e) {
			Settings.Default.UpdatePeriod = fieldUpdatePeriod.SelectedIndex;
		}

		/// <summary>
		/// Opens the Github release section of the current build
		/// </summary>
		private void linkVersion_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
			Process.Start(GITHUB_REPOSITORY + "/releases/tag/v" + linkVersion.Text.Replace(" ", "-"));
		}

		/// <summary>
		/// Opens the Yusuke website
		/// </summary>
		private void linkWebsiteYusuke_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
			Process.Start("http://p.yusukekamiyamane.com");
		}

		/// <summary>
		/// Opens the Xavier website
		/// </summary>
		private void linkWebsiteXavier_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
			Process.Start("http://www.xavierfoucrier.fr");
		}

		/// <summary>
		/// Opens the Github license file
		/// </summary>
		private void linkLicense_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
			Process.Start(GITHUB_REPOSITORY + "/blob/master/LICENSE.md");
		}

		/// <summary>
		/// Hides the settings saved label
		/// </summary>
		private void tabControl_Selecting(object sender, TabControlCancelEventArgs e) {
			labelSettingsSaved.Visible = false;
		}

		/// <summary>
		/// Closes the preferences when the OK button is clicked
		/// </summary>
		private void buttonOK_Click(object sender, EventArgs e) {
			labelSettingsSaved.Visible = false;
			WindowState = FormWindowState.Minimized;
			ShowInTaskbar = false;
			Visible = false;
		}

		/// <summary>
		/// Manages the context menu Synchronize item
		/// </summary>
		private void menuItemSynchronize_Click(object sender, EventArgs e) {
			this.SyncInbox();
		}

		/// <summary>
		/// Manages the context menu MarkAsRead item
		/// </summary>
		private void menuItemMarkAsRead_Click(object sender, EventArgs e) {
			try {

				// displays the sync icon
				notifyIcon.Icon = Resources.sync;
				notifyIcon.Text = translation.sync;

				// gets all unread threads
				UsersResource.ThreadsResource.ListRequest threads = this.service.Users.Threads.List("me");
				threads.LabelIds = "UNREAD";
				IList<Google.Apis.Gmail.v1.Data.Thread> unread = threads.Execute().Threads;

				// loops through all unread threads and removes the "unread" label for each one
				foreach (Google.Apis.Gmail.v1.Data.Thread thread in unread) {
					ModifyThreadRequest request = new ModifyThreadRequest();
					request.RemoveLabelIds = new List<string>() { "UNREAD" };
					this.service.Users.Threads.Modify(request, "me", thread.Id).Execute();
				}

				// restores the default systray icon and text
				notifyIcon.Icon = Resources.normal;
				notifyIcon.Text = translation.noMessage;

				// restores the default tag
				notifyIcon.Tag = null;

				// disables the mark as read menu item
				menuItemMarkAsRead.Enabled = false;
			} catch (Exception exception) {

				// enabled the mark as read menu item
				menuItemMarkAsRead.Enabled = true;

				// displays a balloon tip in the systray with the detailed error message
				notifyIcon.Icon = Resources.warning;
				notifyIcon.Text = translation.markAsReadError;
				notifyIcon.ShowBalloonTip(1500, translation.error, translation.markAsReadErrorOccured + exception.Message, ToolTipIcon.Warning);
			}
		}

		/// <summary>
		/// Delays the inbox sync during a certain time
		/// </summary>
		/// <param name="item">Item selected in the menu</param>
		/// <param name="delay">Delay until the next inbox sync</param>
		private void Timeout(MenuItem item, int delay) {

			// exits if the selected menu item is already checked
			if (item.Checked) {
				return;
			}

			// unchecks all menu items
			foreach (MenuItem i in menuItemTimout.MenuItems) {
				i.Checked = false;
			}

			// displays the user choice
			item.Checked = true;

			// sets the new timer interval depending on the user settings
			timer.Interval = delay;

			// restores the default tag
			notifyIcon.Tag = null;

			// updates the systray icon and text
			if (delay != Settings.Default.TimerInterval) {
				notifyIcon.Icon = Resources.timeout;
				notifyIcon.Text = translation.timeout + " - " + DateTime.Now.AddMilliseconds(delay).ToShortTimeString();
			} else {
				this.SyncInbox();
			}
		}

		/// <summary>
		/// Manages the context menu TimeoutDisabled item
		/// </summary>
		private void menuItemTimeoutDisabled_Click(object sender, EventArgs e) {
			Timeout((MenuItem)sender, Settings.Default.TimerInterval);
		}

		/// <summary>
		/// Manages the context menu Timeout30m item
		/// </summary>
		private void menuItemTimeout30m_Click(object sender, EventArgs e) {
			Timeout((MenuItem)sender, 1000 * 60 * 30);
		}

		/// <summary>
		/// Manages the context menu Timeout1h item
		/// </summary>
		private void menuItemTimeout1h_Click(object sender, EventArgs e) {
			Timeout((MenuItem)sender, 1000 * 60 * 60);
		}

		/// <summary>
		/// Manages the context menu Timeout2h item
		/// </summary>
		private void menuItemTimeout2h_Click(object sender, EventArgs e) {
			Timeout((MenuItem)sender, 1000 * 60 * 60 * 2);
		}

		/// <summary>
		/// Manages the context menu Timeout5h item
		/// </summary>
		private void menuItemTimeout5h_Click(object sender, EventArgs e) {
			Timeout((MenuItem)sender, 1000 * 60 * 60 * 5);
		}

		/// <summary>
		/// Manages the context menu Settings item
		/// </summary>
		private void menuItemSettings_Click(object sender, EventArgs e) {
			Visible = true;
			ShowInTaskbar = true;
			WindowState = FormWindowState.Normal;
		}

		/// <summary>
		/// Manages the context menu exit item
		/// </summary>
		private void menuItemExit_Click(object sender, EventArgs e) {
			this.Close();
		}

		/// <summary>
		/// Manages the systray icon double click
		/// </summary>
		private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e) {
			if (e.Button == MouseButtons.Left) {

				// by default, always open the gmail inbox in a browser
				if (notifyIcon.Tag == null) {
					Process.Start(GMAIL_BASEURL + "/#inbox");
				} else {
					notifyIconInteraction();
				}
			}
		}

		/// <summary>
		/// Manages the systray icon balloon click
		/// </summary>
		private void notifyIcon_BalloonTipClicked(object sender, EventArgs e) {
			if ((Control.MouseButtons & MouseButtons.Right) == MouseButtons.Right) {
				return;
			}

			notifyIconInteraction();
		}

		/// <summary>
		/// Opens the gmail specified view (inbox/message/spam) in a browser
		/// </summary>
		private void notifyIconInteraction() {
			if (notifyIcon.Tag == null) {
				return;
			}

			// opens a browser
			Process.Start(GMAIL_BASEURL + "/" + notifyIcon.Tag);
			notifyIcon.Tag = null;

			// restores the default systray icon and text: pretends that the user had read all his mail, except if the timeout option is activated
			if (timer.Interval == Settings.Default.TimerInterval) {
				notifyIcon.Icon = Resources.normal;
				notifyIcon.Text = translation.noMessage;

				// disables the mark as read menu item
				menuItemMarkAsRead.Enabled = false;
			}
		}

		/// <summary>
		/// Synchronizes the user mailbox on every timer tick
		/// </summary>
		private void timer_Tick(object sender, EventArgs e) {

			// restores the timer interval when the timeout time has elapsed
			if (timer.Interval != Settings.Default.TimerInterval) {
				Timeout(menuItemTimeoutDisabled, Settings.Default.TimerInterval);

				return;
			}

			// synchronizes the inbox
			this.SyncInbox(true);
		}

		/// <summary>
		/// Disconnects the Gmail account from the application
		/// </summary>
		private void buttonGmailDisconnect_Click(object sender, EventArgs e) {

			// asks the user for disconnect
			DialogResult dialog = MessageBox.Show(translation.gmailDisconnectQuestion.Replace("{account_name}", labelEmailAddress.Text), translation.gmailDisconnect, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

			if (dialog == DialogResult.No) {
				return;
			}

			// deletes the local application data folder and the client token file
			if (Directory.Exists(appdata)) {
				Directory.Delete(appdata, true);
			}

			// restarts the application
			this.restart();
		}

		/// <summary>
		/// Restarts the application to apply new user settings
		/// </summary>
		private void linkRestartToApply_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
			this.restart();
		}

		/// <summary>
		/// Restarts the application
		/// </summary>
		private void restart() {

			// starts a new process
			ProcessStartInfo command = new ProcessStartInfo("cmd.exe", "/C ping 127.0.0.1 -n 2 && \"" + Application.ExecutablePath + "\"");
			command.WindowStyle = ProcessWindowStyle.Hidden;
			command.CreateNoWindow = true;
			Process.Start(command);

			// exits the application
			Application.Exit();
		}

		// attempts to reconnect the user mailbox
		private void timerReconnect_Tick(object sender, EventArgs e) {

			// sets the reconnection interval
			timerReconnect.Interval = INTERVAL_RECONNECT * 1000;

			// checks internet connectivity
			if (!this.IsInternetAvailable()) {

				// disables the menu items
				menuItemSynchronize.Enabled = false;
				menuItemTimout.Enabled = false;
				menuItemSettings.Enabled = false;

				// increases the number of reconnection attempt
				this.reconnect++;

				// displays a balloon tip one time in the systray with the detailed reconnection message, depending on the user notification setting
				if (this.reconnect == 1 && Settings.Default.AttemptToReconnectNotification) {
					notifyIcon.Icon = Resources.warning;
					notifyIcon.Text = translation.reconnectAttempt;
					notifyIcon.ShowBalloonTip(450, translation.reconnectAttempt, translation.reconnectNow, ToolTipIcon.Warning);
				}

				// after max unsuccessull reconnection attempts, the application waits for the next sync
				if (reconnect == MAX_AUTO_RECONNECT) {
					timerReconnect.Enabled = false;
					timer.Enabled = true;
					this.reconnect = 1;

					// displays a balloon tip in the systray with the last detailed reconnection message, depending on the user notification setting
					if (Settings.Default.AttemptToReconnectNotification) {
						notifyIcon.Icon = Resources.warning;
						notifyIcon.Text = translation.reconnectFailed;
						notifyIcon.ShowBalloonTip(450, translation.reconnectFailed, translation.reconnectNextTime, ToolTipIcon.Warning);
					}
				}
			} else {

				// restores default operation when internet is available
				timerReconnect.Enabled = false;
				timer.Enabled = true;

				// resets reconnection count
				this.reconnect = 0;

				// syncs the user mailbox
				this.SyncInbox();
			}
		}

		/// <summary>
		/// Check for update
		/// </summary>
		private void buttonCheckForUpdate_Click(object sender, EventArgs e) {
			checkForUpdate();
		}

		/// <summary>
		/// Check for update
		/// </summary>
		private void linkCheckForUpdate_Click(object sender, EventArgs e) {
			linkCheckForUpdate.Image = Resources.update_hourglass;
			linkCheckForUpdate.Enabled = false;
			Cursor.Current = DefaultCursor;
			checkForUpdate();
			linkCheckForUpdate.Enabled = true;
			linkCheckForUpdate.Image = Resources.update_check;
		}

		/// <summary>
		/// Connect to the repository and check if there is an update available
		/// </summary>
		private void checkForUpdate(bool verbose = true) {
			try {

				// gets the list of tags in the Github repository tags webpage
				HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlWeb().Load(GITHUB_REPOSITORY + "/tags");
				HtmlAgilityPack.HtmlNodeCollection collection = document.DocumentNode.SelectNodes("//span[@class='tag-name']");

				if (collection != null && collection.Count > 0) {
					List<string> tags = collection.Select(p => p.InnerText).ToList();

					// the current version tag is not at the top of the list
					if (tags.IndexOf("v" + linkVersion.Text.Replace(" ", "-")) > 0) {
						DialogResult dialog = MessageBox.Show(translation.newVersion.Replace("{version}", tags[0]), "Gmail Notifier Update", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);

						// redirects the user to the Github repository releases webpage
						if (dialog == DialogResult.Yes) {
							Process.Start(GITHUB_REPOSITORY + "/releases/" + tags[0]);
						}
					} else {
						if (verbose) {
							MessageBox.Show(translation.latestVersion, "Gmail Notifier Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
					}
				}
			} catch (Exception) {
				// nothing to catch
			}

			// stores the latest update datetime control
			Settings.Default.UpdateControl = DateTime.Now;

			// updates the update control label
			labelUpdateControl.Text = Settings.Default.UpdateControl.ToString();
		}
	}
}
