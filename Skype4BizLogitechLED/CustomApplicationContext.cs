using LedCSharp;
using Microsoft.Lync.Model;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace Skype4BizLogitechLED
{
    /// <summary>
    /// From https://www.simple-talk.com/dotnet/.net-framework/creating-tray-applications-in-.net-a-practical-guide/
    /// </summary>
    public class CustomApplicationContext : ApplicationContext
    {
        private static readonly string IconFileName = "led.ico";
        private static readonly string DefaultTooltip = "Skype For Business Logitech LED Integration";
        private NotifyIcon notifyIcon;
        private Container components;
        public LyncClient lc;
        public bool firstRun = true;

        /// <summary>
        /// This class should be created and passed into Application.Run( ... )
        /// </summary>
        public CustomApplicationContext()
        {
            InitializeContext();
        }

        private void InitializeContext()
        {
            components = new System.ComponentModel.Container();
            notifyIcon = new NotifyIcon(components)
            {
                ContextMenuStrip = new ContextMenuStrip(),
                Icon = new Icon(IconFileName),
                Text = DefaultTooltip,
                Visible = true
            };
            notifyIcon.ContextMenuStrip.Opening += ContextMenuStrip_Opening;

            SetupLyncLogitechLED();
            //notifyIcon.DoubleClick += notifyIcon_DoubleClick;
            //notifyIcon.MouseUp += notifyIcon_MouseUp;
        }

        private void ContextMenuStrip_Opening(object sender, CancelEventArgs e)
        {
            e.Cancel = false;

            notifyIcon.ContextMenuStrip.Items.Add("Exit", null, ContextMenu_ExitClicked);
        }

        private void ContextMenu_ExitClicked(object sender, EventArgs e)
        {
            ExitThread();
        }

        private void SetupLyncLogitechLED()
        {
            lc = Microsoft.Lync.Model.LyncClient.GetClient();
            if (lc.State == Microsoft.Lync.Model.ClientState.SignedIn)
            {

                lc.Self.Contact.ContactInformationChanged += Contact_ContactInformationChanged;

                ContactAvailability availability = (ContactAvailability)lc.Self.Contact.GetContactInformation(ContactInformationType.Availability);
                ChangeColor(availability);

            }
        }

        private void Contact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            if (lc.State == ClientState.SigningOut || lc.State == ClientState.SignedOut || lc.State == ClientState.SigningIn)
            {
                LogitechGSDK.LogiLedRestoreLighting();
                return;
            }

            Contact self = sender as Contact;
            if (e.ChangedContactInformation.Contains(ContactInformationType.Availability))
            {
                ContactAvailability availability = (ContactAvailability)self.GetContactInformation(ContactInformationType.Availability);

                ChangeColor(availability);
            }
        }

        private void ChangeColor(ContactAvailability availability)
        {
            LogitechGSDK.LogiLedInit();
            if (firstRun)
            {
                LogitechGSDK.LogiLedSaveCurrentLighting();
                firstRun = false;
            }

            switch (availability)
            {
                case ContactAvailability.None:
                    break;
                case ContactAvailability.Free:
                case ContactAvailability.FreeIdle:
                    LogitechGSDK.LogiLedSetLighting(0, 100, 0);
                    break;
                case ContactAvailability.Busy:
                case ContactAvailability.BusyIdle:
                case ContactAvailability.DoNotDisturb:
                    LogitechGSDK.LogiLedSetLighting(100, 0, 0);
                    //LogitechGSDK.LogiLedPulseLighting(100, 0, 0, 10000, 500);
                    break;
                case ContactAvailability.TemporarilyAway:
                case ContactAvailability.Away:
                    LogitechGSDK.LogiLedSetLighting(98, 86, 5);
                    break;
                case ContactAvailability.Offline:
                case ContactAvailability.Invalid:
                default:
                    LogitechGSDK.LogiLedSetLighting(0, 100, 0);
                    break;
            }
        }

        /// <summary>
		/// When the application context is disposed, dispose things like the notify icon.
		/// </summary>
		/// <param name="disposing"></param>
		protected override void Dispose(bool disposing)
        {
            if (disposing && components != null) { components.Dispose(); }

            if (!firstRun)
            {
                LogitechGSDK.LogiLedRestoreLighting();
                LogitechGSDK.LogiLedShutdown();

                lc.Self.Contact.ContactInformationChanged -= Contact_ContactInformationChanged;
                lc = null;
            }
        }
        protected override void ExitThreadCore()
        {
            notifyIcon.Visible = false;

            base.ExitThreadCore();
        }
    }
}
