using System;
using System.Windows.Forms;

namespace CSharpFlexGrid
{
    public partial class SplashScreen : Form
    {
        private System.Windows.Forms.Timer splashTimer;
        private const int SPLASH_DISPLAY_TIME = 3000; // 3 seconds

        public SplashScreen()
        {
            InitializeComponent();
            InitializeSplashTimer();
        }

        private void InitializeSplashTimer()
        {
            splashTimer = new System.Windows.Forms.Timer();
            splashTimer.Interval = SPLASH_DISPLAY_TIME;
            splashTimer.Tick += SplashTimer_Tick;
            splashTimer.Start();
        }

        private void SplashTimer_Tick(object sender, EventArgs e)
        {
            splashTimer.Stop();
            splashTimer.Dispose();

            // Show the main menu form
            MainMenuForm mainForm = new MainMenuForm();
            mainForm.Show();

            // Hide the splash screen
            this.Hide();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if (splashTimer != null)
            {
                splashTimer.Stop();
                splashTimer.Dispose();
            }
        }
    }
}
