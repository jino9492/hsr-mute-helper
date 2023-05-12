using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using NAudio.CoreAudioApi;
using Newtonsoft.Json;


namespace HSR_MUTE_HELPER
{

    public partial class Mixer
    {
        #region Find App
        public static AudioSessionControl GetTargetProgram(){
            string settingJson = File.ReadAllText(Application.StartupPath + "/setting.json");
            dynamic jsonObject = JsonConvert.DeserializeObject(settingJson);

            MMDeviceEnumerator enumerator = new MMDeviceEnumerator();
            MMDevice defaultDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            SessionCollection sessionEnumerator = defaultDevice.AudioSessionManager.Sessions;

            for (int i = 0; i < sessionEnumerator.Count; i++)
            {
                if (Process.GetProcessById((int)sessionEnumerator[i].GetProcessID) != null 
                    && Process.GetProcessById((int)sessionEnumerator[i].GetProcessID).ProcessName.Equals((string)jsonObject["program"], StringComparison.OrdinalIgnoreCase))
                {
                    return sessionEnumerator[i];
                }
            }

            return null;
        }
        #endregion

        #region Mute App Functions
        public static AudioSessionControl target { get; set; } = GetTargetProgram();

        public static void MuteApplication(string appName)
        {
            target.SimpleAudioVolume.Mute = true;
            HSR_MUTE_HELPER.MUTE_BUTTON.BackgroundImage = HSR_MUTE_HELPER.GetImage(
                Application.StartupPath + "/Resources/" + "sound2.png",
                HSR_MUTE_HELPER.buttonWidth, HSR_MUTE_HELPER.buttonHeight);
        }

        public static void UnMuteApplication(string appName)
        {
            target.SimpleAudioVolume.Mute = false;
            HSR_MUTE_HELPER.MUTE_BUTTON.BackgroundImage = HSR_MUTE_HELPER.GetImage(
                Application.StartupPath + "/Resources/" + "sound1.png", 
                HSR_MUTE_HELPER.buttonWidth, HSR_MUTE_HELPER.buttonHeight);
        }
        #endregion
    }

    public partial class HSR_MUTE_HELPER : Form
    {
        public static Button MUTE_BUTTON { get; set; } = new Button();
        public static int buttonWidth { get; set; } = 50;
        public static int buttonHeight { get; set; } = 50;

        public static PictureBox MAIN_IMAGE = new PictureBox();
        public static int imageWidth = 100;
        public static int imageHeight = 100;
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x80;
                return cp;
            }
        }

        public HSR_MUTE_HELPER()
        {
            InitializeComponent();
            CheckTargetState();

            string settingJson = File.ReadAllText(Application.StartupPath + "/setting.json");
            dynamic jsonObject = JsonConvert.DeserializeObject(settingJson);

            #region MUTE_BUTTON
            MUTE_BUTTON.Size = new Size(buttonWidth, buttonHeight);
            MUTE_BUTTON.Location = new Point(
                ((this.ClientSize.Width - MUTE_BUTTON.Size.Width) / 2),
                ((this.ClientSize.Height - MUTE_BUTTON.Size.Height) / 2)
                );

            MUTE_BUTTON.Click += (sender, e) => MuteControl(sender, e, (string)jsonObject["program"]);

            MUTE_BUTTON.FlatStyle = FlatStyle.Flat;
            MUTE_BUTTON.BackColor = Color.White;

            if (Mixer.target != null)
            {
                if (!Mixer.target.SimpleAudioVolume.Mute){
                    MUTE_BUTTON.BackgroundImage = HSR_MUTE_HELPER.GetImage(
                        Application.StartupPath + "/Resources/" + "sound1.png",
                        HSR_MUTE_HELPER.buttonWidth, HSR_MUTE_HELPER.buttonHeight);
                    MUTE_BUTTON.BackgroundImageLayout = ImageLayout.Center;
                }
                else{
                    MUTE_BUTTON.BackgroundImage = HSR_MUTE_HELPER.GetImage(
                        Application.StartupPath + "/Resources/" + "sound2.png",
                        HSR_MUTE_HELPER.buttonWidth, HSR_MUTE_HELPER.buttonHeight);
                    MUTE_BUTTON.BackgroundImageLayout = ImageLayout.Center;
                }
            }
            #endregion

            #region Image
            MAIN_IMAGE.Size = new Size(imageWidth, imageHeight);

            int aj = 5;
            MAIN_IMAGE.Location = new Point(
                ((this.ClientSize.Width - MAIN_IMAGE.Size.Width) / 2) + aj,
                ((this.ClientSize.Height - MAIN_IMAGE.Size.Height) / 2) - 100
                );

            MAIN_IMAGE.MouseDown += (sender, e) => ImageMouseDown(sender, e);
            MAIN_IMAGE.MouseUp   += (sender, e) => ImageMouseUp  (sender, e);
            MAIN_IMAGE.MouseMove += (sender, e) => ImageMouseMove(sender, e);

            MAIN_IMAGE.Image = GetImage(
                Application.StartupPath + "/Resources/" + (string)jsonObject["filename"], 
                imageWidth, imageHeight);
            #endregion


            this.BackColor = Color.Magenta;
            this.TransparencyKey = Color.Magenta;
            this.FormBorderStyle = FormBorderStyle.None;
            this.TopMost = true;
            this.Controls.Add(MAIN_IMAGE);
            this.Controls.Add(MUTE_BUTTON);
            this.ShowInTaskbar = false;
        }

        #region Target Program State Check Function
        private Timer timer;
        public void CheckTargetState()
        {
            
            timer = new Timer();
            timer.Interval = 1000;
            timer.Tick += new EventHandler((sender, e) => {
                if(Mixer.target == null
                || Mixer.target.State == NAudio.CoreAudioApi.Interfaces.AudioSessionState.AudioSessionStateExpired)
                {
                    Application.Exit();
                }
            });
            timer.Start();
        }
        #endregion

        #region Mute Function
        public void MuteControl(object sender, EventArgs e, string programName)
        {
            if (Mixer.target != null)
            {
                if (!Mixer.target.SimpleAudioVolume.Mute)
                    Mixer.MuteApplication(programName);
                else
                    Mixer.UnMuteApplication(programName);
            }
        }
        #endregion

        #region Drag Functions
        private bool isDragging = false;
        private Point lastLocation;
        private void ImageMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = true;
                lastLocation = e.Location;
            }
        }

        private void ImageMouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                this.Location = new Point(
                    (this.Location.X - lastLocation.X) + e.X,
                    (this.Location.Y - lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void ImageMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                Application.Exit();
            }

            isDragging = false;
        }
        #endregion

        #region Load Image
        public static Bitmap GetImage(string url, int width, int height)
        {
            Bitmap icon = new Bitmap(url);
            Bitmap ICON_resize = new Bitmap(icon, new Size(width - 10, height - 10));

            return ICON_resize;
        }
        #endregion
    }
}
