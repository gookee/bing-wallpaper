using Microsoft.Win32;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Linq;
using BingWallpaper.Helper;
using System.Threading;

namespace BingWallpaper
{
    public partial class MainForm : Form, IWallpaperControl
    {
        private BingImageProvider _provider;
        private Settings _settings;
        private HistoryImage _currentWallpaper;
        string CURRENT_FILE_CACHE = "current.img";


        System.Timers.Timer autoChangeTimer;

        #region 控制Interface

        public HistoryImage CurrentWallpaper
        {
            get { return _currentWallpaper; }

            set
            {
                this._currentWallpaper = value;
                this.WallpaperChange(value);
                this.SaveState();
            }
        }

        public delegate void WallpaperChangeHandler(HistoryImage paper);

        public event WallpaperChangeHandler OnWallpaperChange;

        public void WallpaperChange(HistoryImage paper)
        {
            OnWallpaperChange?.Invoke(paper);
        }

        public async void NextWallpaper()
        {
            if (CurrentWallpaper != null)
            {
                var newImage = HistoryImageProvider.Next(this.CurrentWallpaper.Date);
                if (newImage != null)
                {
                    this.CurrentWallpaper = newImage;
                    await UpdateWallpaper();
                }
                else
                {
                    this.CurrentWallpaper = HistoryImageProvider.First();
                    await UpdateWallpaper();
                }
            }
            else
            {
                RandomWallpaper();
            }

        }

        public async void PreWallpaper()
        {
            if (CurrentWallpaper != null)
            {
                var newImage = HistoryImageProvider.Previous(this.CurrentWallpaper.Date);
                if (newImage != null)
                {
                    this.CurrentWallpaper = newImage;
                    await UpdateWallpaper();
                }
                else
                {
                    this.CurrentWallpaper = HistoryImageProvider.Last();
                    await UpdateWallpaper();
                }
            }
            else
            {
                RandomWallpaper();
            }
        }

        public void RandomWallpaper()
        {
            this.SetRandomWallpaper();
        }
        #endregion

        private void ReloadState()
        {
            var image = HistoryImage.LoadFromFile(CURRENT_FILE_CACHE);
            if (image != null)
            {
                this.CurrentWallpaper = image;
            }
        }

        private void SaveState()
        {
            if (CurrentWallpaper != null)
            {
                this.CurrentWallpaper.SaveToFile(CURRENT_FILE_CACHE);
            }
        }

        public MainForm(BingImageProvider provider, Settings settings)
        {
            if (provider == null)
                throw new ArgumentNullException(nameof(provider));
            _provider = provider;

            if (settings == null)
                throw new ArgumentNullException(nameof(settings));
            _settings = settings;

            SetStartup(_settings.LaunchOnStartup);

            InitializeComponent();

            this.KeyPreview = true;
            //添加热键
            HotKey.RegisterHotKey(this.Handle, 110, HotKey.KeyModifiers.Shift, Keys.F5);
            HotKey.RegisterHotKey(this.Handle, 111, HotKey.KeyModifiers.Shift, Keys.Left);
            HotKey.RegisterHotKey(this.Handle, 112, HotKey.KeyModifiers.Shift, Keys.Right);

            AddTrayIcons();

            // 定时更新
            var timer = new System.Timers.Timer();
            timer.Interval = 1000 * 60 * 60 * 24;
            timer.AutoReset = true;
            timer.Enabled = true;
            timer.Elapsed += (s, e) => GetLatestWallpaper();
            timer.Start();

            if (_settings.AutoChange)
            {
                CreateAutoChangeTask();
            }

            this.ReloadState();

            new Thread(() =>
            {
                GetLatestWallpaper();

            }).Start();

            new Thread(() =>
            {
                UpdateLatestDaysImage();
            }).Start();
        }

        private void CreateAutoChangeTask()
        {
            autoChangeTimer = new System.Timers.Timer();
            autoChangeTimer.Interval = getChangeInterval();
            autoChangeTimer.AutoReset = true;
            autoChangeTimer.Enabled = true;
            autoChangeTimer.Elapsed += (s, e) => SetRandomWallpaper();
            autoChangeTimer.Start();
        }

        private int getChangeInterval()
        {
            if (_settings.AutoChangeInterval.Contains(Resource.Minutes))
            {
                return 1000 * 60 * int.Parse(_settings.AutoChangeInterval.Replace(Resource.Minutes, "").Trim());
            }
            return 1000 * 60 * 60 * int.Parse(_settings.AutoChangeInterval.Replace(Resource.Hours, "").Trim());
        }

        public void SetStartup(bool launch)
        {
            RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            if (launch)
            {
                if (rk.GetValue("BingWallpaper") == null)
                    rk.SetValue("BingWallpaper", Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "BingWallpaper.exe"));
            }
            else
            {
                if (rk.GetValue("BingWallpaper") != null)
                    rk.DeleteValue("BingWallpaper");
            }
        }

        public async void GetLatestWallpaper()
        {
            HistoryImage historyImage = null;

            try
            {
                historyImage = await _provider.GetLatestImage();
                HistoryImageProvider.AddImage(historyImage);
            }
            catch
            {
                historyImage = HistoryImageProvider.getRandom();
            }

            if (historyImage != null)
            {
                this.Invoke(new Action(async () =>
                {
                    this.CurrentWallpaper = historyImage;
                    await UpdateWallpaper();
                }));
            }
        }

        private async System.Threading.Tasks.Task UpdateWallpaper()
        {
            if (CurrentWallpaper != null)
            {
                try
                {
                    var img = await CurrentWallpaper.getImage();
                    Wallpaper.Set(img, Wallpaper.Style.Stretched);
                }
                catch { }
            }
        }

        public async void SetRandomWallpaper()
        {
            try
            {
                CurrentWallpaper = HistoryImageProvider.getRandom();
                await UpdateWallpaper();
            }
            catch { }
        }

        #region Tray Icons

        private NotifyIcon _trayIcon;
        private ContextMenu _trayMenu;
        private MenuItem _copyrightLabel;

        public void AddTrayIcons()
        {
            _trayMenu = new ContextMenu();

            _copyrightLabel = new MenuItem(Resource.AppName);
            _trayMenu.MenuItems.Add(_copyrightLabel);
            _trayMenu.MenuItems.Add("-");

            _trayMenu.MenuItems.Add(Resource.WallPaperStory, (s, e) =>
            {
                new WallpaperStoryForm(CurrentWallpaper).ShowDialog();
            });

            _trayMenu.MenuItems.Add(Resource.ForceUpdate, (s, e) => GetLatestWallpaper());

            _trayMenu.MenuItems.Add(Resource.Random, (s, e) => SetRandomWallpaper());

            var save = new MenuItem(Resource.Save);
            save.Click += async (s, e) =>
            {
                if (CurrentWallpaper != null)
                {
                    var fileName = string.Join("_", _settings.ImageCopyright.Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries)).TrimEnd('.');
                    var dialog = new SaveFileDialog
                    {
                        DefaultExt = "jpg",
                        Title = Resource.SaveDialgName,
                        FileName = fileName,
                        Filter = "Jpeg Image|*.jpg",
                    };
                    if (dialog.ShowDialog() == DialogResult.OK && dialog.FileName != "")
                    {
                        var image = await CurrentWallpaper.getImage();
                        image.Save(dialog.FileName, ImageFormat.Jpeg);
                        System.Diagnostics.Process.Start(dialog.FileName);
                    }
                }
            };
            _trayMenu.MenuItems.Add(save);

            _trayMenu.MenuItems.Add("-");

            var launch = new MenuItem(Resource.LaunchOnStartup);
            launch.Checked = _settings.LaunchOnStartup;
            launch.Click += OnStartupLaunch;
            _trayMenu.MenuItems.Add(launch);

            _trayMenu.MenuItems.Add(Resource.UpdateDB, (s, e) => UpdateLocalData());

            var timerChange = new MenuItem(Resource.IntervalChange);
            timerChange.Checked = _settings.AutoChange;

            var timeRanges = new string[] {
                "10" + Resource.Minutes,
                "30" + Resource.Minutes,
                "1" + Resource.Hours,
                "2" + Resource.Hours,
                "3" + Resource.Hours,
                "4" + Resource.Hours,
                "5" + Resource.Hours,
                "6" + Resource.Hours,
                "12" + Resource.Hours
            };

            foreach (var timeRange in timeRanges)
            {
                var rangeMenu = new MenuItem(timeRange);
                rangeMenu.Checked = _settings.AutoChangeInterval == timeRange;
                rangeMenu.Click += RangeMenu_Click; ;
                timerChange.MenuItems.Add(rangeMenu);
            }
            _trayMenu.MenuItems.Add(timerChange);

            _trayMenu.MenuItems.Add("-");

            _trayMenu.MenuItems.Add(Resource.Exit, (s, e) => Application.Exit());

            _trayIcon = new NotifyIcon();
            _trayIcon.Text = Resource.AppName;
            _trayIcon.Icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location);

            _trayIcon.MouseUp += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    MethodInfo mi = typeof(NotifyIcon).GetMethod("ShowContextMenu", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                    mi.Invoke(_trayIcon, null);
                }
            };

            _trayIcon.ContextMenu = _trayMenu;
            _trayIcon.Visible = true;
        }

        private void UpdateLocalData()
        {
            var images = IoliuBingCrawler.LoadHistoryImages();
            if (images.Count > 0)
            {
                HistoryImageProvider.AddBatch(images);
                ShowNotification(Resource.UpdateCount.Replace("{Count}", images.Count.ToString()));
            }
            else
            {
                ShowNotification(Resource.UpdateOver);
            }
        }

        private void UpdateLatestDaysImage()
        {
            var images = IoliuBingCrawler.LoadLatestDaysImages();
            if (images.Count > 0)
            {
                HistoryImageProvider.AddBatch(images);
            }
        }

        private void RangeMenu_Click(object sender, EventArgs e)
        {
            var intervalMenu = (MenuItem)sender;
            foreach (MenuItem subMenu in intervalMenu.Parent.MenuItems)
            {
                subMenu.Checked = false;
            }
            intervalMenu.Checked = !intervalMenu.Checked;

            if (autoChangeTimer != null)
            {
                autoChangeTimer.Stop();
                autoChangeTimer.Dispose();
                autoChangeTimer = null;
            }

            _settings.AutoChange = intervalMenu.Checked;
            _settings.AutoChangeInterval = intervalMenu.Text;

            if (intervalMenu.Checked)
            {
                CreateAutoChangeTask();
            }
        }

        private void OnStartupLaunch(object sender, EventArgs e)
        {
            var launch = (MenuItem)sender;
            launch.Checked = !launch.Checked;
            SetStartup(launch.Checked);
            _settings.LaunchOnStartup = launch.Checked;
        }

        protected override void Dispose(bool isDisposing)
        {
            if (isDisposing)
            {
                HotKey.UnregisterHotKey(this.Handle, 110);
                HotKey.UnregisterHotKey(this.Handle, 111);
                HotKey.UnregisterHotKey(this.Handle, 112);
                _trayIcon.Dispose();
            }

            base.Dispose(isDisposing);
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_HOTKEY = 0x312;
            switch (m.Msg)
            {
                case WM_HOTKEY:
                    switch (m.WParam.ToInt32())
                    {
                        case 110:
                            RandomWallpaper();
                            break;
                        case 111:
                            PreWallpaper();
                            break;
                        case 112:
                            NextWallpaper();
                            break;
                    }
                    break;
            }
            base.WndProc(ref m);
        }

        #endregion

        #region Notifications

        private void ShowNotification(string msg)
        {
            _trayIcon.BalloonTipTitle = Resource.AppName;
            _trayIcon.BalloonTipText = msg;
            _trayIcon.BalloonTipIcon = ToolTipIcon.Info;
            _trayIcon.Visible = true;
            _trayIcon.ShowBalloonTip(5000);
        }

        #endregion
    }
}
