using Microsoft.VisualBasic.ApplicationServices;
using Microsoft.Win32;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace FreeCrosshair
{
    internal class ApplicationBase : WindowsFormsApplicationBase, IDisposable
    {
        #region Dispose Support
        private bool disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposedValue)
            {
                if (disposing)
                {
                    // TODO: 释放托管状态(托管对象)
                    if (this.styleMenuOwner != null)
                    {
                        this.styleMenuOwner.Dispose();
                    }

                    if (this.colorMenuOwner != null)
                    {
                        this.colorMenuOwner.Dispose();
                    }

                    if (this.customColorMenuItem != null)
                    {
                        this.customColorMenuItem.Dispose();
                    }

                    if (this.crosshairImage != null)
                    {
                        this.crosshairImage.Dispose();
                    }

                    if (this.menu != null)
                    {
                        this.menu.Dispose();
                    }

                    if (this.notifyIcon != null)
                    {
                        this.notifyIcon.Dispose();
                    }
                }

                // TODO: 释放未托管的资源(未托管的对象)并重写终结器
                // TODO: 将大型字段设置为 null
                this.disposedValue = true;
            }
        }

        // // TODO: 仅当“Dispose(bool disposing)”拥有用于释放未托管资源的代码时才替代终结器
        // ~ApplicationBase()
        // {
        //     // 不要更改此代码。请将清理代码放入“Dispose(bool disposing)”方法中
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // 不要更改此代码。请将清理代码放入“Dispose(bool disposing)”方法中
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
        #endregion

        #region Windows API
        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern void ReleaseDC(IntPtr hwnd, IntPtr hdc);

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        [DllImport("Shell32.dll")]
        private static extern void SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
        #endregion

        private bool doDrawing = true; //是否绘制标识

        private Color crosshairColor = Color.Empty;
        private Bitmap crosshairImage = null;

        private ContextMenuStrip menu = null;
        private NotifyIcon notifyIcon = null;

        private ToolStripMenuItem styleMenuOwner = null;
        private ToolStripMenuItem colorMenuOwner = null;
        private ToolStripMenuItem customColorMenuItem = null;

        public ApplicationBase() : base()
        {
            base.IsSingleInstance = true;
            base.SaveMySettingsOnExit = true;
            base.ShutdownStyle = ShutdownMode.AfterAllFormsClose;
        }

        protected override bool OnStartup(StartupEventArgs eventArgs)
        {
            base.OnStartup(eventArgs);

            if (this.SetCompatLayer())
            {
                this.menu = new ContextMenuStrip();

                this.notifyIcon = new NotifyIcon()
                {
                    ContextMenuStrip = menu,
                    Icon = Properties.Resources.App,
                    Text = Application.ProductName,
                    Visible = true
                };

                this.notifyIcon.DoubleClick += this.ChangeStyle;

                this.InitContextMenu(menu);
                this.ParseCommandArgs(eventArgs.CommandLine.ToArray());
                this.DrawCrosshairAsync();

                Application.Run();
            }

            return false;
        }

        protected override void OnStartupNextInstance(StartupNextInstanceEventArgs eventArgs)
        {
            base.OnStartupNextInstance(eventArgs);

            if (eventArgs.CommandLine.Count == 0)
            {
                this.ChangeStyle(this.notifyIcon, null);
            }
            else
            {
                this.ParseCommandArgs(eventArgs.CommandLine.ToArray());
            }
        }

        private ToolStripMenuItem CurrentStyleMenuItem
        {
            get => this.styleMenuOwner == null ? null : this.styleMenuOwner.DropDownItems.OfType<ToolStripMenuItem>().FirstOrDefault(x => x.Checked);
        }

        private ToolStripMenuItem CurrentColorMenuItem
        {
            get => this.colorMenuOwner == null ? null : this.colorMenuOwner.DropDownItems.OfType<ToolStripMenuItem>().FirstOrDefault(x => x.Checked);
        }

        private Bitmap CurrentImage
        {
            get => this.CurrentStyleMenuItem == null ? null : (Bitmap)this.CurrentStyleMenuItem.Tag;
        }

        private Color CurrentColor
        {
            get => this.CurrentColorMenuItem == null ? Color.Empty : (Color)this.CurrentColorMenuItem.Tag;
        }

        private bool SetCompatLayer()
        {
            using (var regKey = Registry.CurrentUser.CreateSubKey("Software\\Microsoft\\Windows NT\\CurrentVersion\\AppCompatFlags\\Layers"))
            {
                var value = regKey.GetValue(Application.ExecutablePath, string.Empty) as string;

                if (value == "~ HIGHDPIAWARE")
                {
                    return true;
                }
                else
                {
                    regKey.SetValue(Application.ExecutablePath, "~ HIGHDPIAWARE");
                }
            }

            Application.Restart();
            return false;
        }

        private void ChangeStyle(object o, EventArgs e)
        {
            var menuItems = this.styleMenuOwner.DropDownItems.OfType<ToolStripMenuItem>().ToArray();
            var currentItem = menuItems.FirstOrDefault(x => x.Checked);
            var nextItem = (currentItem == null ? menuItems[0] : menuItems.NextOrDefault(currentItem));

            if (nextItem == null)
            {
                nextItem = menuItems[0];
            }

            this.SetCosshairStyle(nextItem);
        }

        private void SetSelectedStyle(object o, EventArgs e)
        {
            var selectedItem = (ToolStripMenuItem)o;
            this.SetCosshairStyle(selectedItem);
        }

        private void SetSelectedColor(object o, EventArgs e)
        {
            var selectedItem = (ToolStripMenuItem)o;
            this.SetCosshairColor(selectedItem);
        }

        private void SetRandomColor(object o, EventArgs e)
        {
            var random = new Random();
            var color = Color.FromArgb(255, random.Next(256), random.Next(256), random.Next(256));

            var selectedItem = (ToolStripMenuItem)o;
            var colorText = color.ToHexString();

            this.FillImage(color, selectedItem.Image);

            selectedItem.Text = "随机颜色 - " + colorText;
            selectedItem.ForeColor = color;
            selectedItem.Tag = color;

            this.SetCosshairColor(selectedItem);
        }

        private void SetCustomColor(object o, EventArgs e)
        {
            var selectedItem = (ToolStripMenuItem)o;
            selectedItem.Enabled = false;

            Color color;

            using (var dialog = new ColorDialog() { FullOpen = true })
            {
                if (dialog.ShowDialog() == DialogResult.Cancel)
                {
                    selectedItem.Enabled = true;
                    return;
                }

                color = dialog.Color;
            }

            selectedItem.Enabled = true;

            this.SetColorMenuItem(selectedItem, color, "自定义颜色");
            this.SetCosshairColor(selectedItem);
        }

        private void ShowAbout(object o, EventArgs e)
        {
            var selectedItem = (ToolStripMenuItem)o;
            selectedItem.Enabled = false;

            MessageBox.Show("名称：" + Application.ProductName + "\r\n\r\n" +
                            "版本：" + Application.ProductVersion + "\r\n\r\n" +
                            "用法：" + Path.GetFileName(Application.ExecutablePath) + " [-Style=FreeCrosshair2] [-Color=FF00FF00]" + "\r\n\r\n\r\n\r\n" +
                            "参数说明：" + "\r\n\r\n" +
                            "-Style=FreeCrosshair2    准心样式，内置或者自定义准心的样式名称。" + "\r\n\r\n" +
                            "-Color=FF00FF00    准心颜色，内置或者自定义准心的颜色代码。" + "\r\n\r\n\r\n\r\n" +
                            "自定义样式目录：" + Path.Combine(Application.StartupPath, "Styles") + "\r\n\r\n" +
                            "自定义样式文件：PNG图像文件" + "\r\n\r\n" +
                            "自定义颜色目录：" + Path.Combine(Application.StartupPath, "Colors") + "\r\n\r\n" +
                            "自定义颜色文件：CLR文本文件，文件内容为颜色代码",
                            Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);

            selectedItem.Enabled = true;
        }

        private void ExitApp(object o, EventArgs e)
        {
            this.doDrawing = false;
        }

        private void InitContextMenu(ContextMenuStrip owner)
        {
            this.styleMenuOwner = new ToolStripMenuItem("样式(&S)");
            owner.Items.Add(this.styleMenuOwner);

            var defaultStyle1 = new ToolStripMenuItem("FreeCrosshair1", null, this.SetSelectedStyle);
            var defaultStyle2 = new ToolStripMenuItem("FreeCrosshair2", null, this.SetSelectedStyle);

            defaultStyle1.Tag = Properties.Resources.FreeCosshair1;
            defaultStyle2.Tag = Properties.Resources.FreeCosshair2;

            this.styleMenuOwner.DropDownItems.Add(defaultStyle1);
            this.styleMenuOwner.DropDownItems.Add(defaultStyle2);

            var stylesDir = Path.Combine(Application.StartupPath, "Styles");

            if (Directory.Exists(stylesDir))
            {
                var imageFiles = Directory.GetFiles(stylesDir, "*.png");

                if (imageFiles.Length > 0)
                {
                    this.styleMenuOwner.DropDownItems.Add(new ToolStripSeparator());

                    foreach (var imageFile in imageFiles)
                    {
                        var imageFileName = Path.GetFileNameWithoutExtension(imageFile);
                        var menuItem = new ToolStripMenuItem(imageFileName, null, this.SetSelectedStyle);

                        menuItem.Tag = Image.FromFile(imageFile);

                        this.styleMenuOwner.DropDownItems.Add(menuItem);
                    }
                }
            }

            var colorNames = new string[] { "红色(&R)", "黄色(&Y)", "蓝色(&B)", "绿色(&G)", "青色(&C)", "靛色(&I)", "紫色(&P)", "黑色(&B)", "白色(&W)" };
            var colorValues = new string[] { "FFFF0000", "FFFFFF00", "FF0000FF", "FF00FF00", "FF00FFFF", "FF4B0082", "FF800080", "FF000000", "FFFFFFFF" };

            this.colorMenuOwner = new ToolStripMenuItem("颜色(&C)");
            owner.Items.Add(this.colorMenuOwner);

            // 原始颜色，保持PNG图像颜色不变
            var noneColorMenuItem = new ToolStripMenuItem("原始颜色", this.CreateImage(Color.Transparent), this.SetSelectedColor);
            noneColorMenuItem.Tag = Color.Empty;
            this.colorMenuOwner.DropDownItems.Add(noneColorMenuItem);

            this.colorMenuOwner.DropDownItems.Add(new ToolStripSeparator());

            for (int i = 0; i < colorNames.Length; i++)
            {
                var color = Color.FromArgb(Convert.ToInt32(colorValues[i], 16));
                var image = this.CreateImage(color);
                var menuItem = new ToolStripMenuItem(colorNames[i] + " - " + colorValues[i], image, this.SetSelectedColor);

                menuItem.ForeColor = color;
                menuItem.Tag = color;

                this.colorMenuOwner.DropDownItems.Add(menuItem);
            }

            var colorsDir = Path.Combine(Application.StartupPath, "Colors");

            if (Directory.Exists(colorsDir))
            {
                var colorFiles = Directory.GetFiles(colorsDir, "*.clr");

                if (colorFiles.Length > 0)
                {
                    this.colorMenuOwner.DropDownItems.Add(new ToolStripSeparator());

                    foreach (var imageFile in colorFiles)
                    {
                        try
                        {
                            var imageFileName = Path.GetFileNameWithoutExtension(imageFile);
                            var colorValue = File.ReadAllText(imageFile).Trim();
                            var color = Color.FromArgb(Convert.ToInt32(colorValue, 16));
                            var image = this.CreateImage(color);
                            var menuItem = new ToolStripMenuItem(imageFileName + " - " + colorValue, image, this.SetSelectedColor);

                            menuItem.ForeColor = color;
                            menuItem.Tag = color;

                            this.colorMenuOwner.DropDownItems.Add(menuItem);
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                    }
                }
            }

            this.colorMenuOwner.DropDownItems.Add(new ToolStripSeparator());

            // 自定义颜色，弹出颜色对话框
            this.customColorMenuItem = new ToolStripMenuItem("自定义颜色", this.CreateImage(Color.Transparent), this.SetCustomColor);
            this.customColorMenuItem.Tag = Color.Empty;
            this.colorMenuOwner.DropDownItems.Add(this.customColorMenuItem);

            // 随机颜色，不弹出颜色对话框
            var randomColorMenuItem = new ToolStripMenuItem("随机颜色", this.CreateImage(Color.Transparent), this.SetRandomColor);
            randomColorMenuItem.Tag = Color.Empty;
            this.colorMenuOwner.DropDownItems.Add(randomColorMenuItem);

            owner.Items.Add(new ToolStripSeparator());

            // 关于，弹出关于对话框
            var aboutMenuItem = new ToolStripMenuItem("关于(&A)", null, this.ShowAbout);
            owner.Items.Add(aboutMenuItem);

            // 退出
            var exitMenuItem = new ToolStripMenuItem("退出(&X)", null, this.ExitApp);
            owner.Items.Add(exitMenuItem);
        }

        private void SetColorMenuItem(ToolStripMenuItem item, Color color, string text)
        {
            this.FillImage(color, item.Image);

            item.Text = text + " - " + color.ToHexString();
            item.ForeColor = color;
            item.Tag = color;
        }

        private void ParseCommandArgs(string[] args)
        {
            var crosshairStyleText = string.Empty;
            var crosshairColorText = string.Empty;

            foreach (var arg in args)
            {
                var argItems = arg.Split('=');
                var argKey = argItems[0].ToLower();

                switch (argKey)
                {
                    case "/style":
                    case "-style":
                        crosshairStyleText = (argItems.Length == 1 ? string.Empty : argItems[1]);
                        break;
                    case "/color":
                    case "-color":
                        crosshairColorText = (argItems.Length == 1 ? string.Empty : argItems[1]);
                        break;
                    default:
                        break;
                }
            }

            // 默认样式序号：0-FreeCrosshair1
            // 默认颜色序号：3-绿色
            var styleMenuItem = this.CurrentStyleMenuItem == null ? (ToolStripMenuItem)this.styleMenuOwner.DropDownItems[0] : this.CurrentStyleMenuItem;
            var colorMenuItem = this.CurrentColorMenuItem == null ? (ToolStripMenuItem)this.colorMenuOwner.DropDownItems[0] : this.CurrentColorMenuItem;

            if (!string.IsNullOrEmpty(crosshairStyleText))
            {
                var item = this.styleMenuOwner.DropDownItems.OfType<ToolStripMenuItem>().FirstOrDefault(x => x.Text == crosshairStyleText);

                if (item != null)
                {
                    styleMenuItem = item;
                }
            }

            if (!string.IsNullOrEmpty(crosshairColorText))
            {
                var crosshairColor = Color.FromArgb(255, 0, 255, 0);

                if (int.TryParse(crosshairColorText, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var colorValue))
                {
                    crosshairColor = Color.FromArgb(colorValue);
                }

                var item = this.colorMenuOwner.DropDownItems.OfType<ToolStripMenuItem>().FirstOrDefault(x => ((Color)x.Tag) == crosshairColor);

                if (item == null)
                {
                    this.SetColorMenuItem(this.customColorMenuItem, crosshairColor, "自定义颜色");
                    colorMenuItem = this.customColorMenuItem;
                }
            }

            this.SetCosshairStyle(styleMenuItem);
            this.SetCosshairColor(colorMenuItem);
        }

        private void SetCosshairStyle(ToolStripMenuItem menuItem)
        {
            this.ResetMenuItemState(menuItem);

            var newCrossImage = (Bitmap)menuItem.Tag;

            if (this.crosshairImage != newCrossImage)
            {
                this.crosshairColor = this.CurrentColor;
                this.crosshairImage = newCrossImage;

                this.CreateNewCrosshair();
            }
        }

        private void SetCosshairColor(ToolStripMenuItem menuItem)
        {
            this.ResetMenuItemState(menuItem);

            var newCrossColor = (Color)menuItem.Tag;

            if (this.crosshairColor != newCrossColor)
            {
                this.crosshairColor = newCrossColor;
                this.crosshairImage = this.CurrentImage;

                this.CreateNewCrosshair();
            }
        }

        private void CreateNewCrosshair()
        {
            if (this.crosshairColor != Color.Empty)
            {
                var newImage = new Bitmap(this.crosshairImage.Width, this.crosshairImage.Height);

                for (int x = 0; x < this.crosshairImage.Width; x++)
                {
                    for (int y = 0; y < this.crosshairImage.Height; y++)
                    {
                        var color = this.crosshairImage.GetPixel(x, y);

                        if (color.A > 0)
                        {
                            newImage.SetPixel(x, y, Color.FromArgb(color.A, this.crosshairColor.R, this.crosshairColor.G, this.crosshairColor.B));
                        }
                    }
                }

                this.crosshairImage = newImage;
            }

            //this.RefreshDesktop();
        }

        private Image CreateImage(Color color)
        {
            var image = new Bitmap(16, 16);
            this.FillImage(color, image);

            return image;
        }

        private void FillImage(Color color, Image image)
        {
            using (var g = Graphics.FromImage(image))
            using (var brush = new SolidBrush(color))
            {
                g.FillRectangle(brush, 0, 0, image.Width, image.Height);
            }
        }

        private void ResetMenuItemState(ToolStripMenuItem current)
        {
            foreach (var item in current.Owner.Items.OfType<ToolStripMenuItem>().Where(x => x != current && x.Font.Bold))
            {
                item.Checked = false;
                item.Font = new Font(item.Font, FontStyle.Regular);
            }

            current.Checked = true;
            current.Font = new Font(current.Font, FontStyle.Bold);
        }

        private async void DrawCrosshairAsync()
        {
            this.doDrawing = true;

            var dpiScaling = this.GetDpiScaling();

            var w = GetSystemMetrics(0);
            var h = GetSystemMetrics(1);

            var screenWidth = Convert.ToInt32(w * dpiScaling);
            var screenHeight = Convert.ToInt32(h * dpiScaling);

            //中心点
            var imageSize = this.crosshairImage.Size;

            var x = screenWidth / 2 - imageSize.Width / 2;
            var y = screenHeight / 2 - imageSize.Height / 2;

            var hdc = GetDC(IntPtr.Zero);

            using (var g = Graphics.FromHdc(hdc))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;

                while (this.doDrawing)
                {
                    g.DrawImage(this.crosshairImage, x, y);
                    await Task.Delay(10);
                }
            }

            ReleaseDC(IntPtr.Zero, hdc);
            //this.RefreshDesktop();

            Application.Exit();
        }

        private void RefreshDesktop()
        {
            const int SHCNE_ASSOCCHANGED = 0x08000000;
            const int SHCNF_IDLIST = 0x0000;

            SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero);
        }

        private float GetDpiScaling()
        {
#if DEBUG
            try
            {
                var appliedDPI = (int)Registry.GetValue("HKEY_CURRENT_USER\\Control Panel\\Desktop\\WindowMetrics", "AppliedDPI", 96);
                var result = appliedDPI / 96f;

                return result;
            }
            catch (Exception)
            {
                return 1f;
            }
#else
            return 1f;
#endif
        }
    }
}
