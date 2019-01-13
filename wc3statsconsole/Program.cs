using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronOcr;

namespace wc3statsconsole
{
    class Program
    {
        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
        //Mouse actions
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;
        static void Main(string[] args)
        {
            while (true)
            {
                var process = Process.Start(args[0]);
                Task.Delay(15000).Wait();
                if (process != null)
                {
                    IntPtr handle = process.MainWindowHandle;
                    SetForegroundWindow(handle);
                    SendKeys.SendWait("B");
                    Task.Delay(5000).Wait();
                    SendKeys.SendWait(args[1]);

                    //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 800, 800, 0, 0);
                    Task.Delay(1000).Wait();
                    SendKeys.SendWait("{ENTER}");
                    Task.Delay(10000).Wait();

                    var bmpScreenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width,
                        Screen.PrimaryScreen.Bounds.Height,
                        PixelFormat.Format32bppArgb);

                    // Create a graphics object from the bitmap.
                    var gfxScreenshot = Graphics.FromImage(bmpScreenshot);

                    // Take the screenshot from the upper left corner to the right bottom corner.
                    gfxScreenshot.CopyFromScreen(Screen.PrimaryScreen.Bounds.X,
                        Screen.PrimaryScreen.Bounds.Y,
                        0,
                        0,
                        Screen.PrimaryScreen.Bounds.Size,
                        CopyPixelOperation.SourceCopy);

                    // Save the screenshot to the specified path that the user has chosen.
                    bmpScreenshot.Save("Screenshot.png", ImageFormat.Png);
                    process.Kill();
                }

                var Ocr = new AutoOcr();
                var X = 1177; //px
                var Y = 319;
                var Width = 1920;
                var Height = 1080;
                var CropArea = new Rectangle(X, Y, Width, Height);
                var Result = Ocr.Read("Screenshot.png", CropArea);
                var resultText = Result.Text;

                var userStatistics = new UserStatistics(resultText, DateTimeOffset.Now);

                var streamWriter = File.AppendText("wc3Stats.csv");
                streamWriter.WriteLine(userStatistics.ToString());
                streamWriter.Flush();
                streamWriter.Close();

                Task.Delay(3600000).Wait();
            }
        }
    }

    internal class UserStatistics
    {
        public DateTimeOffset TimeStamp { get; }
        public string Users { get; }
        public string TftGames { get; }
        public string BnetUsers { get; }
        public string BnetGames { get; }

        public UserStatistics(string resultText, DateTimeOffset timeStamp)
        {
            ResultText = resultText;
            TimeStamp = timeStamp;

            var replace = resultText.Replace("\r\n", " ");
            var split = replace.Split(' ');

            var bnetUsers = split[15];
            if (!long.TryParse(bnetUsers.Replace("]", "1"), out var _)) bnetUsers = split[16];
            var bnetGames = split[18];
            if (!long.TryParse(bnetGames.Replace("]", "1"), out var _)) bnetGames = split[19];

            Users = split[3].Replace("]", "1");
            TftGames = split[6].Replace("]", "1");
            BnetUsers = bnetUsers.Replace("]", "1");
            BnetGames = bnetGames.Replace("]", "1");
        }

        public string ResultText { get; set; }

        public override string ToString()
        {
            return $"{TimeStamp}, {Users}, {TftGames}, {BnetUsers}, {BnetGames}, {ResultText}";
        }
    }
}
