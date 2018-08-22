using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace Framework.Utils
{
    public class ScreenCapture
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowRect(IntPtr hWnd, ref Rect rect);

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetDesktopWindow();

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static Image CaptureDesktop()
        {
            return CaptureWindow(GetDesktopWindow());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filenamePath"></param>
        /// <param name="imageFormat"></param>
        public static void CaptureDesktopToFile(string filenamePath, ImageFormat imageFormat)
        {
            var dt = CaptureDesktop();
            dt.Save(filenamePath, imageFormat);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static Bitmap CaptureActiveWindow()
        {
            return CaptureWindow(GetForegroundWindow());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static void CaptureActiveWindowToFile(string filenamePath, ImageFormat imageFormat)
        {
            //Image i =  CaptureWindow(GetForegroundWindow());
            Image i = CaptureWindow(GetDesktopWindow());
            i.Save( filenamePath,  imageFormat);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="handle"></param>
        /// <returns></returns>
        public static Bitmap CaptureWindow(IntPtr handle)
        {
            var rect = new Rect();
            GetWindowRect(handle, ref rect);
            var bounds = new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
            var result = new Bitmap(bounds.Width, bounds.Height);

            using (var graphics = Graphics.FromImage(result))
            {
                graphics.CopyFromScreen(new Point(bounds.Left, bounds.Top), Point.Empty, bounds.Size);
            }

            return result;
        }
    }
}
