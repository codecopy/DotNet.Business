using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNet.Business.CopyFromScreen
{
    public class CopyFromScreen
    {
        public static Bitmap GetImgDesk()
        {
            Rectangle rect = System.Windows.Forms.SystemInformation.VirtualScreen;
            //获取屏幕分辨率
            int x_ = rect.Width;
            int y_ = rect.Height;
            //截屏
            Bitmap img = new Bitmap(x_, y_);
            Graphics g = Graphics.FromImage(img);
            g.CopyFromScreen(new Point(0, 0), new Point(0, 0), new Size(x_, y_));
            return img;
        }
    }
}
