using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ColorPickerAddIn
{

    public static class HexColorExtensions
    {
        // PowerPoint's color codes seem to be reversed (i.e., BGR) not RGB
        public static string ToHex(this Color c) => $"#{c.B:X2}{c.G:X2}{c.R:X2}";
        public static int ToInvArgb(this Color c) => (c.A << 24) | (c.R << 16) | (c.G << 8) | c.B;
        public static int GetGrayScale(this Color c) => (c.R + c.G + c.B) / 3;
    }
}
