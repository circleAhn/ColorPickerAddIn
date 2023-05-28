using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.ML.Data;
using Microsoft.ML.Trainers;
using Microsoft.ML;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;


namespace ColorPickerAddIn
{
    internal class Drawer
    {
        Color[] palette;
        PowerPoint.Slide slide;

        public Drawer(Color[] palette, PowerPoint.Slide slide) {
            Init(palette, slide);
        }

        void Init(Color[] palette, PowerPoint.Slide slide)
        {
            this.palette = palette;
            this.slide = slide;
        }

        public void Draw()
        {
            var explainShape = slide.Shapes
                    .AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        50, 20, 200, 500);
            explainShape.Fill.BackColor.RGB = Color.White.ToArgb();

            for (int i = 0; i < palette.Length; ++i)
            {
                var shape = slide.Shapes
                .AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    50, 50 + 30 * i, 200, 200);

                shape.TextFrame.TextRange.Text = palette[i].ToHex();
                if (palette[i].GetGrayScale() < 200)
                {
                    shape.TextFrame.TextRange.Font.Color.RGB = Color.White.ToArgb();
                }
                else
                {
                    shape.TextFrame.TextRange.Font.Color.RGB = Color.Black.ToArgb();
                }
                shape.Fill.BackColor.RGB = palette[i].ToInvArgb();

            }
        }
    }
}
