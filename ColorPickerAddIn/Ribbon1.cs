using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace ColorPickerAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            var numberOfClusters = editBox1.Text;
            if (!numberOfClusters.All(char.IsNumber))
                return;


            var KMeans = new KMeansWithOptions()
            {
                NumberOfClusters = int.Parse(numberOfClusters),
                Option = dropDown1.SelectedItem.Label
            };




            System.Drawing.Color[] Colors = default;
            bool isInvalid = false;

            Random randomGenerator = new Random(DateTime.Now.Second);
            int Rand(int m = 0, int M = 255) => randomGenerator.Next(m, M);
            for (int i = 0; i < 10; ++i)
            {
                var data = dropDown1.SelectedItem.Label;
                if (data == "기본")
                    KMeans.SetParameters(int.Parse(numberOfClusters), 0 + Rand(0, 50), 255 - Rand(0, 50), 0 + Rand(0, 50), 255 - Rand(0, 50), 0 + Rand(0, 50), 255 - Rand(0, 50));
                else if (data == "밝게")
                    KMeans.SetParameters(int.Parse(numberOfClusters), 63 + Rand(0, 50), 255 - Rand(0, 50), 63 + Rand(0, 50), 255 - Rand(0, 50), 63 + Rand(0, 50), 255 - Rand(0, 50));
                else if (data == "어둡게")
                    KMeans.SetParameters(int.Parse(numberOfClusters), 0 + Rand(0, 50), 190 - Rand(0, 50), 0 + Rand(0, 50), 190 - Rand(0, 50), 0 + Rand(0, 50), 190 - Rand(0, 50));
                else if (data == "웜톤")
                    KMeans.SetParameters(int.Parse(numberOfClusters), 0 + Rand(0, 50), 190 - Rand(0, 50), 0 + Rand(0, 50), 190 - Rand(0, 50), 63 + Rand(0, 50), 255 - Rand(0, 50));
                else if (data == "쿨톤")
                    KMeans.SetParameters(int.Parse(numberOfClusters), 63 + Rand(0, 50), 255 - Rand(0, 50), 0 + Rand(0, 50), 190 - Rand(0, 50), 0 + Rand(0, 50), 190 - Rand(0, 50));
                else if (data == "완죤 랜덤")
                    KMeans.SetParameters(int.Parse(numberOfClusters), Rand(20, 100), Rand(150, 235), Rand(20, 100), Rand(150, 235), Rand(20, 100), Rand(150, 235));
                else
                    KMeans.SetParameters(int.Parse(numberOfClusters), 0 + Rand(0, 50), 255 - Rand(0, 50), 0 + Rand(0, 50), 255 - Rand(0, 50), 0 + Rand(0, 50), 255 - Rand(0, 50));

                Colors = KMeans.Do(out isInvalid);

                if (!isInvalid) break;
            }


            var explainShape = slide.Shapes
                    .AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        50, 20, 500, 500);
            explainShape.Fill.BackColor.RGB = System.Drawing.Color.White.ToArgb();

            if (isInvalid)
                explainShape.TextFrame.TextRange.Text = "색선정에 오류가 일부 있습니다. 다시 돌려주세요.";

            for (int i = 0; i < Colors.Length; ++i)
            {
                var shape = slide.Shapes
                .AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    50, 50 + 30 * i, 200, 200);

                shape.TextFrame.TextRange.Text = Colors[i].ToHex();
                shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.White.ToArgb();
                shape.Fill.BackColor.RGB = Colors[i].ToInvArgb();


            }
            
        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
