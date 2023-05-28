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
using static System.Net.Mime.MediaTypeNames;

namespace ColorPickerAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private RandomRGBDataLoader Reload(string selectedItem)
        {
            switch (selectedItem)
            {
                case "밝게":
                    return new RandomRGBDataLoader(0, 50, 63, 255, 63, 255, 63, 255);
                case "어둡게":
                    return new RandomRGBDataLoader(0, 50, 0, 190, 0, 190, 0, 190);
                case "웜톤":
                    return new RandomRGBDataLoader(0, 50, 0, 190, 0, 190, 63, 255);
                case "쿨톤":
                    return new RandomRGBDataLoader(0, 50, 63, 255, 0, 190, 0, 190);
                case "완죤 랜덤":
                    return new RandomRGBDataLoader();
                default:
                    return new RandomRGBDataLoader(0, 50, 0, 255, 0, 255, 0, 255);
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            var numberOfClusters = numberOfClustersBox.Text;
            if (!numberOfClusters.All(char.IsNumber))
                return;


            KMeansColorPicker randomColor = default;
            Color[] palette = default;
            for (int i = 0; i < 5; ++i)
            {
                // Preprocessing
                var rgbDataLoader = Reload(dropDown1.SelectedItem.Label);
                var RGBDataset = rgbDataLoader.RGBDataset;


                // Model Training & Predict
                randomColor = new KMeansColorPicker()
                {
                    NumberOfClusters = int.Parse(numberOfClusters)
                };
                randomColor.Fit(RGBDataset);
                palette = randomColor.Predict();

                if (!randomColor.invalid) break;

            }


            // Drawing
            if (randomColor.invalid)
            {
                MessageBox.Show("색선정에 오류가 일부 있습니다. 다시 돌려주세요.");
                return;
            }

            var drawer = new Drawer(palette, slide);
            drawer.Draw();

        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }




        private void GetBitmap_Click(object sender, RibbonControlEventArgs e)
        {
            var numberOfClusters = numberOfClustersBox.Text;
            if (!numberOfClusters.All(char.IsNumber))
                return;
            

            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                return;

            // Image Preprocessing
            selection.Copy();
            if (!Clipboard.ContainsImage()) 
                return;
            var image = Clipboard.GetImage();
            var rgbDataLoader = new ExtractedRGBDataLoader(image, 50);
            var RGBDataset = rgbDataLoader.RGBDataset;


            // Model Training & Predict
            KMeansColorPicker colorPicker = default;
            Color[] palette = default;
            for (int i = 0; i < 5; ++i)
            {
                colorPicker = new KMeansColorPicker()
                {
                    NumberOfClusters = int.Parse(numberOfClusters)
                };
                colorPicker.Fit(RGBDataset);
                palette = colorPicker.Predict();

                if (!colorPicker.invalid) break;
                
            }


            // Drawing
            if (colorPicker.invalid)
            {
                MessageBox.Show("색선정에 오류가 일부 있습니다. 다시 돌려주세요.");
                return;
            }

            var drawer = new Drawer(palette, slide);
            drawer.Draw();

        }
    }
}
