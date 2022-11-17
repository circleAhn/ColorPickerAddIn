using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.ML;
using Microsoft.ML.Data;
using Microsoft.ML.Trainers;
using System.Drawing;



namespace ColorPickerAddIn
{
    
    public class RGBPointf
    {
        [ColumnName("Features")]
        [VectorType(3)]
        public float[] Features { get; set; }

    }


    public static class HexColorExtensions
    {
        // PowerPoint's color codes seem to be reversed (i.e., BGR) not RGB
        public static string ToHex(this Color c) => $"#{c.B:X2}{c.G:X2}{c.R:X2}";
        public static int ToInvArgb(this Color c) => (c.A << 24) | (c.R << 16) | (c.G << 8) | c.B;
    }


    public class KMeansWithOptions : Setting
    {

        public Color[] Do(out bool isInvalid)
        {
            MLContext mlcontext = new MLContext();
            var dataPoints = GenerateRGBPoints();
            IDataView dataView = mlcontext.Data.LoadFromEnumerable(dataPoints);

            var options = new KMeansTrainer.Options
            {
                NumberOfClusters = NumberOfClusters,
                OptimizationTolerance = 1e-6f
            };

            var model = mlcontext.Clustering.Trainers.KMeans(options).Fit(dataView);
            var modelParams = model.Model;

            VBuffer<float>[] centroids = default;
            modelParams.GetClusterCentroids(ref centroids, out int k);

            
            Color[] colors = GetColors(centroids, out bool invalid);
            isInvalid = invalid;
            
            return colors
                .OrderBy(color => color.GetHue())
                .ThenBy(color => color.R * 3 + color.G * 2 + color.B * 1)
                .Reverse()
                .ToArray();

        }


        IEnumerable<RGBPointf> GenerateRGBPoints()
        {
            Random random = new Random(DateTime.Now.Minute * 60 + DateTime.Now.Second);

            for (int r = RedLowerBound; r <= RedUpperBound; r += 5)
                for (int g = GreenLowerBound; g <= GreenUpperBound; g += 5)
                    for (int b = BlueLowerBound; b <= BlueUpperBound; b += 5)
                    {
                        if (random.Next(10) == 0)
                        {
                            yield return new RGBPointf()
                            {
                                Features = new float[3] { r, g, b }
                            };
                        }
                    }
        }


        Color[] GetColors(VBuffer<float>[] centroids, out bool invalid)
        {
            invalid = false;
            int[][] colors = centroids
                .Select(it => it.GetValues().ToArray().Select(x => (int)x).ToArray())
                .ToArray();

            foreach (var color in colors)
                foreach (var c in color)
                    if (c < 0 || 255 < c) invalid = true;

            return colors
                .Select(color => Color.FromArgb(color[0] & 0xff, color[1] & 0xff, color[2] & 0xff))
                .ToArray();
        }
        

    }

}
