using Microsoft.ML;
using Microsoft.ML.Data;
using Microsoft.ML.Trainers;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;

namespace ColorPickerAddIn
{
    internal abstract class IKMeansModeler
    {
        public int NumberOfClusters;
        protected MLContext mlcontext;
        protected ClusteringPredictionTransformer<KMeansModelParameters> model;
        public bool invalid = false;


        public IKMeansModeler()
        {
            mlcontext = new MLContext();
        }

        public abstract ClusteringPredictionTransformer<KMeansModelParameters> Fit(IDataView rgbDataView);

        public ClusteringPredictionTransformer<KMeansModelParameters> Fit(IEnumerable<RGBPointf> rgbDataset)
        {
            var rgbDataView = mlcontext.Data.LoadFromEnumerable(rgbDataset);
            return Fit(rgbDataView);
        }

        public Color[] Predict()
        {
            var modelParams = model.Model;
            VBuffer<float>[] centroids = default;
            modelParams.GetClusterCentroids(ref centroids, out int _1);
            return BufferToColor(centroids);
        }

        Color[] BufferToColor(VBuffer<float>[] centroids)
        {
            //invalid = false;
            int[][] colors = centroids
                .Select(it => it.GetValues().ToArray().Select(x => (int)x).ToArray())
                .ToArray();

            foreach (var color in colors)
                foreach (var c in color)
                    if (c < 0 || 255 < c)
                    {
                        invalid = true;
                        //MessageBox.Show(c.ToString());
                    }

            return colors
                //.Select(color => Color.FromArgb(Math.Min(Math.Max(color[0], 0), 255), Math.Min(Math.Max(color[1], 0), 255), Math.Min(Math.Max(color[2], 0), 255)))
                .Select(color => Color.FromArgb(color[0] & 0xff, color[1] & 0xff, color[2] & 0xff))
                .OrderBy(color => color.GetHue())
                .ThenBy(color => color.R * 3 + color.G * 2 + color.B * 1)
                .Reverse()
                .ToArray();
        }
    }

    internal class KMeansColorPicker : IKMeansModeler
    {
        //MLContext mlcontext;
        public override ClusteringPredictionTransformer<KMeansModelParameters> Fit(IDataView rgbDataView)
        {
            var options = new KMeansTrainer.Options
            {
                NumberOfClusters = NumberOfClusters,
                OptimizationTolerance = 1e-4f,
                InitializationAlgorithm = KMeansTrainer.InitializationAlgorithm.KMeansPlusPlus
            };
            try
            {
                model = mlcontext.Clustering.Trainers.KMeans(options).Fit(rgbDataView);
            }
            catch
            {
                invalid = true;
            }
            return model;
        }

    }

}
