using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.ML.Data;

namespace ColorPickerAddIn
{

    public class RGBPointf
    {
        [ColumnName("Features")]
        [VectorType(3)]
        public float[] Features { get; set; }

    }

    internal abstract class IDataLoader : Setting
    {
        public IEnumerable<RGBPointf> RGBDataset;
        protected abstract IEnumerable<RGBPointf> GenerateRGBPoints();

    }

    internal class RandomRGBDataLoader : IDataLoader
    {
        public void Init(int randm, int randM, int rLowerPad, int rUpperPad, int gLowerPad, int gUpperPad, int bLowerPad, int bUpperPad)
        {
            Random randomGenerator = new Random(DateTime.Now.Second);
            int Rand() => randomGenerator.Next(randm, randM);
            SetParameters(rLowerPad + Rand(), rUpperPad - Rand(),
                gLowerPad + Rand(), gUpperPad - Rand(),
                bLowerPad + Rand(), bUpperPad - Rand());
        }

        public void Init()
        {
            Random randomGenerator = new Random(DateTime.Now.Second);
            int Rand(int m, int M) => randomGenerator.Next(m, M);
            SetParameters(Rand(20, 100), Rand(150, 235), Rand(20, 100), Rand(150, 235), Rand(20, 100), Rand(150, 235));
        }


        public RandomRGBDataLoader()
        {
            Init();
            RGBDataset = GenerateRGBPoints();
        }

        public RandomRGBDataLoader(int randm, int randM, int rLowerPad, int rUpperPad, int gLowerPad, int gUpperPad, int bLowerPad, int bUpperPad)
        {
            Init(randm, randM, rLowerPad, rUpperPad, bLowerPad, bUpperPad, gLowerPad, gUpperPad);
            RGBDataset = GenerateRGBPoints();
        }

        protected override IEnumerable<RGBPointf> GenerateRGBPoints()
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
    }

    internal class ExtractedRGBDataLoader : IDataLoader
    {
        Bitmap bitmap;
        public int width;
        public int height;
        byte[] bytes;

        void BitmapInit(Bitmap bmp, int scaledWidth, int scaledHeight)
        {
            bitmap = ImageScaler(bmp, scaledWidth, scaledHeight);
            width = bitmap.Width;
            height = bitmap.Height;
            bytes = FastRGBCopy();
            RGBDataset = GenerateRGBPoints();
        }

        public ExtractedRGBDataLoader(Image img, int scaledWidth = -1, int scaledHeight = -1)
        {
            BitmapInit(new Bitmap(img), scaledWidth, scaledHeight);
        }

        public ExtractedRGBDataLoader(Bitmap bmp, int scaledWidth = -1, int scaledHeight = -1)
        {
            BitmapInit(bmp, scaledWidth, scaledHeight);
        }

        protected override IEnumerable<RGBPointf> GenerateRGBPoints()
        {
            Random randNum = new Random();
            int[] noises = Enumerable
                            .Repeat(0, bytes.Length)
                            .Select(i => randNum.Next(-1, 1))
                            .ToArray();

            for (int i = 0; i < noises.Length; i++)
            {
                noises[i] += bytes[i];
                noises[i] = Math.Min(Math.Max(noises[i], 0), 255);
            }

            for (int i = 0; i < bytes.Length; i += 4)
            {
                yield return new RGBPointf()
                {
                    Features = new float[3]
                    {
                        noises[i], noises[i + 1], noises[i + 2]
                    }
                };
            }
        }

        Bitmap ImageScaler(Bitmap bmp, int w, int h)
        {
            if (w > 0)
            {
                if (h < 0)
                {
                    int newHeight = w * bmp.Height / bmp.Width;
                    return new Bitmap(bmp, w, newHeight);
                }
                return new Bitmap(bmp, w, h);
            }
            return bitmap;
        }

        byte[] FastRGBCopy()
        {
            Rectangle rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
            BitmapData bitmapdata = bitmap.LockBits(rect, ImageLockMode.ReadWrite, bitmap.PixelFormat);
            int length = bitmapdata.Stride * bitmapdata.Height;
            byte[] bytes = new byte[length];

            // copy bitmap to byte[]
            Marshal.Copy(bitmapdata.Scan0, bytes, 0, length);
            bitmap.UnlockBits(bitmapdata);

            return bytes;//.Select(i => (sbyte)i).ToArray();
        }
    }

}
