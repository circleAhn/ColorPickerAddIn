﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ColorPickerAddIn
{
    public class Constant
    {
        public const int rgbLowerBound = 0;
        public const int rgbUpperBound = 255;
        public const int numberOfClusters = 7;


    }

    public class Setting
    {
        public int RedLowerBound { get; set; }
        public int RedUpperBound { get; set; }
        public int GreenLowerBound { get; set; }
        public int GreenUpperBound { get; set; }
        public int BlueLowerBound { get; set; }
        public int BlueUpperBound { get; set; }


        public Setting(int redLowerBound = 0, int redUpperBound = 255,
            int greenLowerBound = 0, int greenUpperBound = 255,
            int blueLowerBound = 0, int blueUpperBound = 255)
        {
            SetParameters(redLowerBound, redUpperBound,
                greenLowerBound, greenUpperBound,
                blueLowerBound, blueUpperBound);
        }


        public void SetParameters(int redLowerBound = 0, int redUpperBound = 255,
            int greenLowerBound = 0, int greenUpperBound = 255,
            int blueLowerBound = 0, int blueUpperBound = 255)
        {
            RedLowerBound = redLowerBound;
            RedUpperBound = redUpperBound;
            GreenLowerBound = greenLowerBound;
            GreenUpperBound = greenUpperBound;
            BlueLowerBound = blueLowerBound;
            BlueUpperBound = blueUpperBound;

        }




    }
}
