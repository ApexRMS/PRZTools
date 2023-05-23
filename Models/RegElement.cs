﻿using System;
using ArcGIS.Desktop.Mapping;

namespace NCC.PRZTools
{
    public class RegElement
    {
        public RegElement()
        {

        }

        private int _elementID;

        public int ElementID
        {
            get => _elementID;
            set
            {
                if (value > 999999 || value < 1)
                {
                    throw new Exception("Element ID out of range (1 to 999999)");
                }

                _elementID = value;
                ElementTable = PRZConstants.c_TABLE_REGPRJ_PREFIX_ELEMENT + value.ToString(PRZHelper.CurrentGeodatabaseElementTableNameFormat);
            }
        }

        public string ElementName { get; set; }

        public int ElementType { get; set; }

        public int ElementStatus { get; set; }

        public string ElementUnit { get; set; }

        public string ElementDataPath { get; set; }

        public int ThemeID { get; set; }

        public string ThemeName { get; set; }

        public string ThemeCode { get; set; }

        public string ElementTable { get; private set; }

        public int ElementPresence { get; set; }

        // TODO: remove?
        public int ElementGoal { get; set; } = 50;

        public string WhereClause { get; set; }

        public string LyrxPath { get; set; }

        public string LayerName { get; set; }

        public int LayerType { get; set; }

        public Layer LayerObject { get; set; }

        public string LayerJson { get; set; }

        public string LegendGroup { get; set; }

        public string LegendClass { get; set; }



    }
}

