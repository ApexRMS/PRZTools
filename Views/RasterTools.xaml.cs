﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for RasterTools.xaml
    /// </summary>
    public partial class RasterTools : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        internal RasterToolsVM vm = new RasterToolsVM();

        public RasterTools()
        {
            InitializeComponent();
            this.DataContext = vm;
        }

    }
}
