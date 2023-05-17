﻿using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ProMsgBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;
using PRZH = NCC.PRZTools.PRZHelper;

namespace NCC.PRZTools
{
    public enum NationalGridDimension
    {
        SideLength_1m,
        SideLength_10m,
        SideLength_100m,
        SideLength_1000m,
        SideLength_10000m,
        SideLength_100000m
    }

    /// <summary>
    /// This class models the behaviour of the national prioritization grids.
    /// The various grids all share a common outer extent envelope:
    /// LL (-2500000, 200000), UR (3200000, 4900000)
    /// The grids use the custom Canada Albers spatial reference
    /// The grids all share a similar identifier structure
    /// aaaaaaa[EW]_bbbbbbb_d
    /// aaaaaaa = X-coordinate (unsigned and padded with leading 0s, always 7 digits)
    /// EW = E for positive or zero X-coordinate, and W for negative X-coordinate
    /// bbbbbbb = Y-coordinate (unsigned and padded with leading 0s, always 7 digits)
    /// d = national grid dimension: 0 = 1m, 1 = 10m, 2 = 100m, 3 = 1000m, 4 = 10000m, 5 = 100000m
    /// 
    /// National Grid Maximum Dimensions
    /// 
    /// 1 m:
    /// 10 m:
    /// 100 m:
    /// 1000 m: 5700 columns x 4700 rows    (26,790,000 cells)
    /// 
    /// </summary>
    public static class NationalGrid
    {
        #region CONSTANTS

        public const int c_NATGRID_ENV_XMIN = -2500000;    // Minimum X for National Grid's extent
        public const int c_NATGRID_ENV_YMIN = 200000;      // Minimum Y "
        public const int c_NATGRID_ENV_XMAX = 3200000;     // Maximum X "
        public const int c_NATGRID_ENV_YMAX = 4900000;     // Maximum Y "

        public const int c_NATGRID_ROWS_DIM0 = 0;
        public const int c_NATGRID_ROWS_DIM1 = 0;
        public const int c_NATGRID_ROWS_DIM2 = 0;
        public const int c_NATGRID_ROWS_DIM3 = 4700;
        public const int c_NATGRID_ROWS_DIM4 = 0;
        public const int c_NATGRID_ROWS_DIM5 = 0;

        public const int c_NATGRID_COLUMNS_DIM0 = 0;
        public const int c_NATGRID_COLUMNS_DIM1 = 0;
        public const int c_NATGRID_COLUMNS_DIM2 = 0;
        public const int c_NATGRID_COLUMNS_DIM3 = 5700;
        public const int c_NATGRID_COLUMNS_DIM4 = 0;
        public const int c_NATGRID_COLUMNS_DIM5 = 0;

        // Tiling
        public const int c_NATGRID_TILE_ROW_HEIGHT = 500;
        public const int c_NATGRID_TILE_COL_WIDTH = 500;

        public const int c_NATGRID_TILE_ROW_COUNT = 10; // TODO: Replace with function when other grids are supported
        public const int c_NATGRID_TILE_COLS_COUNT = 12;

        #endregion

        #region STATIC METHODS

        public static (bool success, long cell_number, string message) GetCellNumberFromULXY(int UL_X, int UL_Y, NationalGridDimension dimension)
        {
            try
            {
                int side_length = 0;

                switch (dimension)
                {
                    case NationalGridDimension.SideLength_1m:
                        side_length = 1;
                        break;
                    case NationalGridDimension.SideLength_10m:
                        side_length = 10;
                        break;
                    case NationalGridDimension.SideLength_100m:
                        side_length = 100;
                        break;
                    case NationalGridDimension.SideLength_1000m:
                        side_length = 1000;
                        break;
                    case NationalGridDimension.SideLength_10000m:
                        side_length = 10000;
                        break;
                    case NationalGridDimension.SideLength_100000m:
                        side_length = 100000;
                        break;
                    default:
                        throw new Exception($"invalid gridDimension parameter supplied");
                }

                // Get the distances of the UL_X, UL_Y point from the NatGrid Envelope Upper Left point
                int x_dist_from_UL_origin = UL_X - c_NATGRID_ENV_XMIN;
                int y_dist_from_UL_origin = c_NATGRID_ENV_YMAX - UL_Y;

                int col = (x_dist_from_UL_origin / side_length) + 1;
                int row = (y_dist_from_UL_origin / side_length) + 1;

                return GetCellNumberFromRowColumn(row, col, dimension);
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return (false, -1, $"UL X: {UL_X}, UL Y: {UL_Y}, message={ex.Message}");
            }
        }

        /// <summary>
        /// row and col parameters are 1-based.  Point of origin is upper left corner of grid
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="gridDimension"></param>
        /// <returns></returns>
        public static (bool success, long cell_number, string message) GetCellNumberFromRowColumn(int row, int col, NationalGridDimension dimension)
        {
            try
            {
                // ROWS AND COLUMNS ARE ONE-BASED INDEXES

                int rowcount = 0;
                int colcount = 0;

                switch (dimension)
                {
                    case NationalGridDimension.SideLength_1m:
                        rowcount = c_NATGRID_ROWS_DIM0;
                        colcount = c_NATGRID_COLUMNS_DIM0;
                        break;
                    case NationalGridDimension.SideLength_10m:
                        rowcount = c_NATGRID_ROWS_DIM1;
                        colcount = c_NATGRID_COLUMNS_DIM1;
                        break;
                    case NationalGridDimension.SideLength_100m:
                        rowcount = c_NATGRID_ROWS_DIM2;
                        colcount = c_NATGRID_COLUMNS_DIM2;
                        break;
                    case NationalGridDimension.SideLength_1000m:
                        rowcount = c_NATGRID_ROWS_DIM3;
                        colcount = c_NATGRID_COLUMNS_DIM3;
                        break;
                    case NationalGridDimension.SideLength_10000m:
                        rowcount = c_NATGRID_ROWS_DIM4;
                        colcount = c_NATGRID_COLUMNS_DIM4;
                        break;
                    case NationalGridDimension.SideLength_100000m:
                        rowcount = c_NATGRID_ROWS_DIM5;
                        colcount = c_NATGRID_COLUMNS_DIM5;
                        break;
                    default:
                        throw new Exception($"invalid dimension parameter supplied: {dimension}");
                }

                if (row < 1 | row > rowcount)
                {
                    return (false, -1, $"invalid row {row}");
                }
                if (col < 1 | col > colcount)
                {
                    return (false, -1, $"invalid column {col}");
                }

                long cell_number = ((row - 1) * colcount) + col;

                return (true, cell_number, "success");
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return (false, -1, $"row={row}, col={col}, message={ex.Message}");
            }
        }

        public static async Task<Envelope> GetNatGridEnvelope()
        {
            try
            {
                // Get National SR
                SpatialReference NatSR = await PRZH.GetSR_PRZCanadaAlbers();

                // Construct Envelope
                return EnvelopeBuilderEx.CreateEnvelope(c_NATGRID_ENV_XMIN, c_NATGRID_ENV_YMIN, c_NATGRID_ENV_XMAX, c_NATGRID_ENV_YMAX, NatSR);
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public async static Task<(bool success, Envelope gridEnv, string message, int tilesAcross, int tilesUp)> GetNatGridBoundsFromStudyArea(Geometry geom, NationalGridDimension dimension)
        {
            try
            {
                // Get the National SR
                SpatialReference NatSR = await PRZH.GetSR_PRZCanadaAlbers();

                #region VALIDATION

                // Ensure that geometry is not null or empty
                if (geom == null || geom.IsEmpty)
                {
                    return (false, null, "Geometry is null or empty", 0, 0);
                }

                // Ensure the geometry is either an envelope or a polygon
                if (!(geom is Envelope | geom is Polygon))
                {
                    return (false, null, "Geometry is not of type envelope or polygon", 0, 0);
                }

/*                // Simplify the geometry
                if (!GeometryEngine.Instance.IsSimpleAsFeature(geom))
                {
                    geom = GeometryEngine.Instance.SimplifyAsFeature(geom);
                }*/

                // Project the geometry if required
                SpatialReference geomSR = geom.SpatialReference;

                if (!SpatialReference.AreEqual(NatSR, geomSR))
                {
                    geom = GeometryEngine.Instance.Project(geom, NatSR);
                }

                // Get the National Grid Envelope
                Envelope gridEnv = await GetNatGridEnvelope();

                // Get the geometry Envelope
                Envelope geomEnv = geom.Extent;

                // Get the supplied grid dimension
                int dim = (int)dimension;
                if (dim < 0 | dim > 5)
                {
                    return (false, null, "invalid grid dimension provided", 0, 0);
                }

                #endregion

                #region GENERATE ENVELOPE

                // Side length
                int side_length = (int)Math.Pow(10, dim);

                // Convert envelope coordinate values to integers
                int XMIN = (int)Math.Floor(geomEnv.XMin);
                int YMIN = (int)Math.Floor(geomEnv.YMin);
                int XMAX = (int)Math.Ceiling(geomEnv.XMax);
                int YMAX = (int)Math.Ceiling(geomEnv.YMax);

                // Adjust to specified dimension
                XMIN = (XMIN / side_length) * side_length;
                YMIN = (YMIN / side_length) * side_length;
                XMAX = (XMAX / side_length) * side_length;
                YMAX = (YMAX / side_length) * side_length;

                // Generate final values
                int XMIN_NEW = (XMIN > geomEnv.XMin) ? XMIN - side_length : XMIN;
                int YMIN_NEW = (YMIN > geomEnv.YMin) ? YMIN - side_length : YMIN;
                int XMAX_NEW = (XMAX < geomEnv.XMax) ? XMAX + side_length : XMAX;
                int YMAX_NEW = (YMAX < geomEnv.YMax) ? YMAX + side_length : YMAX;

                // Generate the new envelope
                Envelope outputEnv = EnvelopeBuilderEx.CreateEnvelope(XMIN_NEW, YMIN_NEW, XMAX_NEW, YMAX_NEW, NatSR);

                if (outputEnv == null || outputEnv.IsEmpty)
                {
                    return (false, null, "Output envelope is null or empty", 0, 0);
                }

                // Ensure that the output envelope does not lie outside the national grid outer bounds
                if (XMIN_NEW < c_NATGRID_ENV_XMIN | YMIN_NEW < c_NATGRID_ENV_YMIN | XMAX_NEW > c_NATGRID_ENV_XMAX | YMAX_NEW > c_NATGRID_ENV_YMAX)
                {
                    return (false, null, "Output envelope lies wholly or partly outside of the National Grid outer bounds", 0, 0);
                }

                // Determine row and column count
                int outputWidth = Convert.ToInt32(outputEnv.Width);
                int outputHeight = Convert.ToInt32(outputEnv.Height);

                if (outputWidth == 0 | outputHeight == 0)
                {
                    return (false, null, "Output envelope has either no height or no width", 0, 0);
                }

                int tiles_across = outputWidth / side_length;       // integer division
                int tiles_up = outputHeight / side_length;          // integer division

                if (tiles_across == 0 | tiles_up == 0)
                {
                    return (false, null, "Unable to determine tile counts", 0, 0);
                }

                #endregion

                return (true, outputEnv, "Success", tiles_across, tiles_up);
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return (false, null, ex.Message, 0, 0);
            }
        }

        public static (bool success, HashSet<int> tiles, string message) GetTilesFromExtent(Envelope extent, NationalGridDimension dimension = NationalGridDimension.SideLength_1000m)
        {
            try
            {
                HashSet<int> tiles = new HashSet<int>();
                int side_length = 0;

                switch (dimension)
                {
                    case NationalGridDimension.SideLength_1m:
                        side_length = 1;
                        break;
                    case NationalGridDimension.SideLength_10m:
                        side_length = 10;
                        break;
                    case NationalGridDimension.SideLength_100m:
                        side_length = 100;
                        break;
                    case NationalGridDimension.SideLength_1000m:
                        side_length = 1000;
                        break;
                    case NationalGridDimension.SideLength_10000m:
                        side_length = 10000;
                        break;
                    case NationalGridDimension.SideLength_100000m:
                        side_length = 100000;
                        break;
                    default:
                        throw new Exception($"invalid gridDimension parameter supplied");
                }

                // Find extent in tile units
                int col_left = ((((int)extent.XMin - c_NATGRID_ENV_XMIN) / side_length) + 1) / c_NATGRID_TILE_COL_WIDTH;
                int col_right = ((((int)extent.XMax - c_NATGRID_ENV_XMIN) / side_length) + 1) / c_NATGRID_TILE_COL_WIDTH;
                int row_top = (((c_NATGRID_ENV_YMAX - (int)extent.YMax) / side_length) + 1) / c_NATGRID_TILE_ROW_HEIGHT;
                int row_bottom = (((c_NATGRID_ENV_YMAX - (int)extent.YMin) / side_length) + 1) / c_NATGRID_TILE_ROW_HEIGHT;

                // Add 1-indexed tile ids to list of intersecting tiles
                for(int row = row_top; row <= row_bottom; row++)
                {
                    for(int col = col_left; col <= col_right; col++)
                    {
                        tiles.Add(row * c_NATGRID_TILE_COLS_COUNT + col + 1);
                    }
                }


                return (true, tiles, "success");
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return (false, null, $"Error identifying national grid tiles intersecting study are. Message={ex.Message}");
            }
        }

        public static Dictionary<int, Dictionary<long, int>> GetTilesFromCells(Dictionary<long, int> cnpuid_dict, NationalGridDimension dimension = NationalGridDimension.SideLength_1000m)
        {
            int colcount = 0;

            switch (dimension)
            {
                case NationalGridDimension.SideLength_1m:
                    colcount = c_NATGRID_COLUMNS_DIM0;
                    break;
                case NationalGridDimension.SideLength_10m:
                    colcount = c_NATGRID_COLUMNS_DIM1;
                    break;
                case NationalGridDimension.SideLength_100m:
                    colcount = c_NATGRID_COLUMNS_DIM2;
                    break;
                case NationalGridDimension.SideLength_1000m:
                    colcount = c_NATGRID_COLUMNS_DIM3;
                    break;
                case NationalGridDimension.SideLength_10000m:
                    colcount = c_NATGRID_COLUMNS_DIM4;
                    break;
                case NationalGridDimension.SideLength_100000m:
                    colcount = c_NATGRID_COLUMNS_DIM5;
                    break;
                default:
                    throw new Exception($"invalid dimension parameter supplied: {dimension}");
            }

            Dictionary<int, Dictionary<long, int>> tiles = new Dictionary<int, Dictionary<long, int>>();

            foreach (KeyValuePair<long, int> cnpuid in cnpuid_dict)
            {
                // Convert cell number to tile
                int row = (int)(cnpuid.Key / colcount) / c_NATGRID_TILE_ROW_HEIGHT;
                int col = (int)(cnpuid.Key % colcount) / c_NATGRID_TILE_COL_WIDTH;
                int tile = row * c_NATGRID_TILE_COLS_COUNT + col + 1; // Note tile numbers are 1-indexed

                // Update dictionary of tiles and corresponding cell numbers
                if (!tiles.ContainsKey(tile))
                {
                    tiles.Add(tile, new Dictionary<long, int>());
                }
                tiles[tile].Add(cnpuid.Key, cnpuid.Value);
            }

            return tiles;
        }

        public static Dictionary<int, HashSet<long>> GetTilesFromCells(HashSet<long> cells, NationalGridDimension dimension = NationalGridDimension.SideLength_1000m)
        {
            int colcount = 0;

            switch (dimension)
            {
                case NationalGridDimension.SideLength_1m:
                    colcount = c_NATGRID_COLUMNS_DIM0;
                    break;
                case NationalGridDimension.SideLength_10m:
                    colcount = c_NATGRID_COLUMNS_DIM1;
                    break;
                case NationalGridDimension.SideLength_100m:
                    colcount = c_NATGRID_COLUMNS_DIM2;
                    break;
                case NationalGridDimension.SideLength_1000m:
                    colcount = c_NATGRID_COLUMNS_DIM3;
                    break;
                case NationalGridDimension.SideLength_10000m:
                    colcount = c_NATGRID_COLUMNS_DIM4;
                    break;
                case NationalGridDimension.SideLength_100000m:
                    colcount = c_NATGRID_COLUMNS_DIM5;
                    break;
                default:
                    throw new Exception($"invalid dimension parameter supplied: {dimension}");
            }

            Dictionary<int, HashSet<long>> tiles = new Dictionary<int, HashSet<long>>();

            foreach (long cn in cells) 
            {
                // Convert cell number to tile
                int row = (int)(cn / colcount) / c_NATGRID_TILE_ROW_HEIGHT;
                int col = (int)(cn % colcount) / c_NATGRID_TILE_COL_WIDTH;
                int tile = row * c_NATGRID_TILE_COLS_COUNT + col + 1; // Note tile numbers are 1-indexed

                // Update dictionary of tiles and corresponding cell numbers
                if(!tiles.ContainsKey(tile))
                {
                    tiles.Add(tile, new HashSet<long>());
                }
                tiles[tile].Add(cn);
            }

            return tiles;
        }

        #endregion

        #region FORMER INSTANCE STUFF

        #region FORMER INSTANCE CONSTRUCTORS

        //public NationalGrid(Polygon gridPoly)
        //{
        //    try
        //    {
        //        var res = GenerateFromPolygon(gridPoly);

        //        _cellIsValid = res.success;
        //        _constructorMessage = res.message;
        //    }
        //    catch (Exception ex)
        //    {
        //        ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
        //    }
        //}

        //public NationalGrid(string identifier)
        //{
        //    try
        //    {
        //        var res = GenerateFromIdentifier(identifier);

        //        _cellIsValid = res.success;
        //        _constructorMessage = res.message;
        //    }
        //    catch (Exception ex)
        //    {
        //        ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
        //    }
        //}

        #endregion

        #region FORMER INSTANCE FIELDS

        //private string _cellIdentifier;
        //private bool _cellIsValid = false;
        //private string _constructorMessage;
        //private Polygon _cellPolygon;
        //private Envelope _cellEnvelope;
        //private MapPoint _cellCenterPoint;
        //private double _cellSideLength;
        //private double _cellArea;
        //private double _cell_MinX;
        //private double _cell_MinY;
        //private double _cell_MaxX;
        //private double _cell_MaxY;

        #endregion

        #region FORMER INSTANCE PROPERTIES

        //public Polygon CellPolygon
        //{
        //    get => _cellPolygon;
        //}
        //public Envelope CellEnvelope
        //{
        //    get => _cellEnvelope;
        //}
        //public MapPoint CellCenterPoint
        //{
        //    get => _cellCenterPoint;
        //}
        //public bool CellIsValid
        //{
        //    get => _cellIsValid;
        //}
        //public string ConstructorMessage
        //{
        //    get => _constructorMessage;
        //}
        //public double CellSideLength
        //{
        //    get => _cellSideLength;
        //}
        //public string CellIdentifier
        //{
        //    get => _cellIdentifier;
        //}
        //public double Cell_MinX
        //{
        //    get => _cell_MinX;
        //}
        //public double Cell_MinY
        //{
        //    get => _cell_MinY;
        //}
        //public double Cell_MaxX
        //{
        //    get => _cell_MaxX;
        //}
        //public double Cell_MaxY
        //{
        //    get => _cell_MaxY;
        //}
        //public double CellArea
        //{
        //    get => _cellArea;
        //}

        #endregion

        #region FORMER INSTANCE METHODS

        //private (bool success, string message) GenerateFromPolygon(Polygon cellPoly)
        //{
        //    try
        //    {
        //        // Ensure no null or empty geometry provided
        //        if (cellPoly == null || cellPoly.IsEmpty)
        //        {
        //            return (false, "Cell geometry is null or empty");
        //        }

        //        // Polygon must have valid spatial reference
        //        SpatialReference cellSR = cellPoly.SpatialReference;

        //        if (cellSR == null)
        //        {
        //            return (false, "Cell geometry has null spatial reference");
        //        }

        //        // Polygon spatial reference must be same as National Grid spatial reference
        //        if (!SpatialReference.AreEqual(CANADA_ALBERS_SR, cellSR))
        //        {
        //            return (false, "Cell geometry has incorrect spatial reference");
        //        }

        //        // Polygon should have exactly 5 points (start and end points are coincident)
        //        var cellVertices = cellPoly.Points;
        //        if (cellPoly.PointCount != 5)
        //        {
        //            return (false, "Cell geometry has more than 5 vertices");
        //        }

        //        // Get Polygon Envelope
        //        Envelope cellEnv = cellPoly.Extent;

        //        if (cellEnv == null || cellEnv.IsEmpty)
        //        {
        //            return (false, "Cell geometry envelope is null or empty");
        //        }

        //        // Convert Envelope to Polygon
        //        Polygon cellEnvPoly = PolygonBuilderEx.CreatePolygon(cellEnv, CANADA_ALBERS_SR);

        //        if (cellEnvPoly == null || cellEnvPoly.IsEmpty)
        //        {
        //            return (false, "Envelope to Polygon produces null or empty geometry");
        //        }

        //        // Compare polygons.  They must be identical
        //        bool are_equal = GeometryEngine.Instance.Equals(cellPoly, cellEnvPoly);

        //        if (!are_equal)
        //        {
        //            return (false, "Cell geometry is not equal to its own envelope");
        //        }

        //        // Test squareness by comparing areas (i know geometry is rectangular and axis-aligned with SR)
        //        double perimeter = cellPoly.Length;
        //        double side_length = perimeter / 4.0;
        //        double area_from_sides = side_length * side_length;
        //        double area_from_poly = cellPoly.Area;

        //        if (area_from_sides != area_from_poly)
        //        {
        //            return (false, "Cell geometry is not square");
        //        }

        //        // Ensure that the square side length is 1, 10, 100, 1000, or 10000 meters
        //        if (side_length != 1 & side_length != 10 & side_length != 100 & side_length != 1000 & side_length != 10000 & side_length != 100000)
        //        {
        //            return (false, "Cell geometry has invalid side length for the national grid");
        //        }

        //        // Now assess the square's alignment
        //        double MinX = cellEnv.XMin;
        //        double MinY = cellEnv.YMin;

        //        double remainderX = Math.Abs(MinX % side_length);
        //        double remainderY = Math.Abs(MinY % side_length);

        //        if (remainderX != 0 & remainderY != 0)
        //        {
        //            return (false, $"Cell geometry is misaligned with the {side_length} meter National Grid along both the X and Y axes");
        //        }
        //        else if (remainderX != 0)
        //        {
        //            return (false, $"Cell geometry is misaligned with the {side_length} meter National Grid along the X axis");
        //        }
        //        else if (remainderY != 0)
        //        {
        //            return (false, $"Cell geometry is misaligned with the {side_length} meter National Grid along the Y axis");
        //        }

        //        // Ensure that polygon does not lie even partly outside the National Grid outer bounds
        //        if (cellEnv.XMin < c_NATGRID_ENV_XMIN |
        //            cellEnv.XMax > c_NATGRID_ENV_XMAX |
        //            cellEnv.YMin < c_NATGRID_ENV_YMIN |
        //            cellEnv.YMax > c_NATGRID_ENV_YMAX)
        //        {
        //            return (false, "Cell geometry falls outside the National Grid outer bounds");
        //        }

        //        // Cell Geometry is fully validated!

        //        // Populate Instance Fields
        //        _cellPolygon = cellPoly;
        //        _cellEnvelope = cellEnv;
        //        _cellSideLength = side_length;
        //        _cellCenterPoint = cellEnv.Center;
        //        _cellArea = cellPoly.Area;
        //        _cell_MinX = cellEnv.XMin;
        //        _cell_MinY = cellEnv.YMin;
        //        _cell_MaxX = cellEnv.XMax;
        //        _cell_MaxY = cellEnv.YMax;

        //        // generate identifier from lower left coords and side length
        //        string Xsuffix = (cellEnv.XMin < 0) ? "W" : "E";

        //        int xmin = Convert.ToInt32(cellEnv.XMin);
        //        int ymin = Convert.ToInt32(cellEnv.YMin);

        //        string abscissa = Math.Abs(xmin).ToString("D7");
        //        string ordinate = ymin.ToString("D7");

        //        string dimension = "";

        //        switch(side_length)
        //        {
        //            case 1:
        //                dimension = "0";
        //                break;
        //            case 10:
        //                dimension = "1";
        //                break;
        //            case 100:
        //                dimension = "2";
        //                break;
        //            case 1000:
        //                dimension = "3";
        //                break;
        //            case 10000:
        //                dimension = "4";
        //                break;
        //            case 100000:
        //                dimension = "5";
        //                break;
        //        }

        //        string identifier = abscissa + Xsuffix + "_" + ordinate + "_" + dimension;
        //        _cellIdentifier = identifier;

        //        return (true, "success");
        //    }
        //    catch (Exception ex)
        //    {
        //        ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
        //        return (false, ex.Message);
        //    }
        //}

        //private (bool success, string message) GenerateFromIdentifier(string identifier)
        //{
        //    try
        //    {
        //        // Variables
        //        string identifier_element_1;         // Element 1 of parsed identifier (the X Coordinate + E/W modifier)
        //        string identifier_element_2;         // Element 2 of parsed identifier (The Y Coordinate)
        //        string identifier_element_3;         // Element 3 of parsed identifier (the grid size indicator (0, 1, 2, 3, 4, or 5)

        //        string x_coord_string;
        //        bool xIsPositive;

        //        int x_coord;                    // Canada Albers X coordinate (-9,999,999 to 9,999,999)
        //        int y_coord;                    // Canada Albers Y coordinate (0 to 9,999,999)
        //        int grid_size;                  // NCC National Grid scale (0=1m, 1=10m, 2=100m, 3=1km, 4=10km, 5=100km
        //        int side_length;                // length of the side of a single grid cell in meters, at the associated scale (grid_size)

        //        // Length validation (total length of identifier should always be 18 chars)
        //        if (identifier.Length != 18)
        //        {
        //            return (false, "Identifier string is not 18 characters in length");
        //        }

        //        string[] parsed = identifier.Split('_');    // should contain 3 elements

        //        // parsed array should contain 3 elements
        //        if (parsed.Length != 3)
        //        {
        //            return (false, "Identifier split on '_' does not have 3 elements");
        //        }

        //        // validate the x string (e.g. "3199000E")
        //        identifier_element_1 = parsed[0];
        //        if (identifier_element_1.Length != 8)
        //        {
        //            return (false, "First element of identifier is not 8 characters in length");
        //        }

        //        if (identifier_element_1.EndsWith("E"))
        //        {
        //            xIsPositive = true;
        //        }
        //        else if (identifier_element_1.EndsWith("W"))
        //        {
        //            xIsPositive = false;
        //        }
        //        else
        //        {
        //            return (false, "First element of identifier does not end in 'E' or 'W'");
        //        }

        //        x_coord_string = identifier_element_1.Substring(0, 7);

        //        foreach (char c in x_coord_string)
        //        {
        //            if (c < '0' || c > '9')
        //            {
        //                return (false, "A character in the first identifier element is not a numeric digit 0 to 9");
        //            }
        //        }

        //        if (!int.TryParse(x_coord_string, out x_coord))
        //        {
        //            return (false, "First 7 digits of first element of identifier cannot be parsed to an integer value");
        //        }

        //        // Enforce need for "E" Identifier where X = 0
        //        if (x_coord == 0 && identifier_element_1.EndsWith("W"))
        //        {
        //            return (false, "X coordinate of 0 must have associated 'E' qualifier, not 'W' qualifier");
        //        }

        //        if (!xIsPositive)
        //        {
        //            x_coord *= -1;
        //        }

        //        // validate the y string (e.g. 4899000)
        //        identifier_element_2 = parsed[1];
        //        if (identifier_element_2.Length != 7)
        //        {
        //            return (false, "Second element of identifier is not 7 characters in length");
        //        }

        //        foreach (char c in identifier_element_2)
        //        {
        //            if (c < '0' || c > '9')
        //            {
        //                return (false, "A character in the second identifier element is not a numeric digit 0 to 9");
        //            }
        //        }

        //        if (!int.TryParse(identifier_element_2, out y_coord))
        //        {
        //            return (false, "The second element of the identifier cannot be parsed to an integer value");
        //        }

        //        // Validate the 3rd element (should be 0, 1, 2, 3, 4, 5)
        //        identifier_element_3 = parsed[2];

        //        if (identifier_element_3.Length != 1)
        //        {
        //            return (false, "The third element of the identifier is not 1 character in length");
        //        }

        //        foreach (char c in identifier_element_2)
        //        {
        //            if (c < '0' || c > '9')
        //            {
        //                return (false, "The third element of the identifier contains a character other than numeric digit 0-9");
        //            }
        //        }

        //        if (!int.TryParse(identifier_element_3, out grid_size))
        //        {
        //            return (false, "The third element of the identifier cannot be parsed to an integer value");
        //        }

        //        if (grid_size < 0 | grid_size > 5)
        //        {
        //            return (false, "The third element of the identifier is an integer outside the allowed range of 0-5 inclusive");
        //        }

        //        // Calculate the side length of a single grid cell
        //        side_length = (int)Math.Pow(10, grid_size);

        //        // Assemble an envelope for the grid cell
        //        Envelope cellEnv = EnvelopeBuilderEx.CreateEnvelope(x_coord, y_coord, x_coord + side_length, y_coord + side_length, CANADA_ALBERS_SR);

        //        if (cellEnv == null || cellEnv.IsEmpty)
        //        {
        //            return (false, "Grid cell envelope is null or empty");
        //        }

        //        // Create a Polygon from the envelope
        //        Polygon cellPoly = PolygonBuilderEx.CreatePolygon(cellEnv, CANADA_ALBERS_SR);

        //        if (cellPoly == null || cellPoly.IsEmpty)
        //        {
        //            return (false, "Grid cell polygon is null or empty");
        //        }

        //        // Determine spatial relationship between cell polygon and Grid Bounds
        //        var result = SpatialRelationshipWithOuterBounds(cellPoly);

        //        if (!result.success)
        //        {
        //            return (false, "Unable to determine spatial relationship between cell polygon and national grid bounds");
        //        }

        //        // Validate the spatial relationship
        //        if (result.relationship == BoundaryRelationship.GeometriesAreDisjoint)
        //        {
        //            return (false, "Cell polygon lies entirely outside the national grid bounds");
        //        }
        //        else if (result.relationship == BoundaryRelationship.GeometriesAreEqual)
        //        {
        //            return (false, "Cell polygon is identical to the national grid bounds.  What are you doing over there?");
        //        }
        //        else if (result.relationship == BoundaryRelationship.GeometriesOverlap)
        //        {
        //            return (false, "Cell polygon is partly inside, and partly outside, the national grid bounds.");
        //        }
        //        else if (result.relationship == BoundaryRelationship.GeometriesTouch)
        //        {
        //            return (false, "Cell polygon is adjacent to the national grid bounds, sharing node(s) and/or edge(s)");
        //        }
        //        else if (result.relationship == BoundaryRelationship.GeometryContainsGridEnvelope)
        //        {
        //            return (false, "Cell polygon contains the national grid bounds entirely.  What are you doing over there?");
        //        }
        //        else if (result.relationship == BoundaryRelationship.UndefinedRelationship)
        //        {
        //            return (false, "Spatial relationship between cell polygon and national grid bounds is undetermined.");
        //        }
        //        else if (result.relationship == BoundaryRelationship.GridEnvelopeContainsGeometry)
        //        {
        //            // cell polygon is valid in all ways.  Populate the various properties of this class instance

        //            _cellPolygon = cellPoly;
        //            _cellEnvelope = cellEnv;
        //            _cellSideLength = side_length;
        //            _cellCenterPoint = cellEnv.Center;
        //            _cellArea = cellPoly.Area;
        //            _cellIdentifier = identifier;
        //            _cell_MinX = cellEnv.XMin;
        //            _cell_MinY = cellEnv.YMin;
        //            _cell_MaxX = cellEnv.XMax;
        //            _cell_MaxY = cellEnv.YMax;

        //            return (true, "Cell polygon lies entirely within the national grid bounds.");
        //        }
        //        else
        //        {
        //            return (false, "Unreachable code reached.  Astounding!");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
        //        return (false, ex.Message);
        //    }
        //}




        #endregion

        #endregion

        #region FORMER STATIC STUFF

        //public enum BoundaryRelationship
        //{
        //    GeometriesAreEqual,
        //    GeometriesAreDisjoint,
        //    GridEnvelopeContainsGeometry,
        //    GeometryContainsGridEnvelope,
        //    GeometriesOverlap,
        //    GeometriesTouch,
        //    UndefinedRelationship
        //}

        //public static (bool success, BoundaryRelationship relationship, string message) SpatialRelationshipWithOuterBounds(Geometry geom)
        //{
        //    try
        //    {
        //        // no null geometry!
        //        if (geom == null || geom.IsEmpty)
        //        {
        //            return (false, BoundaryRelationship.UndefinedRelationship, "Null or empty geometry provided");
        //        }

        //        // Ensure the geometry is either an envelope or a polygon
        //        if (!(geom is Envelope | geom is Polygon))
        //        {
        //            return (false, BoundaryRelationship.UndefinedRelationship, "geometry is not of type envelope or polygon");
        //        }

        //        // Simplify the geometry
        //        if (!GeometryEngine.Instance.IsSimpleAsFeature(geom))
        //        {
        //            geom = GeometryEngine.Instance.SimplifyAsFeature(geom);
        //        }

        //        // Project the geometry if required
        //        SpatialReference geomSR = geom.SpatialReference;

        //        if (!SpatialReference.AreEqual(CANADA_ALBERS_SR, geomSR))
        //        {
        //            geom = GeometryEngine.Instance.Project(geom, CANADA_ALBERS_SR);
        //        }

        //        Geometry gridEnvelope = GetNatGridEnvelope();

        //        // Geometries are EQUAL
        //        if (GeometryEngine.Instance.Equals(gridEnvelope, geom))
        //        {
        //            return (true, BoundaryRelationship.GeometriesAreEqual, "Success");
        //        }

        //        // Geometries are DISJOINT
        //        else if (GeometryEngine.Instance.Disjoint(gridEnvelope, geom))
        //        {
        //            return (true, BoundaryRelationship.GeometriesAreDisjoint, "Success");
        //        }

        //        // Grid Envelope CONTAINS Geometry
        //        else if (GeometryEngine.Instance.Contains(gridEnvelope, geom))
        //        {
        //            return (true, BoundaryRelationship.GridEnvelopeContainsGeometry, "Success");
        //        }

        //        // Grid Envelope is WITHIN Geometry (this one is weird and unlikely)
        //        else if (GeometryEngine.Instance.Within(gridEnvelope, geom))
        //        {
        //            return (true, BoundaryRelationship.GeometryContainsGridEnvelope, "Success");
        //        }

        //        // Grid Envelope OVERLAPS Geometry
        //        else if (GeometryEngine.Instance.Overlaps(gridEnvelope, geom))
        //        {
        //            return (true, BoundaryRelationship.GeometriesOverlap, "Success");
        //        }

        //        // Grid Envelope TOUCHES Geometry
        //        else if (GeometryEngine.Instance.Touches(gridEnvelope, geom))
        //        {
        //            return (true, BoundaryRelationship.GeometriesTouch, "Success");
        //        }

        //        // Some other strange spatial relationship we want nothing to do with
        //        else
        //        {
        //            return (false, BoundaryRelationship.UndefinedRelationship, "Undefined Relationship");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
        //        return (false, BoundaryRelationship.UndefinedRelationship, ex.Message);
        //    }
        //}

        /// <summary>
        /// row and col parameters are 1-based.  Point of origin is upper left corner of grid
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="gridDimension"></param>
        /// <returns></returns>
        //public static (bool success, string identifier, string message) GetIdentifierFromRowColumn(int row, int col, NationalGridDimension dimension)
        //{
        //    try
        //    {
        //        // ROWS AND COLUMNS ARE ONE-BASED INDEXES

        //        int side_length = 0;
        //        int rowcount = 0;
        //        int colcount = 0;

        //        switch (dimension)
        //        {
        //            case NationalGridDimension.SideLength_1m:
        //                rowcount = c_NATGRID_ROWS_DIM0;
        //                colcount = c_NATGRID_COLUMNS_DIM0;
        //                side_length = 1;
        //                break;
        //            case NationalGridDimension.SideLength_10m:
        //                rowcount = c_NATGRID_ROWS_DIM1;
        //                colcount = c_NATGRID_COLUMNS_DIM1;
        //                side_length = 10;
        //                break;
        //            case NationalGridDimension.SideLength_100m:
        //                rowcount = c_NATGRID_ROWS_DIM2;
        //                colcount = c_NATGRID_COLUMNS_DIM2;
        //                side_length = 100;
        //                break;
        //            case NationalGridDimension.SideLength_1000m:
        //                rowcount = c_NATGRID_ROWS_DIM3;
        //                colcount = c_NATGRID_COLUMNS_DIM3;
        //                side_length = 1000;
        //                break;
        //            case NationalGridDimension.SideLength_10000m:
        //                rowcount = c_NATGRID_ROWS_DIM4;
        //                colcount = c_NATGRID_COLUMNS_DIM4;
        //                side_length = 10000;
        //                break;
        //            case NationalGridDimension.SideLength_100000m:
        //                rowcount = c_NATGRID_ROWS_DIM5;
        //                colcount = c_NATGRID_COLUMNS_DIM5;
        //                side_length = 100000;
        //                break;
        //            default:
        //                throw new Exception($"invalid gridDimension parameter supplied: {dimension}");
        //        }

        //        if (row < 1 | row > rowcount)
        //        {
        //            return (false, "", $"invalid row {row}");
        //        }
        //        if (col < 1 | col > colcount)
        //        {
        //            return (false, "", $"invalid column {col}");
        //        }

        //        int UL_X = c_NATGRID_ENV_XMIN + (side_length * (col - 1));
        //        int UL_Y = c_NATGRID_ENV_YMAX - (side_length * (row - 1));

        //        string identifier = GetIdentifierFromULXY(UL_X, UL_Y, dimension);

        //        return (true, identifier, "success");
        //    }
        //    catch (Exception ex)
        //    {
        //        ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
        //        return (false, "", $"row={row}, col={col}, message={ex.Message}");
        //    }
        //}

        //public static string GetIdentifierFromULXY(int UL_X, int UL_Y, NationalGridDimension dimension)
        //{
        //    try
        //    {
        //        int side_length = 0;

        //        switch (dimension)
        //        {
        //            case NationalGridDimension.SideLength_1m:
        //                side_length = 1;
        //                break;
        //            case NationalGridDimension.SideLength_10m:
        //                side_length = 10;
        //                break;
        //            case NationalGridDimension.SideLength_100m:
        //                side_length = 100;
        //                break;
        //            case NationalGridDimension.SideLength_1000m:
        //                side_length = 1000;
        //                break;
        //            case NationalGridDimension.SideLength_10000m:
        //                side_length = 10000;
        //                break;
        //            case NationalGridDimension.SideLength_100000m:
        //                side_length = 100000;
        //                break;
        //            default:
        //                return "";
        //        }

        //        int LL_X = UL_X;
        //        int LL_Y = UL_Y - side_length;

        //        string identifier = GetIdentifierFromLLXY(LL_X, LL_Y, dimension);

        //        // do anything here with the identifier?

        //        return identifier;
        //    }
        //    catch (Exception ex)
        //    {
        //        ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
        //        return "";
        //    }
        //}

        //public static string GetIdentifierFromLLXY(int LL_X, int LL_Y, NationalGridDimension dimension)
        //{
        //    try
        //    {
        //        // TODO: Eliminate this function
        //        string x_suffix = (LL_X < 0) ? "W" : "E";

        //        string abscissa = Math.Abs(LL_X).ToString("D7");
        //        string ordinate = LL_Y.ToString("D7");
        //        string dim = ((int)dimension).ToString();

        //        string identifier = abscissa + x_suffix + "_" + ordinate + "_" + dim;
        //        return identifier;
        //    }
        //    catch (Exception ex)
        //    {
        //        ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
        //        return "";
        //    }
        //}

        #endregion

    }
}
