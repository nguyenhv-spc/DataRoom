using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Windows;

namespace Data
{
    public class ExporttoExcel
    {
        public  void Export(List<DataRoom> RoomInput, string fileName)
        {
            if (RoomInput == null || RoomInput.Count == 0)
                throw new Exception("ExportToExcel: Null or empty input Rooms!\n");

            var excelApp = new Excel.Application();
            Excel.Workbook ExcelWorkBook = excelApp.Workbooks.Add();
            var Sheets = ExcelWorkBook.Sheets as Excel.Sheets;
            


            ExportWall(RoomInput, Sheets);
            ExportFloor(RoomInput, Sheets);
            ExportCeiling(RoomInput, Sheets);
            ExportFurniture(RoomInput, Sheets);


            if (!string.IsNullOrEmpty(fileName))
            {
                try
                {
                    ExcelWorkBook.SaveAs(fileName);
                    excelApp.Quit();
                    MessageBox.Show("Excel file saved!");
                }
                catch (Exception ex)
                {
                    throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n" + ex.Message);
                }
            }
            else
            {
                excelApp.Visible = true;
            }
        }

        public void ExportWall(List<DataRoom> RoomInput, Excel.Sheets Sheets)
        {
            #region 'Wall

            var WallSheet = (Excel.Worksheet)Sheets.Add(Sheets[1], Type.Missing, Type.Missing, Type.Missing);
            WallSheet.Name = "Wall_Surfaces";

            // column headings
            var headings = new List<string>();
            headings.Add("Room / Group Number");
            headings.Add("Room / Group Name");
            headings.Add("Element Name");
            headings.Add("Element Surface Name");
            headings.Add("Sub Area Name");
            headings.Add("Sub Area Surface Name");
            headings.Add("Count");
            headings.Add("ID");
            headings.Add("Room ID");
            headings.Add("Surface Material");
            headings.Add("+/-");
            headings.Add("Length / Width");
            headings.Add("Height");
            headings.Add("Area");
            headings.Add("Inner Reveals");

            for (var ii = 0; ii < headings.Count; ii++)
            {
                WallSheet.Cells[1, ii + 1] = headings[ii];
            }

            // rows
            int i = 0;
            foreach (var item in RoomInput)
            {
                WallSheet.Cells[i + 2, 1] = item.RoomNumber;
                WallSheet.Cells[i + 2, 2] = item.RoomName;
                foreach (var Wall in item.WallSet)
                {
                    WallSheet.Cells[i + 2, 3] = Wall.WallName;
                    WallSheet.Cells[i + 2, 4] = Wall.WallSurfaceName;
                    WallSheet.Cells[i + 2, 5] = Wall.SubAreaName;
                    WallSheet.Cells[i + 2, 6] = Wall.SubAreaSurfaceName;
                    WallSheet.Cells[i + 2, 7] = Wall.Count;
                    WallSheet.Cells[i + 2, 8] = Wall.ID;
                    WallSheet.Cells[i + 2, 9] = "";
                    WallSheet.Cells[i + 2, 10] = Wall.SurfaceMaterial;
                    WallSheet.Cells[i + 2, 11] = "";
                    WallSheet.Cells[i + 2, 12] = Wall.LengthWidth;
                    WallSheet.Cells[i + 2, 13] = Wall.Height;
                    WallSheet.Cells[i + 2, 14] = Wall.Area;
                    WallSheet.Cells[i + 2, 15] = Wall.InnerReveals;
                    i += 1;
                }
                foreach (var Material in item.TotalofWallsurfaceoftheroom)
                {
                    WallSheet.Cells[i + 2, 10] = Material.SurfaceMaterial;
                    WallSheet.Cells[i + 2, 14] = Material.Area;
                    WallSheet.Cells[i + 2, 15] = Material.InnerReveals;
                    i += 1;
                }
                WallSheet.Cells[i + 2, 3] = "Total of all wall surfaces of the room";
                WallSheet.Cells[i + 2, 14] = item.RoomTotalAreaofWall;
                i += 1;
            }


            Excel.Range WallRange = WallSheet.UsedRange;
            WallRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            WallRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            foreach (Range item in WallSheet.Columns)
            {
                item.AutoFit();
            }
            #endregion
        }
        public void ExportFloor(List<DataRoom> RoomInput, Excel.Sheets Sheets)
        {
            #region 'Floor

            var FloorSheet = (Excel.Worksheet)Sheets.Add(Sheets[1], Type.Missing, Type.Missing, Type.Missing);
            FloorSheet.Name = "Floor_Surfaces";

            // column headings
            var headings = new List<string>();
            headings.Add("Room / Group Number");
            headings.Add("Room / Group Name");
            headings.Add("Element Name");
            headings.Add("Element Surface Name");
            headings.Add("Sub Area Name");
            headings.Add("Sub Area Surface Name");  
            headings.Add("ID");
            headings.Add("Count");
            headings.Add("Floor Finish");
            headings.Add("+/-");
            headings.Add("Area");
            headings.Add("Door Threshold");

            for (var ii = 0; ii < headings.Count; ii++)
            {
                FloorSheet.Cells[1, ii + 1] = headings[ii];
            }

            // rows
            int i = 0;
            foreach (var item in RoomInput)
            {
                FloorSheet.Cells[i + 2, 1] = item.RoomNumber;
                FloorSheet.Cells[i + 2, 2] = item.RoomName;
                foreach (var Floor in item.FloorSet)
                {
                    FloorSheet.Cells[i + 2, 3] = Floor.FloorName ;
                    FloorSheet.Cells[i + 2, 4] = Floor.FloorSurfaceName;
                    FloorSheet.Cells[i + 2, 5] = Floor.SubAreaName;
                    FloorSheet.Cells[i + 2, 6] = Floor.SubAreaSurfaceName;                 
                    FloorSheet.Cells[i + 2, 7] = Floor.ID;
                    FloorSheet.Cells[i + 2, 8] = Floor.Count;
                    FloorSheet.Cells[i + 2, 9] = Floor.FloorFinish;
                    FloorSheet.Cells[i + 2, 10] = "";
                    FloorSheet.Cells[i + 2, 11] = Floor.Area;
                    FloorSheet.Cells[i + 2, 12] = Floor.DoorThreshold;
                    i += 1;
                }
                foreach (var Material in item.TotalofFloorsurfaceoftheroom)
                {
                    FloorSheet.Cells[i + 2, 9] = Material.FloorFinish;
                    FloorSheet.Cells[i + 2, 11] = Material.Area;
                    FloorSheet.Cells[i + 2, 12] = Material.DoorThreshold;
                    i += 1;
                }
                FloorSheet.Cells[i + 2, 3] = "Total of all floor surfaces of the room";
                FloorSheet.Cells[i + 2, 11] = item.RoomTotalAreaofFloor;
                i += 1;
            }


            Excel.Range FloorRange = FloorSheet.UsedRange;
            FloorRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            FloorRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            foreach (Range item in FloorSheet.Columns)
            {
                item.AutoFit();
            }
            #endregion
        }
        public void ExportCeiling(List<DataRoom> RoomInput, Excel.Sheets Sheets)
        {
            #region 'Ceiling

            var CeilingSheet = (Excel.Worksheet)Sheets.Add(Sheets[1], Type.Missing, Type.Missing, Type.Missing);
            CeilingSheet.Name = "Ceiling_Surfaces";

            // column headings
            var headings = new List<string>();
            headings.Add("Room / Group Number");
            headings.Add("Room / Group Name");
            headings.Add("Element Name");
            headings.Add("Element Surface Name");
            headings.Add("Sub Area Name");
            headings.Add("Sub Area Surface Name");
            headings.Add("ID");
            headings.Add("Count");
            headings.Add("Ceiling Finish");
            headings.Add("+/-");
            headings.Add("Area");

            for (var ii = 0; ii < headings.Count; ii++)
            {
                CeilingSheet.Cells[1, ii + 1] = headings[ii];
            }

            // rows
            int i = 0;
            foreach (var item in RoomInput)
            {
                CeilingSheet.Cells[i + 2, 1] = item.RoomNumber;
                CeilingSheet.Cells[i + 2, 2] = item.RoomName;
                foreach (var Ceiling in item.CeilingSet)
                {
                    CeilingSheet.Cells[i + 2, 3] = Ceiling.CeilingName;
                    CeilingSheet.Cells[i + 2, 4] = Ceiling.CeilingSurfaceName;
                    CeilingSheet.Cells[i + 2, 5] = Ceiling.SubAreaName;
                    CeilingSheet.Cells[i + 2, 6] = Ceiling.SubAreaSurfaceName;
                    CeilingSheet.Cells[i + 2, 7] = Ceiling.ID;
                    CeilingSheet.Cells[i + 2, 8] = Ceiling.Count;
                    CeilingSheet.Cells[i + 2, 9] = Ceiling.CeilingFinish;
                    CeilingSheet.Cells[i + 2, 10] = "";
                    CeilingSheet.Cells[i + 2, 11] = Ceiling.Area;
                    i += 1;
                }
                foreach (var Material in item.TotalofCeilingsurfaceoftheroom)
                {
                    CeilingSheet.Cells[i + 2, 9] = Material.CeilingFinish;
                    CeilingSheet.Cells[i + 2, 11] = Material.Area;
                    i += 1;
                }
                CeilingSheet.Cells[i + 2, 3] = "Total of all ceiling surfaces of the room";
                CeilingSheet.Cells[i + 2, 11] = item.RoomTotalAreaofCeiling;
                i += 1;
            }


            Excel.Range FloorRange = CeilingSheet.UsedRange;
            FloorRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            FloorRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            foreach (Range item in CeilingSheet.Columns)
            {
                item.AutoFit();
            }
            #endregion
        }
        public void ExportFurniture(List<DataRoom> RoomInput, Excel.Sheets Sheets)
        {
            #region 'Furniture

            var FurnitureSheet = (Excel.Worksheet)Sheets.Add(Sheets[1], Type.Missing, Type.Missing, Type.Missing);
            FurnitureSheet.Name = "Furniture_Elements";

            // column headings
            var headings = new List<string>();
            headings.Add("Level");
            headings.Add("Room / Group Number");
            headings.Add("Room / Group Name");
            headings.Add("Category");
            headings.Add("Element Name");
            headings.Add("Picture");
            headings.Add("Description");     
            headings.Add("Count");
            headings.Add("ID");
            headings.Add("Comments");


            for (var ii = 0; ii < headings.Count; ii++)
            {
                FurnitureSheet.Cells[1, ii + 1] = headings[ii];
            }

            // rows
            int i = 0;
            foreach (var item in RoomInput)
            {
                FurnitureSheet.Cells[i + 2, 1] = item.RoomLevel;
                FurnitureSheet.Cells[i + 2, 2] = item.RoomNumber;
                FurnitureSheet.Cells[i + 2, 3] = item.RoomName;
                foreach (var Furniture in item.FurnitureSet)
                {
                    FurnitureSheet.Cells[i + 2, 4] = Furniture.Category;
                    FurnitureSheet.Cells[i + 2, 5] = Furniture.FurnitureName;
                    FurnitureSheet.Cells[i + 2, 6] = "";
                    FurnitureSheet.Cells[i + 2, 7] = "";
                    FurnitureSheet.Cells[i + 2, 8] = Furniture.Count;
                    FurnitureSheet.Cells[i + 2, 9] = Furniture.ID;
                    FurnitureSheet.Cells[i + 2, 10] = Furniture.Comments;

                    i += 1;
                }

                FurnitureSheet.Cells[i + 2, 5] = "Total of all furniture elements of the room";
                FurnitureSheet.Cells[i + 2, 8] = item.RoomTotalFurniture;
                i += 1;
            }


            Excel.Range FloorRange = FurnitureSheet.UsedRange;
            FloorRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            FloorRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            foreach (Range item in FurnitureSheet.Columns)
            {
                item.AutoFit();
            }
            #endregion
        }
    }
}