using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Media3D;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Excel;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;

namespace Data
{
    public class ExportUsingOpenXml
    {
        WorkbookPart workbookPart_Pulic = null;
        WorkbookStylesPart stylesPart = null;
        public void ExportDataRoom(List<DataRoom> RoomInput, string fileName)
        {
            using (var workbook = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // Tạo hoặc chỉnh sửa tài liệu Excel và Stylesheet tại đây
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                #region Tạo Sheets
                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                uint sheetId = 1;
                if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = "Wall_Surfaces" };
                sheets.Append(sheet);
                #endregion

                stylesPart = workbook.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();
                NewCellFormat();

                workbookPart_Pulic = workbookPart;

                

                // Lưu tài liệu Excel và Stylesheet
            }

            using (var workbook = SpreadsheetDocument.Open(fileName, true))
            {

                stylesPart = workbook.WorkbookPart.WorkbookStylesPart;
                Stylesheet stylesheet = stylesPart.Stylesheet;
                //var workbookPart = workbook.WorkbookPart;

                //ExportWall(RoomInput, workbook);
                //ExportFloor(RoomInput, workbook);
                //ExportCeiling(RoomInput, workbook);
                //ExportFurniture(RoomInput, workbook);

            }
        }
        public void ExportWall(List<DataRoom> RoomInput, SpreadsheetDocument workbook)
        {
            #region Tạo Sheets
            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
            sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

            DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
            string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

            uint sheetId = 1;
            if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = "Wall_Surfaces" };
            sheets.Append(sheet);
            #endregion
            #region Thêm hàng header
            DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

            List<String> columns = new List<string>();

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

            foreach (var item in headings)
            {
                columns.Add(item);

                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(item);
                headerRow.AppendChild(cell);
            }

            sheetData.AppendChild(headerRow);
            #endregion
            #region Thêm data
            foreach (var item in RoomInput)
            {     
                #region Thêm hàng elements
                foreach (var Wall in item.WallSet)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRowData = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    AddNewRow( item.RoomNumber.ToString(), newRowData);
                    AddNewRow(item.RoomName.ToString(), newRowData);
                    AddNewRow(Wall.WallName.ToString(), newRowData);
                    AddNewRow(Wall.WallSurfaceName.ToString(), newRowData);
                    AddNewRow(Wall.SubAreaName.ToString(), newRowData);
                    AddNewRow(Wall.SubAreaSurfaceName.ToString(), newRowData);
                    AddNewRow(Wall.Count.ToString(), newRowData);
                    AddNewRow(Wall.ID.ToString(), newRowData);
                    AddNewRow("", newRowData);
                    AddNewRow(Wall.SurfaceMaterial.ToString(), newRowData);
                    AddNewRow("", newRowData);
                    AddNewRow(Wall.LengthWidth.ToString(), newRowData);
                    AddNewRow(Wall.Height.ToString(), newRowData);
                    AddNewRow(Wall.Area.ToString(), newRowData);
                    AddNewRow(Wall.InnerReveals.ToString(), newRowData);
                    sheetData.AppendChild(newRowData);
                }
                #endregion
                #region Them hàng vật liệu
                foreach (var Material in item.TotalofWallsurfaceoftheroom)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRowMaterial = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow(Material.SurfaceMaterial.ToString(), newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow(Material.Area.ToString(), newRowMaterial);
                    AddNewRow(Material.InnerReveals.ToString(), newRowMaterial);
                    sheetData.AppendChild(newRowMaterial);
                }
                #endregion
                #region Thêm hàng total
                DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("Total of all wall surfaces of the room", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow(item.RoomTotalAreaofWall.ToString(), newRow);
                AddNewRow("", newRow);

                sheetData.AppendChild(newRow);
                #endregion
            }
            #endregion

            
        }
        public void ExportFloor(List<DataRoom> RoomInput, SpreadsheetDocument workbook)
        {
            #region Tạo Sheets
            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
            sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

            DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
            string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

            uint sheetId = 1;
            if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = "Floor_Surfaces" };
            sheets.Append(sheet);
            #endregion
            #region Thêm hàng header
            DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

            List<String> columns = new List<string>();

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

            foreach (var item in headings)
            {
                columns.Add(item);

                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(item);
                headerRow.AppendChild(cell);
            }

            sheetData.AppendChild(headerRow);
            #endregion
            #region Thêm data
            foreach (var item in RoomInput)
            {
                #region Thêm hàng elements
                foreach (var Floor in item.FloorSet)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRowData = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    AddNewRow(item.RoomNumber.ToString(), newRowData);
                    AddNewRow(item.RoomName.ToString(), newRowData);

                    AddNewRow(Floor.FloorName.ToString(), newRowData);
                    AddNewRow(Floor.FloorSurfaceName.ToString(), newRowData);
                    AddNewRow(Floor.SubAreaName.ToString(), newRowData);
                    AddNewRow(Floor.SubAreaSurfaceName.ToString(), newRowData);
                    AddNewRow(Floor.ID.ToString(), newRowData);
                    AddNewRow(Floor.Count.ToString(), newRowData);
                    AddNewRow(Floor.FloorFinish, newRowData);
                    AddNewRow("", newRowData);
                    AddNewRow(Floor.Area.ToString(), newRowData);
                    AddNewRow(Floor.DoorThreshold.ToString(), newRowData);
                    sheetData.AppendChild(newRowData);
                }
                #endregion
                #region Them hàng vật liệu
                foreach (var Material in item.TotalofFloorsurfaceoftheroom)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRowMaterial = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow(Material.FloorFinish.ToString(), newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow(Material.Area.ToString(), newRowMaterial);
                    AddNewRow(Material.DoorThreshold.ToString(), newRowMaterial);
                    sheetData.AppendChild(newRowMaterial);
                }
                #endregion
                #region Thêm hàng total
                DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("Total of all floor surfaces of the room", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);      
                AddNewRow(item.RoomTotalAreaofFloor.ToString(), newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                sheetData.AppendChild(newRow);
                #endregion
            }
            #endregion
        }
        public void ExportCeiling(List<DataRoom> RoomInput, SpreadsheetDocument workbook)
        {
            #region Tạo Sheets
            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
            sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

            DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
            string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

            uint sheetId = 1;
            if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = "Ceiling_Surfaces" };
            sheets.Append(sheet);
            #endregion
            #region Thêm hàng header
            DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

            List<String> columns = new List<string>();

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

            foreach (var item in headings)
            {
                columns.Add(item);

                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(item);
                headerRow.AppendChild(cell);
            }

            sheetData.AppendChild(headerRow);
            #endregion
            #region Thêm data
            foreach (var item in RoomInput)
            {
                #region Thêm hàng elements
                foreach (var Ceiling in item.CeilingSet)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRowData = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    AddNewRow(item.RoomNumber.ToString(), newRowData);
                    AddNewRow(item.RoomName.ToString(), newRowData);

                    AddNewRow(Ceiling.CeilingName.ToString(), newRowData);
                    AddNewRow(Ceiling.CeilingSurfaceName.ToString(), newRowData);
                    AddNewRow(Ceiling.SubAreaName.ToString(), newRowData);
                    AddNewRow(Ceiling.SubAreaSurfaceName.ToString(), newRowData);
                    AddNewRow(Ceiling.ID.ToString(), newRowData);
                    AddNewRow(Ceiling.Count.ToString(), newRowData);
                    AddNewRow(Ceiling.CeilingFinish.ToString(), newRowData);
                    AddNewRow("", newRowData);
                    AddNewRow(Ceiling.Area.ToString(), newRowData);
                    sheetData.AppendChild(newRowData);
                }
                #endregion
                #region Them hàng vật liệu
                foreach (var Material in item.TotalofCeilingsurfaceoftheroom)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRowMaterial = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow(Material.CeilingFinish.ToString(), newRowMaterial);
                    AddNewRow("", newRowMaterial);
                    AddNewRow(Material.Area.ToString(), newRowMaterial);
                    sheetData.AppendChild(newRowMaterial);
                }
                #endregion
                #region Thêm hàng total
                DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("Total of all ceiling surfaces of the room", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow(item.RoomTotalAreaofCeiling.ToString(), newRow);
                AddNewRow("", newRow);
                sheetData.AppendChild(newRow);
                #endregion
            }
            #endregion
            
        }
        public void ExportFurniture(List<DataRoom> RoomInput, SpreadsheetDocument workbook)
        {
            #region Tạo Sheets
            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
            sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

            DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
            string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

            uint sheetId = 1;
            if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = "Furniture_Elements" };
            sheets.Append(sheet);
            #endregion
            #region Thêm hàng header
            DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

            List<String> columns = new List<string>();

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

            foreach (var item in headings)
            {
                columns.Add(item);

                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(item);
                headerRow.AppendChild(cell);
            }

            sheetData.AppendChild(headerRow);
            #endregion
            #region Thêm data
            foreach (var item in RoomInput)
            {
                #region Thêm hàng elements
                foreach (var Furniture in item.FurnitureSet)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRowData = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    AddNewRow(item.RoomLevel.ToString(), newRowData);
                    AddNewRow(item.RoomNumber.ToString(), newRowData);
                    AddNewRow(item.RoomName.ToString(), newRowData);

                    AddNewRow(Furniture.Category.ToString(), newRowData);
                    AddNewRow(Furniture.FurnitureName.ToString(), newRowData);
                    AddNewRow("", newRowData);
                    AddNewRow("", newRowData);
                    AddNewRow(Furniture.Count.ToString(), newRowData);
                    AddNewRow(Furniture.ID.ToString(), newRowData);
                    AddNewRow(Furniture.Comments.ToString(), newRowData);
                    sheetData.AppendChild(newRowData);
                }
                #endregion
            
                #region Thêm hàng total
                DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow("Total of all furniture elements of the room", newRow);
                AddNewRow("", newRow);
                AddNewRow("", newRow);
                AddNewRow(item.RoomTotalFurniture.ToString(), newRow);
                sheetData.AppendChild(newRow);
                #endregion
            }
            #endregion
        }
        
        public void AddNewRow(string value, DocumentFormat.OpenXml.Spreadsheet.Row newRow)
        {
            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value);
            //NewCellFormat();
            cell.StyleIndex = 1;
            newRow.AppendChild(cell);
        }
        public DocumentFormat.OpenXml.Spreadsheet.CellFormat NewCellFormat()
        {
            // Tạo 1 danh sách CellFormats, danh sách Fonts, danh sách Boders
            stylesPart.Stylesheet.CellFormats = new DocumentFormat.OpenXml.Spreadsheet.CellFormats();
            stylesPart.Stylesheet.Fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts();
            stylesPart.Stylesheet.Borders = new DocumentFormat.OpenXml.Spreadsheet.Borders();

            // Tạo 1 đối tượng CellFormat
            DocumentFormat.OpenXml.Spreadsheet.CellFormat cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();

            // Tạo 1 đối tượng Font
            DocumentFormat.OpenXml.Spreadsheet.Font font = new DocumentFormat.OpenXml.Spreadsheet.Font();
            font.FontName = new DocumentFormat.OpenXml.Spreadsheet.FontName { Val = "Time New Roman" };
            font.FontSize = new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 13 };
            font.Color = new Color { Rgb = "FF0000" }; // Màu đỏ


            // Tạo 1 đối tượng Border
            DocumentFormat.OpenXml.Spreadsheet.Border border = new DocumentFormat.OpenXml.Spreadsheet.Border();

            DocumentFormat.OpenXml.Spreadsheet.LeftBorder leftBorder = new DocumentFormat.OpenXml.Spreadsheet.LeftBorder() { Style = BorderStyleValues.Thick };
            Color color1 = new Color() { Indexed = (UInt32Value)64U };
            leftBorder.Append(color1);

            DocumentFormat.OpenXml.Spreadsheet.RightBorder rightBorder = new DocumentFormat.OpenXml.Spreadsheet.RightBorder() { Style = BorderStyleValues.Dotted };
            Color color2 = new Color() { Indexed = (UInt32Value)64U };
            rightBorder.Append(color2);

            DocumentFormat.OpenXml.Spreadsheet.TopBorder topBorder = new DocumentFormat.OpenXml.Spreadsheet.TopBorder() { Style = BorderStyleValues.Dashed };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };
            topBorder.Append(color3);

            DocumentFormat.OpenXml.Spreadsheet.BottomBorder bottomBorder = new DocumentFormat.OpenXml.Spreadsheet.BottomBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };
            bottomBorder.Append(color4);

            DiagonalBorder diagonalBorder = new DiagonalBorder();

            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);
            border.Append(bottomBorder);
            border.Append(diagonalBorder);

            // Tạo 1 đối tượng Aignment
            DocumentFormat.OpenXml.Spreadsheet.Alignment alignment = new DocumentFormat.OpenXml.Spreadsheet.Alignment();
            alignment.Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Left;
            alignment.Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top;
            alignment.WrapText = true;

            // Thêm các format vào Stylesheet           
            stylesPart.Stylesheet.Fonts.AppendChild(font);
            stylesPart.Stylesheet.Borders.AppendChild(border);

            //Gán các ID của các Format cho CellFormat           
            cellFormat.Alignment = alignment;
            cellFormat.ApplyAlignment = true;

            cellFormat.FontId = 1;
            //cellFormat.ApplyFont = true;

            cellFormat.BorderId = 1;
            cellFormat.ApplyBorder = true;

            //cellFormat.NumberFormatId = 0; // Kiểu của cell - 0 = Mặc định

            //Thêm CellFormat vào danh sách CellFormats
            stylesPart.Stylesheet.CellFormats.AppendChild(cellFormat);
            return cellFormat;
        }
    }
}
