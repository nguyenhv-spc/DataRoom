using System.Collections.Generic;

namespace Data
{
    public class DataRoom
    {
        public DataRoom()
        { }

        public string RoomName { get; set; }
        public int RoomNumber { get; set; }
        public int RoomID { get; set; }
        public double RoomVolume { get; set; }
        public string RoomLevel { get; set; }
        public double RoomArea { get; set; }
        public List<DataWall> WallSet { get; set; }
        public List<DataMaterials> TotalofWallsurfaceoftheroom { get; set; }
        public double RoomTotalAreaofWall { get; set; }
        public List<DataFloor> FloorSet { get; set; }
        public List<DataMaterials> TotalofFloorsurfaceoftheroom { get; set; }
        public double RoomTotalAreaofFloor { get; set; }
        public List<DataCeiling> CeilingSet { get; set; }
        public List<DataMaterials> TotalofCeilingsurfaceoftheroom { get; set; }
        public double RoomTotalAreaofCeiling { get; set; }
        public List<DataFurniture> FurnitureSet { get; set; }
        public double RoomTotalFurniture { get; set; }
    }

    public class DataWall
    {
        public string WallName { get; set; }
        public string WallSurfaceName { get; set; }
        public string SubAreaName { get; set; }
        public string SubAreaSurfaceName { get; set; }
        public int Count { get; set; }
        public int ID { get; set; }
        public int RoomID { get; set; }
        public string SurfaceMaterial { get; set; }
        // +/-
        public string PS { get; set; }
        public double LengthWidth { get; set; }
        public double Height { get; set; }
        public double Area { get; set; }
        public int InnerReveals { get; set; }
    }

    public class DataFloor
    {
        public string FloorName { get; set; }
        public string FloorSurfaceName { get; set; }
        public string SubAreaName { get; set; }
        public string SubAreaSurfaceName { get; set; }
        public int Count { get; set; }
        public int ID { get; set; }
        public string FloorFinish { get; set; }
        public double Area { get; set; }
        public int DoorThreshold { get; set; }
    }

    public class DataCeiling
    {
        public string CeilingName { get; set; }
        public string CeilingSurfaceName { get; set; }
        public string SubAreaName { get; set; }
        public string SubAreaSurfaceName { get; set; }
        public int Count { get; set; }
        public int ID { get; set; }
        public string CeilingFinish { get; set; }
        public double Area { get; set; }
    }

    public class DataFurniture
    {
        public string Category { get; set; }
        public string FurnitureName { get; set; }
        public string Description { get; set; }
        public int Count { get; set; }
        public int ID { get; set; }
        public string Comments { get; set; }
    }

    public class DataMaterials
    {
        public string SurfaceMaterial { get; set; }
        public string FloorFinish { get; set; }
        public string CeilingFinish { get; set; }
        public double Area { get; set; }
        public int InnerReveals { get; set; }
        public int DoorThreshold { get; set; }
    }
}