using System.Collections.Generic;

namespace Data
{
    public class GetData
    {
        public void FixData()
        {
            int Roomcount = 3;

            var Rooms = new List<DataRoom>();
            DataRoom Room = null;
            for (int i = 0; i < Roomcount; i++)
            {
                Room = new DataRoom();
                Room.RoomLevel = "Level" + i;
                Room.RoomName = "Name" + i;
                Room.RoomNumber = i;
                Room.RoomID = i;

                FixDataWall(Room);
                FixDataFloor(Room);
                FixDataCeiling(Room);
                FixDataFurniture(Room);

                Rooms.Add(Room);
            }
            Module.DataRooms = Rooms;
        }

        public void FixDataWall(DataRoom Room)
        {
            #region 'Khai bao cac Wall trong Rooms

            int Wallcount = 2;

            var Walls = new List<DataWall>();
            DataWall Wall = null;

            for (int i = 0; i < Wallcount; i++)
            {
                Wall = new DataWall();
                Wall.WallName = "WallName" + i;
                Wall.WallSurfaceName = "WallSurfaceName" + i;
                Wall.SubAreaName = "SubAreaName" + i;
                Wall.SubAreaSurfaceName = "SubAreaSurfaceName" + i;
                Wall.Count = i;
                Wall.ID = i;
                Wall.RoomID = i;
                Wall.SurfaceMaterial = "SurfaceMaterial" + i;
                Wall.PS = "";
                Wall.LengthWidth = i;
                Wall.Height = i;
                Wall.Area = i;
                Wall.InnerReveals = i;

                Walls.Add(Wall);
            }
            Room.WallSet = Walls;

            #endregion 'Khai bao cac Wall trong Rooms

            #region 'Get sufacemateril of wall

            var MatertialsofWall = new List<DataMaterials>();
            DataMaterials MatertialWall = null;
            var SurfaceMaterials = new List<string>();
            foreach (var item in Walls)
            {
                MatertialWall = new DataMaterials();
                if (SurfaceMaterials.Contains(item.SurfaceMaterial))
                {
                    foreach (var item1 in MatertialsofWall)
                    {
                        if (item1.SurfaceMaterial == item.SurfaceMaterial)
                        {
                            item1.Area = item1.Area + item.Area;
                        }
                    }
                }
                else
                {
                    SurfaceMaterials.Add(item.SurfaceMaterial);
                    MatertialWall.SurfaceMaterial = item.SurfaceMaterial;
                    MatertialWall.Area = item.Area;
                    MatertialsofWall.Add(MatertialWall);
                }
            }

            Room.TotalofWallsurfaceoftheroom = MatertialsofWall;
            foreach (var item in MatertialsofWall)
            {
                Room.RoomTotalAreaofWall = Room.RoomTotalAreaofWall + item.Area;
            }

            #endregion 'Get sufacemateril of wall
        }

        public void FixDataFloor(DataRoom Room)
        {
            #region 'Khai bao cac Floor trong Rooms

            int Floorcount = 4;

            var Floors = new List<DataFloor>();
            DataFloor Floor = null;

            for (int i = 0; i < Floorcount; i++)
            {
                Floor = new DataFloor();
                Floor.FloorName = "FloorName" + i;
                Floor.FloorSurfaceName = "FloorSurfaceName" + i;
                Floor.SubAreaName = "SubAreaName" + i;
                Floor.SubAreaSurfaceName = "SubAreaSurfaceName" + i;
                Floor.Count = i;
                Floor.ID = i;
                Floor.FloorFinish = "FloorFinish" + i;
                Floor.Area = i;
                Floor.DoorThreshold = i;

                Floors.Add(Floor);
            }
            Room.FloorSet = Floors;

            #endregion 'Khai bao cac Floor trong Rooms

            #region 'Get sufacemateril of floor

            var MatertialsofFloors = new List<DataMaterials>();
            DataMaterials MatertialFloor = null;
            var FloorFisnih = new List<string>();
            foreach (var item in Floors)
            {
                MatertialFloor = new DataMaterials();
                if (FloorFisnih.Contains(item.FloorFinish))
                {
                    foreach (var item1 in MatertialsofFloors)
                    {
                        if (item1.FloorFinish == item.FloorFinish)
                        {
                            item1.Area = item1.Area + item.Area;
                        }
                    }
                }
                else
                {
                    FloorFisnih.Add(item.FloorFinish);
                    MatertialFloor.FloorFinish = item.FloorFinish;
                    MatertialFloor.Area = item.Area;
                    MatertialFloor.DoorThreshold = item.DoorThreshold;
                    MatertialsofFloors.Add(MatertialFloor);
                }
            }

            Room.TotalofFloorsurfaceoftheroom = MatertialsofFloors;
            foreach (var item in MatertialsofFloors)
            {
                Room.RoomTotalAreaofFloor = Room.RoomTotalAreaofFloor + item.Area;
            }

            #endregion 'Get sufacemateril of floor
        }

        public void FixDataCeiling(DataRoom Room)
        {
            #region 'Khai bao cac Ceiling trong Rooms

            int Ceilingcount = 3;

            var Ceilings = new List<DataCeiling>();
            DataCeiling Ceiling = null;

            for (int i = 0; i < Ceilingcount; i++)
            {
                Ceiling = new DataCeiling();
                Ceiling.CeilingName = "CeilingName" + i;
                Ceiling.CeilingSurfaceName = "CeilingSurfaceName" + i;
                Ceiling.SubAreaName = "SubAreaName" + i;
                Ceiling.SubAreaSurfaceName = "SubAreaSurfaceName" + i;
                Ceiling.Count = i;
                Ceiling.ID = i;
                Ceiling.CeilingFinish = "CeilingFinish" + i;
                Ceiling.Area = i;

                Ceilings.Add(Ceiling);
            }

            Room.CeilingSet = Ceilings;

            #endregion 'Khai bao cac Ceiling trong Rooms

            #region 'Get sufacemateril of ceiling

            var MatertialsofCeilings = new List<DataMaterials>();
            DataMaterials MatertialsofCeiling = null;
            var CeilingFinish = new List<string>();
            foreach (var item in Ceilings)
            {
                MatertialsofCeiling = new DataMaterials();
                if (CeilingFinish.Contains(item.CeilingFinish))
                {
                    foreach (var item1 in MatertialsofCeilings)
                    {
                        if (item1.SurfaceMaterial == item.CeilingFinish)
                        {
                            item1.Area = item1.Area + item.Area;
                        }
                    }
                }
                else
                {
                    CeilingFinish.Add(item.CeilingFinish);
                    MatertialsofCeiling.CeilingFinish = item.CeilingFinish;
                    MatertialsofCeiling.Area = item.Area;
                    MatertialsofCeilings.Add(MatertialsofCeiling);
                }
            }

            Room.TotalofCeilingsurfaceoftheroom = MatertialsofCeilings;
            foreach (var item in MatertialsofCeilings)
            {
                Room.RoomTotalAreaofCeiling = Room.RoomTotalAreaofCeiling + item.Area;
            }

            #endregion 'Get sufacemateril of ceiling
        }

        public void FixDataFurniture(DataRoom Room)
        {
            #region 'Khai bao cac Furniture trong Rooms

            int Furniturecount = 2;

            var Furnitures = new List<DataFurniture>();
            DataFurniture Furniture = null;

            for (int i = 0; i < Furniturecount; i++)
            {
                Furniture = new DataFurniture();
                Furniture.Category = "Category" + i;
                Furniture.FurnitureName = "FurnitureSurfaceName" + i;
                Furniture.Description = "Description" + i;
                Furniture.Count = i;
                Furniture.ID = i;
                Furniture.Comments = "Comments" + i;

                Furnitures.Add(Furniture);
            }
            Room.FurnitureSet = Furnitures;

            #endregion 'Khai bao cac Furniture trong Rooms

            #region 'Get count of Furniture

            foreach (var item in Furnitures)
            {
                Room.RoomTotalFurniture = Room.RoomTotalFurniture + item.Count;
            }

            #endregion 'Get count of Furniture
        }
    }
}