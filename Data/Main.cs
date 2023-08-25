using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Data
{
    public class Main
    {
        
        public void Run()
        {
            // Get data
            GetData CreateData = new GetData();
            CreateData.FixData();


            using (SaveFileDialog SaveFD = new SaveFileDialog())
            {
                SaveFD.Title = "Lưu file Data";
                SaveFD.Filter = "Excel Files (*xlsx)|*.xlsx";
                SaveFD.RestoreDirectory = true;
                if (SaveFD.ShowDialog() == DialogResult.OK)
                {
                    ExporttoExcel Excel = new ExporttoExcel();
                    Excel.Export(Module.DataRooms, SaveFD.FileName);
                }

            }

            
        }
    }
}
