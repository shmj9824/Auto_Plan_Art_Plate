using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Auto_Plan_Art_Plate
{
    internal class Dataset_excel : Call_Excel
    {
        List<Plate_data_item> pdis;
        public override void Load_File()
        {
            ws_excel = wb_excel.Worksheets[1];
            pdis = new List<Plate_data_item>();
            int row_init = 2;

            while (ws_excel.Cells[row_init, 1].Value != null)
            {
                string name = Convert.ToString(ws_excel.Cells[row_init, 1].Value);
                int col = Convert.ToInt32(Convert.ToString(ws_excel.Cells[row_init, 2].Value));
                int row = Convert.ToInt32(Convert.ToString(ws_excel.Cells[row_init, 3].Value));
                int rotate = Convert.ToInt32(Convert.ToString(ws_excel.Cells[row_init, 4].Value));
                int back = Convert.ToInt32(Convert.ToString(ws_excel.Cells[row_init,5].Value));
                pdis.Add(new Plate_data_item(name, col, row, rotate, back));
                row_init++;
            }
        }
        public List<Plate_data_item> get_list_excel()
        { 
            return pdis;
        }
    }
    
    class Plate_data_item
    {
        public string name { get; set; }
        public int col { get; set; }
        public int row { get; set; }
        public int rotate { get; set; }
        public int back_board { get; set; }
        public Plate_data_item(string name,int col,int row,int rot,int back)
        { 
            this.name = name;
            this.col = col;
            this.row = row;
            rotate = rot;
            back_board = back;
        }
    }
}
