using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Auto_Plan_Art_Plate
{
    internal class Import_Vaude_Care : Call_Excel
    {
        List<string> eng_care = new List<string>();
        public Import_Vaude_Care()
        { 
        
        }

        public override void Load_File()
        {
            ws_excel = wb_excel.Worksheets["洗語"];

            int row_index = 6;

            while (ws_excel.Cells[row_index, 5].text != "")
            { 
                eng_care.Add(ws_excel.Cells[row_index, 5].text);
                row_index++;
            }
        }
        public List<string> Get_eng_list()
        { 
            return eng_care;
        }
    }
}
