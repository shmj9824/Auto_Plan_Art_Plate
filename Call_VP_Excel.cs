using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
using Excel = Microsoft.Office.Interop.Excel;

namespace Auto_Plan_Art_Plate
{
    enum VP_Mode
    {
        check,
        plate
    }
    class VP_data
    {
        public int no { get; set; }
        public string item1 { get; set; }
        public string item2 { get; set; }
        public string item3 { get; set; }
        public string item4 { get; set; }
        public string item5 { get; set; }
        public string item6 { get; set; }
        public string item7 { get; set; }
        public string item8 { get; set; }
        public string item9 { get; set; }
    }
    internal class Call_VP_Excel : Call_Excel
    {
        VP_Mode mode_str;
        
        public Call_VP_Excel(VP_Mode mode_str)
        {
            this.mode_str = mode_str;
        }
        public override void Load_File()
        {
            if (mode_str == VP_Mode.check)
                VP_check();
            else if (mode_str == VP_Mode.plate)
                VP_plate();
        }
        public virtual void VP_check()
        {}
        public virtual void VP_plate()
        {}

        public int get_Start_row(Excel._Worksheet worksheet)
        {
            int i = 1;
            while (Convert.ToString(worksheet.Cells[i, 1].Value) != "1")
            {
                i++;
            }
            return i;
        }
        public int get_End_row(Excel._Worksheet worksheet, int start_row)
        {
            int end_r = start_row;
            while (Convert.ToString(worksheet.Cells[end_r, 1].Value) != null)
            {
                end_r++;
            }
            return --end_r;
        }
    }

    class VP_Tyvek_Excel : Call_VP_Excel
    {
        public List<VP_data> vp_Datas = new List<VP_data>();
        bool old_version = false;
        public VP_Tyvek_Excel(VP_Mode mode_str) : base(mode_str)
        { 
            
        }

        public void Set_old_version_value(bool aa)
        {
            old_version = aa;
        }
        public override void VP_check()
        {
            ws_excel = wb_excel.Worksheets[1];

            int data_start_row = get_Start_row(ws_excel);
            int data_end_row = get_End_row(ws_excel, data_start_row);

            for (int i = data_start_row; i <= data_end_row; i++)
            {
                VP_data vp_d = new VP_data { };
                if (!old_version)
                {
                    vp_d = new VP_data
                    {
                        no = Convert.ToInt32(Convert.ToString(ws_excel.Cells[i, 1].value)),

                        item1 = Convert.ToString(ws_excel.Cells[i, 2].value),
                        item2 = Convert.ToString(ws_excel.Cells[i, 4].value),
                        item3 = Convert.ToString(ws_excel.Cells[i, 5].value),
                        item4 = Convert.ToString(ws_excel.Cells[i, 6].value),
                        item5 = Convert.ToString(ws_excel.Cells[i, 7].value),
                        item9 = Convert.ToString(ws_excel.Cells[i, 8].value),
                        item6 = Convert.ToString(ws_excel.Cells[i, 9].value),
                        item7 = Convert.ToString(ws_excel.Cells[i, 10].value),
                        item8 = Convert.ToString(ws_excel.Cells[i, 3].value),
                    };
                }
                else 
                {
                    vp_d = new VP_data
                    {
                        no = Convert.ToInt32(Convert.ToString(ws_excel.Cells[i, 1].value)),
                        item1 = Convert.ToString(ws_excel.Cells[i, 2].value),
                        item2 = Convert.ToString(ws_excel.Cells[i, 4].value),
                        item3 = Convert.ToString(ws_excel.Cells[i, 5].value),
                        item4 = Convert.ToString(ws_excel.Cells[i, 6].value),
                        item5 = Convert.ToString(ws_excel.Cells[i, 7].value),
                        item6 = Convert.ToString(ws_excel.Cells[i, 8].value),
                        item7 = Convert.ToString(ws_excel.Cells[i, 9].value),
                        item8 = Convert.ToString(ws_excel.Cells[i, 3].value),
                    };
                }
                
                vp_Datas.Add(vp_d);
            }
        }

        public override void VP_plate()
        {
            ws_excel = wb_excel.Worksheets[1];

            int data_start_row = get_Start_row(ws_excel);
            int data_end_row = get_End_row(ws_excel, data_start_row);

            for (int i = data_start_row; i <= data_end_row; i++)
            {
                VP_data vp_d = new VP_data { };
                if (!old_version)
                {
                    vp_d = new VP_data
                    {
                        no = Convert.ToInt32(Convert.ToString(ws_excel.Cells[i, 1].value)),
                        item1 = Convert.ToString(ws_excel.Cells[i, 4].value),
                        item2 = Convert.ToString(ws_excel.Cells[i, 5].value),
                        item3 = Convert.ToString(ws_excel.Cells[i, 6].value),
                        item4 = Convert.ToString(ws_excel.Cells[i, 8].value),
                        item5 = Convert.ToString(ws_excel.Cells[i, 9].value),
                        item7 = Convert.ToString(ws_excel.Cells[i, 7].value)
                    };
                }
                else
                {
                    vp_d = new VP_data
                    {
                        no = Convert.ToInt32(Convert.ToString(ws_excel.Cells[i, 1].value)),
                        item1 = Convert.ToString(ws_excel.Cells[i, 4].value),
                        item2 = Convert.ToString(ws_excel.Cells[i, 5].value),
                        item3 = Convert.ToString(ws_excel.Cells[i, 6].value),
                        item4 = Convert.ToString(ws_excel.Cells[i, 7].value),
                        item5 = Convert.ToString(ws_excel.Cells[i, 8].value),
                        item6 = Convert.ToString(ws_excel.Cells[i, 9].value)
                    };
                }
                
                string get_val = Convert.ToString(ws_excel.Cells[i, 10].value);
                if (get_val != null)
                    if (get_val.Contains("PE"))
                        vp_d.item6 = get_val;
                vp_Datas.Add(vp_d);
            }
        }
    }

    
}
