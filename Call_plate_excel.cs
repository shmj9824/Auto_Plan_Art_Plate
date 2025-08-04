using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Auto_Plan_Art_Plate
{
    class Call_plate_excel : Call_Excel
    {
        public int plate_count;
        public int data_count;
        int plate_int = 0;
        List<Dictionary<int, int>> list_plate_data = new List<Dictionary<int, int>>();
        public string mess;
        public string mess_plate_sum = "";
        public Call_plate_excel(int plate_int)
        {
            this.plate_int = plate_int;
        }
        public List<Dictionary<int, int>> Get_list_plate_data()
        { 
            return list_plate_data;
        }
        public override void Load_File()
        {
            ws_excel = wb_excel.Worksheets[1];

            //表頭列
            int table_head_row = 1;
            //總車數的行號
            int sum_data_col = 1;

            while (Convert.ToString(ws_excel.Cells[table_head_row, 1].Value) != "#")
            {
                table_head_row++;
            }

            int data_start_row = table_head_row + 1;

            while (Convert.ToString(ws_excel.Cells[1, sum_data_col].Value) != "總車數")
                sum_data_col++;
            //(G,2)
            mess = Convert.ToString(ws_excel.Cells[2, 7].Value);

            //取得資料筆數
            int data_end_row = data_start_row;
            while (ws_excel.Cells[data_end_row, 1].Value != null)
            {
                data_end_row++;
            }
            data_count = data_end_row - data_start_row;
            
            //取得模板資料的行號
            int plate_col_init = sum_data_col + 2;
            int plate_col_end = plate_col_init;
            while (ws_excel.Cells[data_start_row - 1, plate_col_end].Value != null)
            {
                if (Convert.ToString(ws_excel.Cells[data_start_row - 2, plate_col_end].Value) == null)
                    break;
                plate_col_end++;
            }
            //計算版數-最後之資料尾減頭
            plate_count = plate_col_end - plate_col_init;
            

            for (int i = plate_col_init; i < plate_col_end; i++)
            {
                int row_step = 0;
                for (int j = data_start_row; j < data_end_row; j++)
                {
                    if (Convert.ToString(ws_excel.Cells[j, i].value) != "")
                        row_step += Convert.ToInt32(ws_excel.Cells[j, i].value);
                }
                if (plate_int != row_step)
                    mess_plate_sum = "第" + i.ToString() + "列模板加總錯誤";
            }

            for (int i = plate_col_init; i < plate_col_end; i++)
            {
                Dictionary<int, int> plate_dict = new Dictionary<int, int>();
                int plan_num_sum = 0;
                for (int j = data_start_row; j < data_end_row; j++)
                {
                    int item_no = Convert.ToInt32(ws_excel.Cells[j, 1].value);

                    if (ws_excel.Cells[j, i].value != null)
                    { 
                        int plan_num = Convert.ToInt32(ws_excel.Cells[j, i].value);
                        plan_num_sum += plan_num;

                        plate_dict.Add(item_no, plan_num);
                    }
                }
                
                list_plate_data.Add(plate_dict);
            }
            
        }
    }
}
