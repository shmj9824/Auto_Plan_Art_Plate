using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Auto_Plan_Art_Plate
{
    class VP_Ilina_Excel : Call_VP_Excel
    {
        public List<VP_data> vp_Datas = new List<VP_data>();
        List<Ilina_vp_data> Vp_ilina_datas = new List<Ilina_vp_data>();
        public VP_Ilina_Excel(VP_Mode mode_str) : base(mode_str)
        {

        }
        /*
        public override void VP_check()
        {
            ws_excel = wb_excel.Worksheets[1];
            int data_start_row = get_Start_row(ws_excel);
            int data_end_row = get_End_row(ws_excel, data_start_row);

            for (int i = data_start_row; i <= data_end_row; i++)
            {
                VP_data vp_d = new VP_data
                {
                    no = 0,
                    item1 = Convert.ToString(ws_excel.Cells[i, 3].value),
                    item2 = Convert.ToString(ws_excel.Cells[i, 4].value),
                    item3 = Convert.ToString(ws_excel.Cells[i, 5].value),
                    item4 = Convert.ToString(ws_excel.Cells[i, 6].value),
                    item5 = Convert.ToString(ws_excel.Cells[i, 7].value),
                };
                vp_Datas.Add(vp_d);
            }
        }
        */
        public override void VP_check()
        {
            int sheet_name_init_x = 4;
            int sheet_name_init_y = 11;

            //訪問所有在檔案內的Sheet
            for (int u = 1; u <= wb_excel.Worksheets.Count; u++)
            {
                ws_excel = wb_excel.Worksheets[u];

                string file_sheet_name = Convert.ToString(ws_excel.Cells[sheet_name_init_x, sheet_name_init_y].value);
                string file_po_num;

                //抓取貨號列
                int point_row = 1;
                //抓取資料開頭列
                while (Convert.ToString(ws_excel.Cells[point_row, 1].value) != "no                     (欄位：9)")
                {
                    point_row++;
                }
                int tabel_head_row = point_row;
                int data_row = point_row + 2;
                if (Convert.ToString(ws_excel.Cells[data_row, 1].value) == null)
                    data_row++;

                //開始取資料
                int check_col = 6;

                while (Convert.ToString(ws_excel.Cells[data_row, check_col].value) != null)
                {
                    if (Convert.ToString(ws_excel.Cells[data_row, 1].value) == null)
                        data_row++;

                    if (Convert.ToString(ws_excel.Cells[data_row, check_col].value) == null)
                    {
                        if (Convert.ToString(ws_excel.Cells[data_row + 1, check_col].value) != null)
                        {
                            data_row += 1;
                            if (Convert.ToString(ws_excel.Cells[data_row, check_col].value) == "TOTAL")
                                break;
                        }
                        else
                            break;
                    }
                    //貨號
                    string product = Convert.ToString(ws_excel.Cells[data_row, 1].value);

                    //color
                    string color = Convert.ToString(ws_excel.Cells[data_row, 2].value);

                    //size
                    string size = Convert.ToString(ws_excel.Cells[data_row, 3].value);

                    //price
                    string price = Convert.ToString(ws_excel.Cells[data_row, 5].value);
                    //barcode
                    string barcode = Convert.ToString(ws_excel.Cells[data_row, 6].value);
                    //string barcode = product + "-" + color + "-" + size;

                    Vp_ilina_datas.Add(new Ilina_vp_data(file_sheet_name,product,color,size,price,barcode));
                    data_row++;
                }
            }
            
        }

        public override void VP_plate()
        {
            ws_excel = wb_excel.Worksheets[1];
            int data_start_row = get_Start_row(ws_excel);
            int data_end_row = get_End_row(ws_excel, data_start_row);

            for (int i = data_start_row; i <= data_end_row; i++)
            {
                VP_data vp_d = new VP_data
                {
                    no = Convert.ToInt32(Convert.ToString(ws_excel.Cells[i, 1].value)),
                    item1 = Convert.ToString(ws_excel.Cells[i, 4].value),
                    item2 = Convert.ToString(ws_excel.Cells[i, 5].value),
                    item3 = Convert.ToString(ws_excel.Cells[i, 6].value),
                    item4 = "$" + Convert.ToString(ws_excel.Cells[i, 7].value),
                };
                string str_item2 = vp_d.item2;
                string str_item3 = vp_d.item3;
                while (str_item2.Length < 4)
                    str_item2 += " ";
                while (str_item3.Length < 4)
                    str_item3 += " ";
                vp_d.item5 = vp_d.item1 + " " + str_item2 + str_item3;
                vp_Datas.Add(vp_d);
            }
        }
        public List<Ilina_vp_data> Get_Ilina_Vp_Datas() { return Vp_ilina_datas; }
    }

    class Ilina_vp_data
    {
        public string Sheet_name { get; set; }
        public string No_Product { get; set; }
        public string Color { get; set; }
        public string Size { get; set; }
        public string Price { get; set; }
        public string Code128 { get; set; }
        public Ilina_vp_data(string sheetName,string no_p, string color, string size, string price, string code)
        {
            Sheet_name = sheetName;
            this.No_Product = no_p;
            this.Color = color;
            this.Size = size;
            this.Price = price;
            this.Code128 = code;
        }
    }
}
