using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace Auto_Plan_Art_Plate
{

    class VP_Bonbon_Excel : Call_VP_Excel
    {
        //public string _test_str;
        public List<Bonbon_vp_data> bonbon_Vps = new List<Bonbon_vp_data>();
        public Dictionary<string,int> dic_art_no_col = new Dictionary<string,int>();
        public Dictionary<string, (string, string)> dic_po_no = new Dictionary<string, (string, string)>();
        Dictionary<string, (string, string)> dic_Ingredient = new Dictionary<string, (string, string)>();
        public VP_Bonbon_Excel(VP_Mode mode_str) : base(mode_str)
        { }
        public VP_Bonbon_Excel(VP_Mode mode_str, Dictionary<string, (string, string)> include1) : base(mode_str)
        { 
            this.dic_Ingredient = include1;
        }
        public override void VP_check()
        {
            //tuple<x,y>
            Tuple<int, int> sheet_name_init = new Tuple<int, int>(4, 8);
            Tuple<int, int> po_num_init = new Tuple<int, int>(4, 11);

            //訪問所有在檔案內的Sheet
            for (int u = 1; u <= wb_excel.Worksheets.Count; u++)
            {
                ws_excel = wb_excel.Worksheets[u];

                string file_sheet_name = Convert.ToString(ws_excel.Cells[sheet_name_init.Item1,sheet_name_init.Item2].value);
                string file_po_num;
                if (Convert.ToString(ws_excel.Cells[po_num_init.Item1, po_num_init.Item2].value) != "PO#")
                    file_po_num = Convert.ToString(ws_excel.Cells[po_num_init.Item1, po_num_init.Item2].value);
                else
                    file_po_num = Convert.ToString(ws_excel.Cells[po_num_init.Item1, po_num_init.Item2 + 1].value);

                //抓取貨號列
                int point_row = 1;
                while (Convert.ToString(ws_excel.Cells[point_row, 1].value) != "貨  號")
                {
                    point_row++;
                }
                int product_number_row = ++point_row;

                //抓取資料開頭列
                while (Convert.ToString(ws_excel.Cells[point_row, 1].value) != "no                     (欄位：9)")
                {
                    point_row++;
                }
                int tabel_head_row = point_row;
                int data_row = point_row + 2;
                if (Convert.ToString(ws_excel.Cells[data_row, 1].value) == null)
                    data_row++;

                //畫稿編號
                int art_no_col = 11;
                while (art_no_col < 50 && (Convert.ToString(ws_excel.Cells[product_number_row, art_no_col].value) == null || !Convert.ToString(ws_excel.Cells[product_number_row, art_no_col].value).Contains("-")))
                {
                    art_no_col++;
                }
                if (art_no_col > 30)
                    throw new IndexOutOfRangeException();
            
                int temp_art_no_col = art_no_col;
                //取得產品的價錢與成分
                //Queue -> price,item,ingredient,art_no
                Dictionary<string, List<string>> product_detail = new Dictionary<string, List<string>>();

                while (Convert.ToString(ws_excel.Cells[product_number_row, 1].value) != null)
                {
                    string product_name = Convert.ToString(ws_excel.Cells[product_number_row, 1].value);
                    product_name = product_name.Trim();
                    List<string> product = new List<string>();
                    string price = Convert.ToString(ws_excel.Cells[product_number_row, 2].value) + " " + Convert.ToString(ws_excel.Cells[product_number_row, 3].value);
                    //售價
                    product.Add(price);
                    //item
                    product.Add(Convert.ToString(ws_excel.Cells[product_number_row, 4].value)); 
                    //成分
                    double nylon_value = Convert.ToDouble(ws_excel.Cells[product_number_row, 7].value) * 100;
                    double spandex_value = Convert.ToDouble(ws_excel.Cells[product_number_row, 10].value) < 100 ? Convert.ToDouble(ws_excel.Cells[product_number_row, 10].value): Convert.ToDouble(ws_excel.Cells[product_number_row, 9].value);
                    spandex_value *= 100;
                    string ingredient = "vải " + Convert.ToString(ws_excel.Cells[product_number_row, 6].value) + " "
                        + nylon_value.ToString() + "% " 
                        + Convert.ToString(ws_excel.Cells[product_number_row, 8].value) + " "
                        + spandex_value.ToString() + "%";
                    product.Add(ingredient);

                    //art_no
                    //拿畫稿編號
                    int count = 1;
                    while (Convert.ToString(ws_excel.Cells[product_number_row, art_no_col].value) != null)
                    {
                        string produce = Convert.ToString(ws_excel.Cells[product_number_row, 1].value);
                        produce = produce.Trim();
                        if (!dic_art_no_col.ContainsKey(produce))
                            dic_art_no_col.Add(produce, art_no_col);
                        count++;
                        art_no_col++;
                    }
                    //把畫稿編號在檔案的位址存入product[3],[4]裡
                    string col_count_for_product = (count - 1).ToString();
                    product.Add(col_count_for_product);
                    product.Add(product_number_row.ToString());

                    product_detail.Add(product_name, product);

                    //把檔案sheet上方的序號和PO放入dic_po_no
                    dic_po_no.Add(product_name, (file_sheet_name, file_po_num));

                    product_number_row++;
                    art_no_col = temp_art_no_col; 
                }

                string product_check = "";
                string color_check = "";
                string art_no;
                //開始取資料
                int check_col = 7;
                int get_col_num = 0;
                while (Convert.ToString(ws_excel.Cells[data_row, check_col].value) != null)
                {
                    if (Convert.ToString(ws_excel.Cells[data_row, 1].value) == null)
                        data_row++;

                    if (Convert.ToString(ws_excel.Cells[data_row, check_col].value) == null)
                    {
                        if (Convert.ToString(ws_excel.Cells[data_row + 1, check_col].value) != null)
                            data_row += 1;
                        else
                            break;
                    }
                    //貨號
                    string product = Convert.ToString(ws_excel.Cells[data_row, 1].value);
                
                    //color
                    string color = Convert.ToString(ws_excel.Cells[data_row, 2].value);
                
                    //size
                    string size = Convert.ToString(ws_excel.Cells[data_row, 3].value);
                    //barcode
                    //string barcode = Convert.ToString(ws_excel.Cells[data_row, 4].value);
                    string barcode = product + "-" + color + "-" + size;

                    List<string> q_name = product_detail[product];

                    //同貨號且多個畫稿編號時，判斷顏色是否改變而更換畫稿編號
                    int product_art_no = Convert.ToInt32(q_name[3]);
                    if (product != product_check || get_col_num == 0)
                    { 
                        get_col_num = dic_art_no_col[product];
                        color_check = color;
                    }

                    if (product_art_no > 1)
                    {
                        if (color != color_check)
                            get_col_num++;
                    }
                    product_check = product;
                    color_check = color;

                    art_no = Convert.ToString(ws_excel.Cells[Convert.ToInt32(q_name[4]), get_col_num].value);
                    bonbon_Vps.Add(new Bonbon_vp_data(product
                        , color, size, barcode
                        ,q_name[0],q_name[1],q_name[2],art_no));
                    data_row++;
                }
            }
        }

        public override void VP_plate()
        {
            ws_excel = wb_excel.Worksheets[1];

            int row_end = 4;
            string art_no = ws_excel.Cells[row_end, 2].text;
            while (ws_excel.Cells[row_end, 4].text != "")
            {
                try
                {
                    string product_num = ws_excel.Cells[row_end, 4].text;

                    if (ws_excel.Cells[row_end, 2].value != null)
                        art_no = Convert.ToString(ws_excel.Cells[row_end, 2].value);

                    string color = Convert.ToString(ws_excel.Cells[row_end, 5].value);
                    string size = Convert.ToString(ws_excel.Cells[row_end, 6].value);
                    /*
                    string b_product = product_num;
                    while (b_product.Length < 9)
                        b_product += " ";
                    string b_color = color;
                    while (b_color.Length < 4)
                        b_color += " ";
                    string b_size = size;
                    while (b_size.Length < 4)
                        b_size += " ";
                    string barcode = b_product + b_color + b_size;
                    */
                    string barcode = product_num.Trim() + "-" + color + "-" + size;

                    string price = Convert.ToString(ws_excel.Cells[row_end, 7].value) + " " + Convert.ToString(ws_excel.Cells[row_end, 8].value);

                    string item = dic_Ingredient[art_no].Item1;
                    string ingredient = dic_Ingredient[art_no].Item2;

                    bonbon_Vps.Add(new Bonbon_vp_data(product_num, color, size, barcode, price, item, ingredient, art_no));

                    row_end++;
                    if (row_end == 86)
                    { 
                        //string bb = "";
                
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine(e.StackTrace.ToString());
                    Console.WriteLine(row_end.ToString());
                }
            }
        }
    }

    class Bonbon_Reference_file : Call_Excel
    {
        public List<Bonbon_vp_data> bonbon_Vps;
        public Dictionary<string, (string, string)> art_no_Ingredient = new Dictionary<string, (string, string)>();
        string mode;
        public Bonbon_Reference_file(List<Bonbon_vp_data> bonbon_Vps,string mode)
        {
            this.bonbon_Vps = bonbon_Vps;
            this.mode = mode;
        }
        public Bonbon_Reference_file(string mode)
        { 
            this.mode = mode;
        }

        public override void Load_File()
        {
            if (mode == "write")
            {
                string temp_art_no = "";
                string temp_product = "";
                ws_excel = wb_excel.Worksheets[1];
                int row_dock = 2;
                foreach (Bonbon_vp_data _Vp_Data in bonbon_Vps)
                {
                    if (_Vp_Data.Art_no != temp_art_no || _Vp_Data.Product_num != temp_product)
                    {
                        temp_product = _Vp_Data.Product_num;
                        temp_art_no = _Vp_Data.Art_no;
                        ws_excel.Cells[row_dock, 1].value = _Vp_Data.Art_no;
                        ws_excel.Cells[row_dock, 2].value = _Vp_Data.Product_num;
                        ws_excel.Cells[row_dock, 3].value = _Vp_Data.Color;
                        ws_excel.Cells[row_dock, 4].value = _Vp_Data.Item;
                        string temp = _Vp_Data.Ingredient;
                        string[] temp_arr = temp.Split(' ');

                        ws_excel.Cells[row_dock, 5].value = temp_arr[1] + " " + temp_arr[2];
                        ws_excel.Cells[row_dock, 6].value = temp_arr[3] + " " + temp_arr[4];
                        row_dock++;
                    }
                }
            }
            else if (mode == "read")
            {
                bonbon_Vps = new List<Bonbon_vp_data>();
                int row_end = 2;
                ws_excel = wb_excel.Worksheets[1];

                while (ws_excel.Cells[row_end, 1].text != "" || Convert.ToString(ws_excel.Cells[row_end, 1].text).Contains("-"))
                {
                    string art_no = Convert.ToString(ws_excel.Cells[row_end, 1].value);

                    string item = Convert.ToString(ws_excel.Cells[row_end, 4].value);

                    string Ingredient = "vải ";

                    Ingredient += Convert.ToString(ws_excel.Cells[row_end, 5].value) + " ";
                    Ingredient += Convert.ToString(ws_excel.Cells[row_end, 6].value);

                    art_no_Ingredient.Add(art_no, (item, Ingredient));
                    row_end++;
                }
            }
        }
    }


    class Bonbon_vp_data
    { 
        public string Product_num { get; set; }
        public string Color { get; set; }
        public string Size { get; set; }
        public string Barcode { get; set; }
        public string Prize { get; set; }
        public string Item { get; set; }
        public string Ingredient { get; set; }
        public string Art_no { get; set; }
        public Bonbon_vp_data(string Product,string color,string size,string barcode,string prize,string it,string ingredient,string an)
        { 
            this.Product_num = Product;
            this.Color = color;
            this.Size = size;
            this.Barcode = barcode;
            this.Prize = prize;
            this.Item = it;
            this.Ingredient = ingredient;
            this.Art_no = an;
        }
    }
}
