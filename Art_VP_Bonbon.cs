using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;


namespace Auto_Plan_Art_Plate
{
    internal class Art_VP_Bonbon
    {
        public List<Bonbon_vp_data> bonbon_Vps;
        public Dictionary<string, (string, string)> dic_po_no;
        string fn;
        SaveFileDialog saveFileDialog1;
        bool plus_year;
        public Art_VP_Bonbon(string file_name, List<Bonbon_vp_data> vP_Datas, Dictionary<string, (string, string)> dic_po_no,bool plus_year, SaveFileDialog saveFileDialog1)
        { 
            fn = file_name;
            bonbon_Vps = vP_Datas;
            this.dic_po_no = dic_po_no;
            this.saveFileDialog1 = saveFileDialog1;
            this.plus_year = plus_year;
        }
        public Art_VP_Bonbon(string file_name, List<Bonbon_vp_data> vP_Datas, SaveFileDialog saveFileDialog1)
        {
            fn = file_name;
            bonbon_Vps = vP_Datas;
            this.saveFileDialog1 = saveFileDialog1;
        }
        Dictionary<string, double> dic_switch_pos_x;
        Dictionary<string, double> dic_switch_pos_y;
        Illustrator.Layers mother_ls;
        public void Bonbon_check_vp()
        {
            Illustrator.Application app = new Illustrator.Application();
            Illustrator.Document doc_art = app.Open(fn);
            //取得母版圖稿
            mother_ls = doc_art.Layers;
            Illustrator.Layer mother_l = Getbyname_layer("母版", mother_ls);

            //拋出取得位置的thread
            Thread th_g_position_dic = new Thread(new ThreadStart(Get_position_dic_thread));
            th_g_position_dic.Start();

            //第一張圖的定位點
            double page_first_x = Getbyname_pathitem("1", Getbyname_layer("定位線", mother_ls).PathItems).Left;
            double page_first_y = Getbyname_pathitem("1", Getbyname_layer("定位線", mother_ls).PathItems).Top;
            //下張圖與起始定位點的距離
            double step_range_x = Get_step_range_x(Getbyname_layer("定位線", mother_ls));
            double step_range_y = Get_step_range_y(Getbyname_layer("定位線", mother_ls));
            //取得匯出樣板
            Illustrator.Document doc_vp_out = app.Open(System.Environment.CurrentDirectory + "\\bonbon匯出樣板.ai");
            Illustrator.Layers vp_layers = doc_vp_out.Layers;
            Illustrator.Layer vp_oring = vp_layers[1];

            //畫稿配圖
            Illustrator.GroupItem ori_pic_group = Getbyname_groupitem("畫稿配圖", mother_l.GroupItems);
            ori_pic_group.Duplicate(vp_oring);

            Illustrator.GroupItem pic_group = vp_oring.GroupItems[1];
            pic_group.Left = ori_pic_group.Left;
            pic_group.Top = ori_pic_group.Top;

            double pic_loc_x = ori_pic_group.Left;
            double pic_loc_y = ori_pic_group.Top;

            Illustrator.GroupItem original_g = Getbyname_groupitem("範本", mother_l.GroupItems);
            Illustrator.TextFrame textFrame = Getbyname_TextFrame("說明", mother_l.TextFrames);
            textFrame.Duplicate(vp_oring);

            Illustrator.TextFrame describe_tf = vp_oring.TextFrames[1];
            describe_tf.Left = textFrame.Left;
            describe_tf.Top = textFrame.Top;
            double describe_loc_x = textFrame.Left;
            double describe_loc_y = textFrame.Top;

            th_g_position_dic.Join();

            //背面
            double back_art_x = dic_switch_pos_x["背面"];
            double back_art_y = dic_switch_pos_y["背面"];
            Illustrator.GroupItem back_group_m = Getbyname_groupitem("背面", mother_l.GroupItems);
            back_group_m.Duplicate(vp_oring);
            Illustrator.GroupItem back_g = vp_oring.GroupItems[1];
            back_g.Left = back_art_x;
            back_g.Top = back_art_y;
            DateTime today = DateTime.Today;
            Illustrator.TextFrame year_t = Getbyname_TextFrame("year", back_g.TextFrames);
            string year_ = today.Year.ToString();
            if (plus_year)
                year_ = (today.Year + 1).ToString();
            year_t.Contents += year_;

            //開始匯出
            //紀錄已匯出的畫稿編號跟顏色與當前比較，檢查是否換頁
            string temp_art_no = bonbon_Vps[0].Art_no;
            string temp_color = bonbon_Vps[0].Color;
            string temp_file_sheet_name = dic_po_no[bonbon_Vps[0].Product_num].Item1;
            int step_on_page = 1;

            describe_tf.Contents = Set_describe_str(describe_tf.Contents, bonbon_Vps[0]);

            int data_count = 0;
            int page_count = 1;
            foreach (Bonbon_vp_data bb_Vp in bonbon_Vps)
            {
                if (dic_po_no[bb_Vp.Product_num].Item1 != temp_file_sheet_name)
                {
                    double get_left = Getbyname_pathitem("1", Getbyname_layer("定位線", mother_ls).PathItems).Left;
                    page_first_x = get_left;
                    page_first_y += step_range_y;
                    temp_art_no = bb_Vp.Art_no;
                    temp_color = bb_Vp.Color;
                    //畫稿頁面的其他圖像
                    pic_loc_x = ori_pic_group.Left;
                    ori_pic_group.Duplicate(vp_oring);
                    pic_group = vp_oring.GroupItems[1];

                    pic_group.Left = pic_loc_x;

                    pic_loc_y += step_range_y;
                    pic_group.Top = pic_loc_y;
                    Illustrator.TextFrame page = Getbyname_TextFrame("page", pic_group.TextFrames);
                    page_count = 1;
                    page.Contents = "P" + 1;
                    
                    temp_file_sheet_name = dic_po_no[bb_Vp.Product_num].Item1;
                    //文字窗格的文字處理
                    describe_loc_x = textFrame.Left;
                    textFrame.Duplicate(vp_oring);
                    describe_tf = vp_oring.TextFrames[1];
                    describe_tf.Left = describe_loc_x;

                    describe_tf.Contents = Set_describe_str(describe_tf.Contents,bb_Vp);

                    describe_loc_y += step_range_y;
                    describe_tf.Top = describe_loc_y;

                    //背面
                    back_art_y += step_range_y;
                    back_art_x = dic_switch_pos_x["背面"];
                    back_group_m.Duplicate(vp_oring);
                    back_g = vp_oring.GroupItems[1];
                    back_g.Left = back_art_x;
                    back_g.Top = back_art_y;

                    year_t = Getbyname_TextFrame("year", back_g.TextFrames);
                    year_t.Contents += year_;

                    step_on_page = 1;
                }
                else
                { 
                    if (bb_Vp.Art_no != temp_art_no || bb_Vp.Color != temp_color || step_on_page > 7)
                    {
                        page_first_x += step_range_x;
                        temp_art_no = bb_Vp.Art_no;
                        temp_color = bb_Vp.Color;

                        //圖稿
                        pic_loc_x += step_range_x;
                        ori_pic_group.Duplicate(vp_oring);
                        pic_group = vp_oring.GroupItems[1];
                        pic_group.Left = pic_loc_x;
                        pic_group.Top = pic_loc_y;
                        Illustrator.TextFrame page = Getbyname_TextFrame("page", pic_group.TextFrames);
                        page_count++;
                        page.Contents = "P" + page_count.ToString();
                        
                        //說明
                        describe_loc_x += step_range_x;
                        textFrame.Duplicate(vp_oring);
                        describe_tf = vp_oring.TextFrames[1];
                        describe_tf.Left = describe_loc_x;
                        describe_tf.Contents = Set_describe_str(describe_tf.Contents,bb_Vp);
                        describe_tf.Top = describe_loc_y;

                        //背面
                        back_art_x += step_range_x;
                        back_group_m.Duplicate(vp_oring);
                        back_g = vp_oring.GroupItems[1];
                        back_g.Left = back_art_x;
                        back_g.Top = back_art_y;
                        year_t = Getbyname_TextFrame("year", back_g.TextFrames);
                        year_t.Contents += year_;

                        step_on_page = 1;
                    }
                }

                double print_x = page_first_x + dic_switch_pos_x[step_on_page.ToString()];
                double print_y = page_first_y + dic_switch_pos_y[step_on_page.ToString()];
                original_g.Duplicate(vp_oring);
                //畫稿的全部圖形
                Illustrator.GroupItem print_group = vp_oring.GroupItems[1];
                print_group.Left = print_x;
                print_group.Top = print_y;
                print_group.Name = bb_Vp.Barcode;

                Illustrator.GroupItem data_group = Getbyname_groupitem("資料", print_group.GroupItems);

                //產生barcode，取得定位線條
                Illustrator.PathItem position_path = Getbyname_pathitem("pos", data_group.PathItems);
                Art_Barcode art_b = new Art_Barcode(bb_Vp.Barcode, 31, 8.7, data_group, position_path);

                Thread th_set_barcode = new Thread(new ThreadStart(art_b.Print_in_art_barcode));
                th_set_barcode.Start();

                Illustrator.TextFrames data_tfs = data_group.TextFrames;

                if (step_on_page != 1)
                    Getbyname_pathitem("ellipse",print_group.PathItems).Delete();

                Getbyname_TextFrame("Item", data_tfs).Contents = bb_Vp.Item;
                Getbyname_TextFrame("Ingredient", data_tfs).Contents = bb_Vp.Ingredient;
                Getbyname_TextFrame("Product", data_tfs).Contents = bb_Vp.Product_num;
                Getbyname_TextFrame("Size", data_tfs).Contents = bb_Vp.Size;
                Getbyname_TextFrame("Color", data_tfs).Contents = bb_Vp.Color;
                Getbyname_TextFrame("Price", data_tfs).Contents = bb_Vp.Prize;
                Getbyname_TextFrame("Barcode", data_tfs).Contents = bb_Vp.Barcode;

                th_set_barcode.Join();
                
                step_on_page++;
                data_count++;
            }

            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                Illustrator.IllustratorSaveOptions saveOptions = new Illustrator.IllustratorSaveOptions();
                saveOptions.Compatibility = Illustrator.AiCompatibility.aiIllustrator17;
                saveOptions.FlattenOutput = Illustrator.AiOutputFlattening.aiPreserveAppearance;
                doc_vp_out.SaveAs(saveFileDialog1.FileName, saveOptions);
                doc_vp_out.Close();

                MessageBox.Show("輸出資料OK", "提醒");
            }
            saveFileDialog1.FileName = "";

            doc_art.Close();
        }

        public void Bonbon_plate_vp()
        {
            Illustrator.Application app = new Illustrator.Application();
            Illustrator.Document doc_art = app.Open(fn);
            //取得母版圖稿
            mother_ls = doc_art.Layers;
            Illustrator.Layer mother_l = Getbyname_layer("母版", mother_ls);

            double page_first_x = 5 * 2.835;
            double page_first_y = -10 * 2.835;

            Illustrator.GroupItem ori_group = Getbyname_groupitem("備拼", mother_l.GroupItems);

            Illustrator.Document doc_vp_out = app.Open(System.Environment.CurrentDirectory + "\\母版樣板.ai");
            Illustrator.Layers vp_layers = doc_vp_out.Layers;

            Illustrator.Layer vp_oring = vp_layers[1];


            int data_count = 0;
            foreach (Bonbon_vp_data _Vp_Data in bonbon_Vps)
            {
                if (data_count % 15 == 0 & data_count != 0)
                { 
                    page_first_y -= 10 * 2.835 + ori_group.Height;
                    page_first_x = 5 * 2.835;
                }
                else if (data_count > 0)
                    page_first_x += 5 * 2.835 + ori_group.Width;

                data_count++;

                vp_layers.Add();
                Illustrator.Layer vp_l = vp_layers[1];
                vp_l.Name = "圖層 " + data_count.ToString();

                ori_group.Duplicate(vp_l);
                Illustrator.GroupItem dup_group = vp_l.GroupItems[1];
                dup_group.Left = page_first_x;
                dup_group.Top = page_first_y;
                //條碼
                Illustrator.PathItem position_path = Getbyname_pathitem("pos", Getbyname_groupitem("資料",dup_group.GroupItems).PathItems);
                Art_Barcode art_b = new Art_Barcode(_Vp_Data.Barcode, 31, 8.7, Getbyname_groupitem("資料", dup_group.GroupItems), position_path);
                Thread th_set_barcode = new Thread(new ThreadStart(art_b.Print_in_art_barcode));
                th_set_barcode.Start();

                //序號
                Illustrator.TextFrame no_item = Getbyname_TextFrame("No",dup_group.TextFrames);
                no_item.Contents = data_count.ToString();
                //資料
                Illustrator.GroupItem data_group = Getbyname_groupitem("資料", dup_group.GroupItems);

                Illustrator.TextFrames data_tfs = data_group.TextFrames;
                Getbyname_TextFrame("Item", data_tfs).Contents = _Vp_Data.Item;
                Getbyname_TextFrame("Ingredient", data_tfs).Contents = _Vp_Data.Ingredient;
                Getbyname_TextFrame("Product", data_tfs).Contents = _Vp_Data.Product_num;
                Getbyname_TextFrame("Size", data_tfs).Contents = _Vp_Data.Size;
                Getbyname_TextFrame("Color", data_tfs).Contents = _Vp_Data.Color;
                Getbyname_TextFrame("Price", data_tfs).Contents = _Vp_Data.Prize;
                Getbyname_TextFrame("Barcode", data_tfs).Contents = _Vp_Data.Barcode;

                th_set_barcode.Join();
            }
            vp_oring.Delete();
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                Illustrator.IllustratorSaveOptions saveOptions = new Illustrator.IllustratorSaveOptions();
                saveOptions.Compatibility = Illustrator.AiCompatibility.aiIllustrator17;
                saveOptions.FlattenOutput = Illustrator.AiOutputFlattening.aiPreserveAppearance;
                doc_vp_out.SaveAs(saveFileDialog1.FileName, saveOptions);
                doc_vp_out.Close();


                MessageBox.Show("輸出資料OK", "提醒");
            }
            saveFileDialog1.FileName = "";

            doc_art.Close();
        }
        string Set_describe_str(string describle_str, Bonbon_vp_data bonbon_Vp)
        {
            string temp_str;
            string[] part_describle = describle_str.Split('[');
            temp_str = part_describle[0] + bonbon_Vp.Art_no;
            DateTime today = DateTime.Today;
            string today_str = today.Year.ToString();
            today_str += (today.Month.ToString().Length < 2) ? "/0" + today.Month.ToString() : "/" + today.Month.ToString();
            today_str += (today.Day.ToString().Length < 2) ? "/0" + today.Day.ToString() : "/" + today.Day.ToString();

            temp_str += part_describle[1] + today_str;
            //temp_str += part_describle[2] + "序號";
            temp_str += part_describle[2] + dic_po_no[bonbon_Vp.Product_num].Item1;
            //temp_str += part_describle[3] + "po";
            temp_str += part_describle[3] + dic_po_no[bonbon_Vp.Product_num].Item2;
            temp_str += part_describle[4] + bonbon_Vp.Product_num;
            temp_str += part_describle[5] + bonbon_Vp.Color;
            return temp_str;
        }

        void Get_position_dic_thread()
        {
            dic_switch_pos_x = Get_positon_dic(Getbyname_layer("定位線", mother_ls), "x");
            dic_switch_pos_y = Get_positon_dic(Getbyname_layer("定位線", mother_ls), "y");
        }
        public double Get_step_range_x(Illustrator.Layer layer)
        {
            Illustrator.PathItem frist_page = Getbyname_pathitem("1", layer.PathItems);
            Illustrator.PathItem secord_page = Getbyname_pathitem("第二頁", layer.PathItems);

            double x = secord_page.Left - frist_page.Left;
            return x;
        }
        public double Get_step_range_y(Illustrator.Layer layer)
        {
            Illustrator.PathItem frist_row = Getbyname_pathitem("1", layer.PathItems);
            Illustrator.PathItem secord_row = Getbyname_pathitem("第二列", layer.PathItems);

            double y = secord_row.Top - frist_row.Top;
            return y;
        }

        public Dictionary<string, double> Get_positon_dic(Illustrator.Layer layer,string mode)
        {
            Dictionary<string,double> position_dic = new Dictionary<string,double>();
            Illustrator.PathItem loc_path = Getbyname_pathitem("1", layer.PathItems);
            for (int i = 1; i <= 7; i++)
            {
                Illustrator.PathItem pos_path = Getbyname_pathitem(i.ToString(), layer.PathItems);
                
                double pos_value = 0;
                if (mode == "x")
                    pos_value = pos_path.Left - loc_path.Left;
                else if (mode == "y")
                    pos_value = pos_path.Top - loc_path.Top;
                position_dic.Add(i.ToString(), pos_value);
            }
            Illustrator.PathItem back_path = Getbyname_pathitem("背面", layer.PathItems);
            if (mode == "x")
                position_dic.Add("背面", back_path.Left);
            else if (mode == "y")
                position_dic.Add("背面", back_path.Top);
            return position_dic;
        }

        public Illustrator.Layer Getbyname_layer(string key, Illustrator.Layers i_layers)
        {
            foreach (Illustrator.Layer layer in i_layers)
            {
                if (layer.Name == key)
                    return layer;
            }
            return null;
        }
        public Illustrator.GroupItem Getbyname_groupitem(string key, Illustrator.GroupItems i_groups)
        {
            foreach (Illustrator.GroupItem groupItem in i_groups)
            {
                if (groupItem.Name == key)
                    return groupItem;
            }
            return null;
        }
        public Illustrator.TextFrame Getbyname_TextFrame(string key, Illustrator.TextFrames i_tf)
        {
            foreach (Illustrator.TextFrame textFrame in i_tf)
            {
                if (textFrame.Name == key)
                    return textFrame;
            }
            return null;
        }
        public Illustrator.PathItem Getbyname_pathitem(string key, Illustrator.PathItems i_pis)
        {
            foreach (Illustrator.PathItem pathItem in i_pis)
            {
                if (pathItem.Name == key)
                    return pathItem;
            }
            return null;
        }
    }
}
