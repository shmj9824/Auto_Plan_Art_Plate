using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Auto_Plan_Art_Plate
{
    internal class Art_plate
    {
        public bool check_doc_art(string file_art, int art_unit_sum)
        {
            bool check_message = false;
            Illustrator.Application app = new Illustrator.Application();

            //畫稿來源
            Illustrator.Document doc_art = app.Open(file_art);

            Illustrator.Layers lrs_plan = doc_art.Layers;

            List<string> layers_name = new List<string>();
            for (int i = 1; i <= lrs_plan.Count; i++)
            {
                layers_name.Add(lrs_plan[i].Name);
            }

            if (layers_name.Count == art_unit_sum)
                check_message = true;
            doc_art.Close();
            app.Quit();
            return check_message;
        }

        public void Tyvek_plate(ref int rotate_num, ref int max_col, ref int max_row, string file_art, string file_plate, string plate_explan, ref List<Dictionary<int, int>> Plate_data_list, ref SaveFileDialog saveFileDialog1)
        {
            //rotate_num 畫稿旋轉角度，file_art 原稿檔案，file_plate 模板檔案，plate_explan 模板說明，Plate_data_list 模板內容(模板數量、編號、數量)

            Illustrator.Application app = new Illustrator.Application();
            //畫稿來源
            Illustrator.Document doc_art = app.Open(file_art);
            Illustrator.Layers lrs_plan = doc_art.Layers;
            //模板範本
            Illustrator.Document doc_output = app.Open(file_plate);
            Illustrator.Layers layers_output = doc_output.Layers;
            //模板說明
            Illustrator.TextFrame art_plate_name = getByname_layer("版1", layers_output).TextFrames[1];
            art_plate_name.Contents = plate_explan;
            art_plate_name.Name = "模板說明";

            foreach (Dictionary<int, int> obj_item_num in Plate_data_list)
            {
                //多版以圖層區分
                Illustrator.Layer lay_output = layers_output.Add();

                //取基準線的定位點
                Illustrator.PathItem position_path = getByname_layer("版1", layers_output).PathItems[1];
                double location_p_x = position_path.Left;
                double location_p_y = position_path.Top;
                //畫稿的長度(y)和寬度(x)
                double obj_height = 0;
                double obj_width = 0;
                Illustrator.GroupItem check = lrs_plan[1].GroupItems[1];
                obj_height = check.Height;
                obj_width = check.Width;

                int row_obj_num = 0;

                int max_paint_x_num = max_col;
                int max_paint_y_num = max_row;

                int obj_x_num = 1;
                int obj_y_num = 0;

                foreach (KeyValuePair<int, int> k_pair in obj_item_num)
                {
                    Illustrator.GroupItem groupItem = lrs_plan["圖層 " + k_pair.Key.ToString()].GroupItems[1];

                    for (int i = 1; i <= k_pair.Value; i++)
                    {
                        groupItem.Duplicate(lay_output);
                        //複製和新增物件都是從1開始
                        Illustrator.GroupItem dup_group = lay_output.GroupItems[1];

                        dup_group.Rotate(rotate_num);
                        if (rotate_num == 90)
                        {
                            dup_group.Top = location_p_y - obj_width * obj_y_num;
                            dup_group.Left = location_p_x;
                            dup_group = null;

                            obj_y_num++;
                            if (obj_y_num == max_paint_y_num)
                            {
                                obj_y_num = 0;
                                obj_x_num++;

                                location_p_x += obj_height;
                            }
                        }
                        else if (rotate_num == 0)
                        {
                            dup_group.Top = location_p_y - obj_height * obj_y_num;
                            dup_group.Left = location_p_x;

                            dup_group = null;
                            obj_y_num++;
                            if (obj_y_num == max_paint_y_num)
                            {
                                obj_y_num = 0;
                                obj_x_num++;

                                location_p_x += obj_width;
                            }
                        }
                    }
                    row_obj_num += k_pair.Value;
                }
            }

            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                Illustrator.IllustratorSaveOptions saveOptions = new Illustrator.IllustratorSaveOptions();
                saveOptions.Compatibility = Illustrator.AiCompatibility.aiIllustrator17;
                saveOptions.FlattenOutput = Illustrator.AiOutputFlattening.aiPreserveAppearance;
                doc_output.SaveAs(saveFileDialog1.FileName, saveOptions);
                doc_output.Close();

                MessageBox.Show("輸出資料OK", "提醒");
            }
            saveFileDialog1.FileName = "";

            doc_art.Close();
            //app.Quit();
        }

        public void iliea_plate(ref int rotate_num, ref int max_col, ref int max_row, string file_art, string file_plate, string plate_explan, ref List<Dictionary<int, int>> Plate_data_list, ref SaveFileDialog saveFileDialog1)
        {
            //需要由左而右排列
            Illustrator.Application app = new Illustrator.Application();
            //畫稿來源
            Illustrator.Document doc_art = app.Open(file_art);
            Illustrator.Layers lrs_plan = doc_art.Layers;
            //模板範本
            Illustrator.Document doc_output = app.Open(file_plate);
            Illustrator.Layers layers_output = doc_output.Layers;
            //模板說明
            Illustrator.TextFrame art_plate_name = getByname_layer("版1", layers_output).TextFrames[1];
            art_plate_name.Contents = plate_explan;
            art_plate_name.Name = "模板說明";

            foreach (Dictionary<int, int> obj_item_num in Plate_data_list)
            {
                //多版以圖層區分
                Illustrator.Layer lay_output = layers_output.Add();

                //取基準線的定位點
                Illustrator.PathItem position_path = getByname_layer("版1", layers_output).PathItems[1];
                double location_p_x = position_path.Left;
                double location_p_y = position_path.Top;
                //畫稿的長度(y)和寬度(x)
                double obj_height = 0;
                double obj_width = 0;
                Illustrator.GroupItem check = lrs_plan[1].GroupItems[1];
                obj_height = check.Height;
                obj_width = check.Width;

                //int row_obj_num = 0;

                int max_paint_x_num = max_col;
                int max_paint_y_num = max_row;
                //物件的初始值
                int obj_x_num = 0;
                int obj_y_num = 1;
                //空白間隔
                double interval_x = 4.99 * 2.835;
                double interval_y = 5.88 * 2.835;
                //間隔次數
                int step_interval_x = 0;

                Illustrator.GroupItem gi_no = lay_output.GroupItems.Add();
                gi_no.Name = "product_no";

                foreach (KeyValuePair<int, int> k_pair in obj_item_num)
                {
                    Illustrator.GroupItem groupItem = lrs_plan["圖層 " + k_pair.Key.ToString()].GroupItems[1];
                    
                    for (int i = 1; i <= k_pair.Value; i++)
                    {
                        groupItem.Duplicate(lay_output);
                        //複製和新增物件都是從1開始
                        Illustrator.GroupItem dup_group = lay_output.GroupItems[1];
                        Illustrator.TextFrame no_tf = Getbyname_TextFrame("No", dup_group.TextFrames);
                        dup_group.Name = no_tf.Contents;
                        dup_group.Rotate(rotate_num);
                        if (rotate_num == 0)
                        {
                            dup_group.Top = location_p_y;
                            //每幾格加一次
                            if (obj_x_num % 8 == 0 && obj_x_num > 0)
                                step_interval_x++;
                            dup_group.Left = location_p_x + obj_width * obj_x_num + interval_x * step_interval_x;
                            dup_group = null;

                            obj_x_num++;
                            if (obj_x_num == max_paint_x_num)
                            {
                                obj_x_num = 0;
                                obj_y_num++;

                                location_p_y -= obj_height;
                                if (obj_y_num > 0)
                                    location_p_y -= interval_y;
                                step_interval_x = 0;
                            }
                        }
                        else if (rotate_num == 90)
                        {
                            dup_group.Top = location_p_y;
                            //每幾格加一次
                            if (obj_x_num % 8 == 0 && obj_x_num > 0)
                                step_interval_x++;
                            dup_group.Left = location_p_x + obj_height * obj_x_num + interval_x * step_interval_x;
                            dup_group = null;

                            obj_x_num++;
                            if (obj_x_num == max_paint_x_num)
                            {
                                obj_x_num = 0;
                                obj_y_num++;

                                location_p_y -= obj_width;
                                if (obj_y_num > 0)
                                    location_p_y -= interval_y;
                                step_interval_x = 0;
                            }
                        }
                        no_tf.Move(gi_no, Illustrator.AiElementPlacement.aiPlaceAtEnd);
                    }   
                }
            }
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                Illustrator.IllustratorSaveOptions saveOptions = new Illustrator.IllustratorSaveOptions();
                saveOptions.Compatibility = Illustrator.AiCompatibility.aiIllustrator17;
                saveOptions.FlattenOutput = Illustrator.AiOutputFlattening.aiPreserveAppearance;
                doc_output.SaveAs(saveFileDialog1.FileName, saveOptions);
                doc_output.Close();

                MessageBox.Show("輸出資料OK", "提醒");
            }
            saveFileDialog1.FileName = "";

            doc_art.Close();
            //app.Quit();
        }

        public void bonbon_plate(ref int rotate_num, ref int max_col, ref int max_row, string file_art, string file_plate, string plate_explan, ref List<Dictionary<int, int>> Plate_data_list, ref SaveFileDialog saveFileDialog1)
        {
            //需要由左而右排列
            Illustrator.Application app = new Illustrator.Application();
            //畫稿來源
            Illustrator.Document doc_art = app.Open(file_art);
            Illustrator.Layers lrs_plan = doc_art.Layers;
            //模板範本
            Illustrator.Document doc_output = app.Open(file_plate);
            Illustrator.Layers layers_output = doc_output.Layers;
            //模板說明
            Illustrator.TextFrame art_plate_name = Getbyname_TextFrame("說明", getByname_layer("版1", layers_output).TextFrames);
            art_plate_name.Contents = plate_explan;
            art_plate_name.Name = "模板說明";

            foreach (Dictionary<int, int> obj_item_num in Plate_data_list)
            {
                //多版以圖層區分
                Illustrator.Layer lay_output = layers_output.Add();

                //取基準線的定位點
                Illustrator.PathItem position_path = Getbyname_pathitem("1", getByname_layer("版1", layers_output).PathItems);
                double location_p_x = position_path.Left;
                double location_p_y = position_path.Top;
                //畫稿的長度(y)和寬度(x)
                double obj_height = 0;
                double obj_width = 0;
                Illustrator.GroupItem check = lrs_plan[1].GroupItems[1];
                obj_height = check.Height;
                obj_width = check.Width;

                //int row_obj_num = 0;

                int max_paint_x_num = max_col;
                int max_paint_y_num = max_row;
                //物件的初始值
                int obj_x_num = 0;
                int obj_y_num = 1;
                
                Illustrator.PathItem secord = Getbyname_pathitem("第二列", getByname_layer("版1", layers_output).PathItems);
                double interval_y = secord.Top - position_path.Top;
                
                Illustrator.GroupItem gi_no = lay_output.GroupItems.Add();
                gi_no.Name = "product_no";

                foreach (KeyValuePair<int, int> k_pair in obj_item_num)
                {
                    Illustrator.GroupItem groupItem = lrs_plan["圖層 " + k_pair.Key.ToString()].GroupItems[1];

                    for (int i = 1; i <= k_pair.Value; i++)
                    {
                        groupItem.Duplicate(lay_output);
                        //複製和新增物件都是從1開始
                        Illustrator.GroupItem dup_group = lay_output.GroupItems[1];
                        Illustrator.TextFrame no_tf = Getbyname_TextFrame("No", dup_group.TextFrames);
                        dup_group.Name = no_tf.Contents;
                        dup_group.Rotate(rotate_num);
                        if (rotate_num == 0)
                        {
                            dup_group.Top = location_p_y;
                            dup_group.Left = location_p_x + obj_width * obj_x_num;
                            dup_group = null;

                            obj_x_num++;
                            if (obj_x_num == max_paint_x_num)
                            {
                                obj_x_num = 0;
                                obj_y_num++;

                                //location_p_y -= obj_height;
                                if (obj_y_num > 0)
                                    location_p_y += interval_y;
                            }
                        }
                        else 
                        {
                            
                        }
                        no_tf.Move(gi_no, Illustrator.AiElementPlacement.aiPlaceAtEnd);
                    }
                }
            }
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                Illustrator.IllustratorSaveOptions saveOptions = new Illustrator.IllustratorSaveOptions();
                saveOptions.Compatibility = Illustrator.AiCompatibility.aiIllustrator17;
                saveOptions.FlattenOutput = Illustrator.AiOutputFlattening.aiPreserveAppearance;
                doc_output.SaveAs(saveFileDialog1.FileName, saveOptions);
                doc_output.Close();

                MessageBox.Show("輸出資料OK", "提醒");
            }
            saveFileDialog1.FileName = "";

            doc_art.Close();
        }

        public void Basic_bonbon_plate(ref int rotate_num, ref int max_col, ref int max_row, string file_art, string file_plate, string plate_explan, ref List<Dictionary<int, int>> Plate_data_list,ref DateTime tod, ref SaveFileDialog saveFileDialog1)
        {
            //需要由左而右排列
            Illustrator.Application app = new Illustrator.Application();
            //畫稿來源
            Illustrator.Document doc_art = app.Open(file_art);
            Illustrator.Layers lrs_plan = doc_art.Layers;
            //模板範本
            Illustrator.Document doc_output = app.Open(file_plate);
            Illustrator.Layers layers_output = doc_output.Layers;
            //模板說明
            Illustrator.TextFrame art_plate_name = Getbyname_TextFrame("說明", getByname_layer("版1", layers_output).TextFrames);
            art_plate_name.Contents = plate_explan;
            art_plate_name.Name = "模板說明";
            //出版日期
            string plate_day = (tod.Year - 1911).ToString() + "." + tod.Month + "." + tod.Day + "出版";
            Illustrator.TextFrame art_plate_day = Getbyname_TextFrame("出版", getByname_layer("版1", layers_output).TextFrames);
            art_plate_day.Contents = plate_day;
            
            foreach (Dictionary<int, int> obj_item_num in Plate_data_list)
            {
                //多版以圖層區分
                Illustrator.Layer lay_output = layers_output.Add();

                //取基準線的定位點
                Illustrator.PathItem position_path = Getbyname_pathitem("1", getByname_layer("版1", layers_output).PathItems);
                double location_p_x = position_path.Left;
                double location_p_y = position_path.Top;
                //畫稿的長度(y)和寬度(x)
                double obj_height = 0;
                double obj_width = 0;
                Illustrator.GroupItem check = lrs_plan[1].GroupItems[1];
                obj_height = check.Height;
                obj_width = check.Width;

                //int row_obj_num = 0;

                int max_paint_x_num = max_col;
                int max_paint_y_num = max_row;
                //物件的初始值
                int obj_x_num = 0;
                int obj_y_num = 1;

                Illustrator.PathItem secord = Getbyname_pathitem("第二列", getByname_layer("版1", layers_output).PathItems);
                double interval_y = secord.Top - position_path.Top;

                Illustrator.GroupItem gi_no = lay_output.GroupItems.Add();
                gi_no.Name = "product_no";

                foreach (KeyValuePair<int, int> k_pair in obj_item_num)
                {
                    Illustrator.GroupItem groupItem = lrs_plan["圖層 " + k_pair.Key.ToString()].GroupItems[1];

                    for (int i = 1; i <= k_pair.Value; i++)
                    {
                        groupItem.Duplicate(lay_output);
                        //複製和新增物件都是從1開始
                        Illustrator.GroupItem dup_group = lay_output.GroupItems[1];
                        Illustrator.TextFrame no_tf = Getbyname_TextFrame("No", dup_group.TextFrames);
                        dup_group.Name = no_tf.Contents;
                        dup_group.Rotate(rotate_num);
                        if (rotate_num == 0)
                        {
                            dup_group.Top = location_p_y;
                            dup_group.Left = location_p_x + obj_width * obj_x_num;
                            dup_group = null;

                            obj_x_num++;
                            if (obj_x_num == max_paint_x_num)
                            {
                                obj_x_num = 0;
                                obj_y_num++;

                                //location_p_y -= obj_height;
                                if (obj_y_num > 0)
                                    location_p_y += interval_y;
                            }
                        }
                        else
                        {

                        }
                        no_tf.Move(gi_no, Illustrator.AiElementPlacement.aiPlaceAtEnd);
                    }
                }
            }
            
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                Illustrator.IllustratorSaveOptions saveOptions = new Illustrator.IllustratorSaveOptions();
                saveOptions.Compatibility = Illustrator.AiCompatibility.aiIllustrator17;
                saveOptions.FlattenOutput = Illustrator.AiOutputFlattening.aiPreserveAppearance;
                doc_output.SaveAs(saveFileDialog1.FileName, saveOptions);
                doc_output.Close();

                MessageBox.Show("輸出資料OK", "提醒");
            }
            saveFileDialog1.FileName = "";

            doc_art.Close();
        }

        public Illustrator.Layer getByname_layer(string name,Illustrator.Layers layers)
        {
            foreach (Illustrator.Layer la in layers)
            { 
                if (la.Name == name)
                    return la;
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
