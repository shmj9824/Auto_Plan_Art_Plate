using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;

using BarcodeStandard;
//using BarcodeLib;
using SkiaSharp;
using Type = BarcodeStandard.Type;

namespace Auto_Plan_Art_Plate
{
    internal class Art_VP_File
    {
        public List<VP_data> vp_Datas = new List<VP_data>();
        string vp_file;
        
        SaveFileDialog saveFileDialog1;

        public Art_VP_File(string file_name,List<VP_data> vP_Datas, SaveFileDialog saveFileDialog1)
        {
            vp_file = file_name;
            vp_Datas = vP_Datas;
            this.saveFileDialog1 = saveFileDialog1;
        }
        public void Ilina_vp()
        {
            Illustrator.Application app = new Illustrator.Application();
            Illustrator.Document doc_art = app.Open(vp_file);
            //取得母版圖稿
            Illustrator.Layers lrs_plan = doc_art.Layers;
            Illustrator.Layer lr_plan = Getbyname_layer("母版", lrs_plan);

            string group_name = "核稿";
            if (vp_Datas[0].no == 1)
                group_name = "拼模";
            //Illustrator.GroupItem Inheritance_item = lr_plan.GroupItems[1];
            Illustrator.GroupItem Inheritance_item = Getbyname_groupitem(group_name,lr_plan.GroupItems);
            //取得母版圖層
            Illustrator.Document doc_vp_out = app.Open(System.Environment.CurrentDirectory + "\\母版樣板.ai");
            Illustrator.Layers vp_layers = doc_vp_out.Layers;

            Illustrator.Layer vp_oring = vp_layers[1];

            //取定位點
            double location_p_x = 21 * 2.835 -1.55;
            double location_p_y = -19 * 2.835;
            int print_num = 1;

            foreach (VP_data vp in vp_Datas)
            {
                vp_layers.Add();
                Illustrator.Layer vp_l = vp_layers[1];
                vp_l.Name = "圖層 " + print_num.ToString();

                Inheritance_item.Duplicate(vp_l);
                Illustrator.GroupItem dup_group = vp_l.GroupItems[1];
                dup_group.Top = location_p_y;
                dup_group.Left = location_p_x;
                //產生barcode
                //取得定位線條
                Illustrator.PathItem position_path = Getbyname_pathitem("pos", dup_group.PathItems);
                Art_Barcode art_b = new Art_Barcode(vp.item5, 28, 12, dup_group, position_path);
                Thread th_set_barcode = new Thread(new ThreadStart(art_b.Print_in_art_barcode));
                th_set_barcode.Start();

                location_p_x += dup_group.Width + 12 * 2.835 -1.3;
                if (print_num % 12 == 0)
                {
                    location_p_x += (-12 + 21 + 5 + 21 + 9) * 2.835 -1.3;
                    location_p_y = -19 * 2.835;
                }
                else if (print_num % 4 == 0)
                {
                    location_p_y -= dup_group.Height + 7.5 * 2.835;
                    location_p_x -= dup_group.Width * 4 + 12 * 4 * 2.835;
                }
                if (vp.no != 0)
                    Getbyname_TextFrame("No", dup_group.TextFrames).Contents = vp.no.ToString();

                Getbyname_TextFrame("style_1", dup_group.TextFrames).Contents = vp.item1;
                Getbyname_TextFrame("style_2", dup_group.TextFrames).Contents = vp.item1;

                Getbyname_TextFrame("color_1", dup_group.TextFrames).Contents = vp.item2;
                Getbyname_TextFrame("color_2", dup_group.TextFrames).Contents = vp.item2;

                Getbyname_TextFrame("size_1", dup_group.TextFrames).Contents = vp.item3;
                Getbyname_TextFrame("size_2", dup_group.TextFrames).Contents = vp.item3;

                Getbyname_TextFrame("price_1", dup_group.TextFrames).Contents = vp.item4;
                Getbyname_TextFrame("price_2", dup_group.TextFrames).Contents = vp.item4;

                th_set_barcode.Join();
                
                print_num++;
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
            //app.Quit();
        }
        
        public Illustrator.Layer Getbyname_layer(string key,Illustrator.Layers i_layers)
        {
            foreach (Illustrator.Layer layer in i_layers)
            { 
                if(layer.Name == key)
                    return layer;
            }
            return null;
        }
        public Illustrator.GroupItem Getbyname_groupitem(string key,Illustrator.GroupItems i_groups)
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

        public void Print_barcode_inai(string print_data,Illustrator.GroupItem gi,double x,double y)
        {
            double point_x = x;
            double point_y = y;
            //int point_width = 1;
        }

        public void Test_barcode_ai()
        {
            Illustrator.Application app = new Illustrator.Application();
            Illustrator.Document doc_art = app.Open("D:\\程式文案\\016畫稿自動拼模\\條碼\\物件抓取.ai");

            //取得母版圖稿
            Illustrator.Layers lrs_plan = doc_art.Layers;
            Illustrator.Layer lr_plan = Getbyname_layer("test", lrs_plan);
            
            //Illustrator.PluginItem test_plugini = lr_plan.PluginItems[1];
            
            //取得母版圖層
            Illustrator.Document doc_vp_out = app.Open(System.Environment.CurrentDirectory + "\\母版樣板.ai");
            Illustrator.Layers vp_layers = doc_vp_out.Layers;
            
            Illustrator.GroupItem test_gi = Getbyname_groupitem("test", doc_art.GroupItems);
            
            try
            {
                Illustrator.PathItem test_path = lr_plan.PathItems[1];
                //Illustrator.GraphItem test_graphi = lr_plan.GraphItems[1];
                //Illustrator.GroupItem test_groupi = lr_plan.GroupItems[1];
                //Illustrator.MeshItem test_meshi = lr_plan.MeshItems[1];
                //Illustrator.PageItems test_pagei = lr_plan.PageItems[1];
                //Illustrator.PlacedItem test_placei = lr_plan.PlacedItems[1];
                //Illustrator.PluginItem test_plugini = lr_plan.PluginItems[1];
                //Illustrator.RasterItem test_rasteri = lr_plan.RasterItems[1];
                //Illustrator.SymbolItem test_symboli = lr_plan.SymbolItems[1]; 
                //Illustrator.CompoundPathItem test_compound = lr_plan.CompoundPathItems[1];
                string test_name = test_path.Name;
            }
            catch (Exception)
            {
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
            //app.Quit();
        }

        public void test_print_barcode()
        {
            Illustrator.Application app = new Illustrator.Application();

            Illustrator.Document doc_art = app.Open("D:\\程式文案\\016畫稿自動拼模\\條碼\\物件抓取.ai");

            //取得母版圖稿
            Illustrator.Layers lrs_plan = doc_art.Layers;
            Illustrator.Layer lr_plan = Getbyname_layer("test", lrs_plan);

            Illustrator.PathItem pos_path = Getbyname_pathitem("pos", lr_plan.PathItems);

            //barcode
            Barcode barcode = new Barcode();
            //barcode.Encode(BarcodeLib.TYPE.CODE128, "ABCABCABCABC").Encode();
            barcode.Encode(Type.Code128B, "ABCABCABCABC");
            //barcode.Encode(Type.Code128B, "ABCDEFGHIJKLM").Encode();
            //barcode.Encode(Type.Code128B, "NOPQRSTUVWXYZ").Encode();

            string bar_data = barcode.EncodedValue;

            Illustrator.GroupItem bar_group = lr_plan.GroupItems.Add();

            bar_group.Name = "barcode";

            bar_group.Top = pos_path.Top;
            bar_group.Left = pos_path.Left;

            double point_x = pos_path.Left + 5;
            double exam_x = pos_path.Left + 5;

            double point_width = 0.25;

            
            for (int i = 0; i < bar_data.Length; i++)
            { 
                char c = bar_data[i];
                double line_width = 0.25;
                double poetd = 0;
                if (c == '1')
                {
                    if (bar_data.Length > i + 1)
                    { 
                        if (bar_data[i + 1] == '0')
                        {
                            poetd = 1;
                            point_width = 1 * point_width;
                        }
                        else
                        {
                            if (bar_data.Length > i + 2)
                            {
                                if (bar_data[i + 2] == '0')
                                {
                                    line_width = 2 * point_width;
                                    poetd = 2;
                                    i++;
                                }
                                else
                                {
                                    if (bar_data[i + 3] == '0')
                                    {
                                        line_width = 3 * point_width;
                                        poetd = 3;
                                        i += 2;
                                    }
                                    else
                                    {
                                        line_width = 4 * point_width;
                                        poetd = 4;
                                        i += 3;
                                    }
                                }
                            }
                            else
                            {
                                line_width = 2 * point_width;
                                poetd = 2;
                                i++;
                            }
                        }
                    }
                    /*
                    以線寬調整
                    if (bar_data.Length > i + 1)
                    {
                        if (bar_data[i + 1] == '0')
                        {
                            poetd = 1;
                            point_width = 1 * point_width;
                        }
                        else
                        {
                            if (bar_data.Length > i + 2)
                            {
                                if (bar_data[i + 2] == '0')
                                {
                                    line_width = 2 * point_width;
                                    poetd = 1.5;
                                    point_x += point_width * 0.5;
                                    i++;
                                }
                                else
                                {
                                    if (bar_data[i + 3] == '0')
                                    {
                                        line_width = 3 * point_width;
                                        poetd = 2;
                                        point_x += point_width;
                                        i += 2;
                                    }
                                    else
                                    {
                                        line_width = 4 * point_width;
                                        poetd = 2.5;
                                        point_x += point_width * 1.5;
                                        i += 3;
                                    }
                                }
                            }
                            else
                            {
                                line_width = 2 * point_width;
                                poetd = 1.5;
                                point_x += point_width * 0.5;
                                i++;
                            }
                        }
                    }
                    pathItem.StrokeWidth = line_dth;
                    */

                    pos_path.Duplicate(bar_group);
                    Illustrator.PathItem pathItem = bar_group.PathItems[1];

                    pathItem.Top = pos_path.Top;
                    pathItem.Left = point_x;
                    pathItem.Height = 12;
                    //pathItem.StrokeWidth = line_width;
                    pathItem.Width = line_width;
                    point_x += point_width * poetd;
                }
                else {
                    point_width = 0.25;
                    /*
                    pos_path.Duplicate(bar_group);
                    Illustrator.PathItem pathItem = bar_group.PathItems[1];

                    pathItem.Top = pos_path.Top;
                    pathItem.Left = point_x;
                    pathItem.Height = 12;
                    //pathItem.StrokeWidth = 1;
                    pathItem.Width = line_width;
                    Illustrator.RGBColor color_ai = new Illustrator.RGBColor();
                    color_ai.Red = 0;
                    color_ai.Blue = 255;
                    color_ai.Green = 0;
                    //pathItem.StrokeColor = color_ai;
                    pathItem.FillColor = color_ai;
                    */
                    
                    point_x += point_width;
                }
            }

            point_x = 5;
            point_width = 0.25;

            foreach (char ch in bar_data)
            {
                if (ch == '1')
                {
                    pos_path.Duplicate(lr_plan);
                    Illustrator.PathItem pathItem = lr_plan.PathItems[1];
                    pathItem.Top = pos_path.Top - 24;
                    pathItem.Left = exam_x;
                    pathItem.Height = 12;
                    //pathItem.StrokeWidth = point_width;
                    pathItem.Width = point_width;
                }

                exam_x += point_width;
            }
            bar_group.Height = 12 * 2.835;
            bar_group.Width = 28 * 2.835;


            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                Illustrator.IllustratorSaveOptions saveOptions = new Illustrator.IllustratorSaveOptions();
                saveOptions.Compatibility = Illustrator.AiCompatibility.aiIllustrator17;
                saveOptions.FlattenOutput = Illustrator.AiOutputFlattening.aiPreserveAppearance;
                doc_art.SaveAs(saveFileDialog1.FileName, saveOptions);
                doc_art.Close();


                MessageBox.Show("輸出資料OK", "提醒");
            }

            saveFileDialog1.FileName = "";

            //doc_art.Close();
            app.Quit();
        }

        public void create_barcode_fun_test()
        {
            Illustrator.Application app = new Illustrator.Application();

            Illustrator.Document doc_database = app.Open(System.Environment.CurrentDirectory + "\\條碼範本.ai");
            Illustrator.Layer lay_base = Getbyname_layer("樣板", doc_database.Layers);

            Illustrator.Document doc_art = app.Open("D:\\程式文案\\016畫稿自動拼模\\條碼\\物件抓取.ai");

            //取得母版圖稿
            Illustrator.Layers lrs_plan = doc_art.Layers;
            Illustrator.Layer lr_plan = Getbyname_layer("test", lrs_plan);

            Illustrator.PathItem pos_path = Getbyname_pathitem("pos", lr_plan.PathItems);


            //barcode
            Barcode barcode = new Barcode();
            string test_str = "ABCABCABCABC";
            barcode.Encode(Type.Code128B, test_str);
            //barcode.Encode(Type.Code128B, test_str).Encode();
            //barcode.Encode(Type.Code128B, "ABCDEFGHIJKLM").Encode();
            //barcode.Encode(Type.Code128B, "NOPQRSTUVWXYZ").Encode();

            string bar_data = barcode.EncodedValue;

            Illustrator.GroupItem bar_group = lr_plan.GroupItems.Add();

            bar_group.Name = "barcode";

            bar_group.Top = pos_path.Top;
            bar_group.Left = pos_path.Left;

            double point_x = pos_path.Left + 5;
            double exam_x = pos_path.Left + 5;

            double point_width = 0.25;

            Illustrator.GroupItem apply_print = Getbyname_groupitem("開始",lay_base.GroupItems);

            apply_print.Duplicate(lr_plan);
            Illustrator.GroupItem per_group = lr_plan.GroupItems[1];

            per_group.Top = pos_path.Top;
            per_group.Left = point_x;
            per_group.Width = point_width * 10;
            point_x += 2.5;

            apply_print = Getbyname_layer("ABC", doc_database.Layers).GroupItems[1];
            for (int i = 0; i < 4; i++)
            {
                apply_print.Duplicate(lr_plan);
                per_group = lr_plan.GroupItems[1];
                per_group.Top = pos_path.Top;
                per_group.Left = point_x;
                per_group.Width = point_width * 10 * 3;
                point_x += 2.5 * 3;
            }
            /*
            foreach (char ch in test_str)
            {
                apply_print = Getbyname_groupitem(ch.ToString(), lay_base.GroupItems);
                //apply_print = Getbyname_layer(ch.ToString(), doc_database.Layers).GroupItems[1];

                apply_print.Duplicate(lr_plan);
                per_group = lr_plan.GroupItems[1];
                per_group.Top = pos_path.Top;
                per_group.Left = point_x;
                per_group.Width = point_width * 10;
                point_x += 2.5;
            }
            */
            //從檢查碼開始
            int str_index = bar_data.Length - 25;

            for (int i = str_index; i < bar_data.Length; i++)
            {
                char c = bar_data[i];
                double line_width = 0.25;
                double poetd = 0;
                if (c == '1')
                {
                    if (bar_data.Length > i + 1)
                    {
                        if (bar_data[i + 1] == '0')
                        {
                            poetd = 1;
                            point_width = 1 * point_width;
                        }
                        else
                        {
                            if (bar_data.Length > i + 2)
                            {
                                if (bar_data[i + 2] == '0')
                                {
                                    line_width = 2 * point_width;
                                    poetd = 2;
                                    i++;
                                }
                                else
                                {
                                    if (bar_data[i + 3] == '0')
                                    {
                                        line_width = 3 * point_width;
                                        poetd = 3;
                                        i += 2;
                                    }
                                    else
                                    {
                                        line_width = 4 * point_width;
                                        poetd = 4;
                                        i += 3;
                                    }
                                }
                            }
                            else
                            {
                                line_width = 2 * point_width;
                                poetd = 2;
                                i++;
                            }
                        }
                    }
                    pos_path.Duplicate(bar_group);
                    Illustrator.PathItem pathItem = bar_group.PathItems[1];

                    pathItem.Top = pos_path.Top;
                    pathItem.Left = point_x;
                    pathItem.Height = 12;
                    pathItem.Width = line_width;
                    point_x += point_width * poetd;
                }
                else
                {
                    point_width = 0.25;
                    point_x += point_width;
                }
            }

            point_x = 5;
            point_width = 0.25;

            foreach (char ch in bar_data)
            {
                if (ch == '1')
                {
                    pos_path.Duplicate(lr_plan);
                    Illustrator.PathItem pathItem = lr_plan.PathItems[1];
                    pathItem.Top = pos_path.Top - 24;
                    pathItem.Left = exam_x;
                    pathItem.Height = 12;
                    //pathItem.StrokeWidth = point_width;
                    pathItem.Width = point_width;
                }

                exam_x += point_width;
            }
            bar_group.Height = 12 * 2.835;
            bar_group.Width = 28 * 2.835;


            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                Illustrator.IllustratorSaveOptions saveOptions = new Illustrator.IllustratorSaveOptions();
                saveOptions.Compatibility = Illustrator.AiCompatibility.aiIllustrator17;
                saveOptions.FlattenOutput = Illustrator.AiOutputFlattening.aiPreserveAppearance;
                doc_art.SaveAs(saveFileDialog1.FileName, saveOptions);
                doc_art.Close();


                MessageBox.Show("輸出資料OK", "提醒");
            }

            saveFileDialog1.FileName = "";

            //doc_art.Close();
            app.Quit();
        }

        public void Control_artboard()
        {
            Illustrator.Application app = new Illustrator.Application();

            //app.ActiveDocument.Selection = null;
            Illustrator.Document doc_art = app.Open("D:\\程式文案\\016畫稿自動拼模\\條碼\\物件抓取.ai");

            Illustrator.Artboard artboard = doc_art.Artboards[1];

            dynamic aaa = artboard.ArtboardRect;
            aaa[0] = artboard.ArtboardRect[0] + 100 * 2.835;
            aaa[2] = artboard.ArtboardRect[2];

            doc_art.Artboards.Add(aaa);

            Illustrator.Artboard new_ab = doc_art.Artboards[2];
            new_ab.Name = "123";
            dynamic bbb = new_ab.ArtboardRect;
            new_ab.ArtboardRect[0] = bbb[0] + 30;

            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                Illustrator.IllustratorSaveOptions saveOptions = new Illustrator.IllustratorSaveOptions();
                saveOptions.Compatibility = Illustrator.AiCompatibility.aiIllustrator17;
                saveOptions.FlattenOutput = Illustrator.AiOutputFlattening.aiPreserveAppearance;
                doc_art.SaveAs(saveFileDialog1.FileName, saveOptions);
                doc_art.Close();
                MessageBox.Show("輸出資料OK", "提醒");
            }

            saveFileDialog1.FileName = "";

            //doc_art.Close();
            app.Quit();
        }
    }
}
