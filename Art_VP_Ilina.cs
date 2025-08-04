using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using Illustrator;

namespace Auto_Plan_Art_Plate
{
    internal class Art_VP_Ilina
    {
        string file_name;
        List<Ilina_vp_data> ilina_Vp_Datas = new List<Ilina_vp_data>();
        SaveFileDialog saveFileDialog1;
        public Art_VP_Ilina(string file_name, List<Ilina_vp_data> vP_Datas, SaveFileDialog saveFileDialog1) 
        {
            this.file_name = file_name;
            ilina_Vp_Datas = vP_Datas;
            this.saveFileDialog1 = saveFileDialog1;
        }

        public void Ilina_check_vp()
        {
            Illustrator.Application app = new Illustrator.Application();
            Illustrator.Document doc_art = app.Open(file_name);
            //取得母版圖稿
            Illustrator.Layers lrs_plan = doc_art.Layers;
            Illustrator.Layer lr_plan = Getbyname_layer("母版", lrs_plan);

            string group_name = "核稿";
            /*
            if (vp_Datas[0].no == 1)
                group_name = "拼模";
            */
            //Illustrator.GroupItem Inheritance_item = lr_plan.GroupItems[1];
            Illustrator.GroupItem Inheritance_item = Getbyname_groupitem(group_name, lr_plan.GroupItems);
            //取得母版圖層
            Illustrator.Document doc_vp_out = app.Open(System.Environment.CurrentDirectory + "\\母版樣板.ai");
            Illustrator.Layers vp_layers = doc_vp_out.Layers;

            Illustrator.Layer vp_oring = vp_layers[1];

            //取定位點
            double location_p_x = 21 * 2.835 - 1.55;
            double location_p_y = -19 * 2.835;
            //取原始距離
            double org_loc_p_x = location_p_x;
            double org_loc_p_y = location_p_y;

            int print_count = 1;
            int print_num = 1;
            string temp = ilina_Vp_Datas[0].Sheet_name;

            foreach (Ilina_vp_data vp in ilina_Vp_Datas)
            {
                vp_layers.Add();
                Illustrator.Layer vp_l = vp_layers[1];
                vp_l.Name = "圖層 " + print_count.ToString();

                Inheritance_item.Duplicate(vp_l);
                Illustrator.GroupItem dup_group = vp_l.GroupItems[1];
                //判斷轉換序號時，需要整筆圖面換列顯示
                if (vp.Sheet_name != temp)
                {
                    temp = vp.Sheet_name;
                    print_num = 1;

                    location_p_x = org_loc_p_x;
                    dup_group.Left = org_loc_p_x;

                    location_p_y -= (dup_group.Height + 7.5 * 3 * 2.835);
                    dup_group.Top = location_p_y;
                    org_loc_p_y -= (dup_group.Height * 3 + 7.5 * (3+2) * 2.835);
                }
                else 
                {
                    dup_group.Top = location_p_y;
                    dup_group.Left = location_p_x;
                }
                //產生barcode
                //取得定位線條
                Illustrator.PathItem position_path = Getbyname_pathitem("pos", dup_group.PathItems);
                Art_Barcode art_b = new Art_Barcode(vp.Code128, 28, 12, dup_group, position_path);
                Thread th_set_barcode = new Thread(new ThreadStart(art_b.Print_in_art_barcode));
                th_set_barcode.Start();

                location_p_x += dup_group.Width + 12 * 2.835 - 1.3;
                if (print_num % 12 == 0)
                {
                    location_p_x += (-12 + 21 + 5 + 21 + 9) * 2.835 - 1.3;
                    location_p_y = org_loc_p_y;
                }
                else if (print_num % 4 == 0)
                {
                    location_p_y -= dup_group.Height + 7.5 * 2.835 ;
                    location_p_x -= dup_group.Width * 4 + 12 * 4 * 2.835 - 5.19;
                }
                /*
                if (vp.no != 0)
                    Getbyname_TextFrame("No", dup_group.TextFrames).Contents = vp.no.ToString();
                */
                Getbyname_TextFrame("style_1", dup_group.TextFrames).Contents = vp.No_Product;
                Getbyname_TextFrame("style_2", dup_group.TextFrames).Contents = vp.No_Product;

                Getbyname_TextFrame("color_1", dup_group.TextFrames).Contents = vp.Color;
                Getbyname_TextFrame("color_2", dup_group.TextFrames).Contents = vp.Color;

                Getbyname_TextFrame("size_1", dup_group.TextFrames).Contents = vp.Size;
                Getbyname_TextFrame("size_2", dup_group.TextFrames).Contents = vp.Size;

                Getbyname_TextFrame("price_1", dup_group.TextFrames).Contents = vp.Price;
                Getbyname_TextFrame("price_2", dup_group.TextFrames).Contents = vp.Price;

                th_set_barcode.Join();

                print_count++;
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
        }

        public void Ilina_plate_vp()
        { 
            
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
