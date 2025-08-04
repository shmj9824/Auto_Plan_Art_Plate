using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Auto_Plan_Art_Plate
{
    internal class Art_VP_Rei_Id : Art_VP_File
    {
        //public List<VP_data> vp_Datas = new List<VP_data>();
        string vp_file;
        bool old_version = false;
        SaveFileDialog saveFileDialog1;

        public Art_VP_Rei_Id(string file_name, List<VP_data> vP_Datas, SaveFileDialog saveFileDialog1) : base(file_name, vP_Datas, saveFileDialog1)
        {
            vp_file = file_name;
            vp_Datas = vP_Datas;
            this.saveFileDialog1 = saveFileDialog1;
        }

        public void Set_old_version_value(bool aa)
        { 
            old_version = aa;
        }

        public void REI_ID_forplate()
        {
            Illustrator.Application app = new Illustrator.Application();
            Illustrator.Document doc_art = app.Open(vp_file);
            //取得母版圖稿
            Illustrator.Layers lrs_plan = doc_art.Layers;
            Illustrator.Layer lr_plan = Getbyname_layer("母版", lrs_plan);
            

            //取得母版圖層
            Illustrator.Document doc_vp_out = app.Open(System.Environment.CurrentDirectory + "\\母版樣板.ai");
            Illustrator.Layers vp_layers = doc_vp_out.Layers;

            Illustrator.Layer vp_oring = vp_layers[1];

            //取定位點
            double location_p_x = 20 * 2.835;
            double location_p_y = -20 * 2.835;
            int print_num = 1;

            foreach (VP_data vp in vp_Datas)
            {
                string lay_na = "拼模";
                if (vp.item6 != null)
                    if (vp.item6.Contains("PER"))
                        lay_na = "羽絨用_拼模";
                Illustrator.GroupItem Inheritance_item = Getbyname_groupitem(lay_na, lr_plan.GroupItems);

                vp_layers.Add();
                Illustrator.Layer vp_l = vp_layers[1];
                vp_l.Name = "圖層 " + print_num.ToString();

                Inheritance_item.Duplicate(vp_l);
                Illustrator.GroupItem dup_group = vp_l.GroupItems[1];
                dup_group.Top = location_p_y;
                dup_group.Left = location_p_x;
                location_p_x += dup_group.Width + 10 * 2.835;

                if (print_num % 6 == 0)
                {
                    location_p_y -= dup_group.Height + 10 * 2.835;
                    location_p_x = 20 * 2.835;
                }
                if (!old_version)
                {
                    Getbyname_TextFrame("No", dup_group.TextFrames).Contents = vp.no.ToString();
                    Getbyname_TextFrame("Sku", dup_group.TextFrames).Contents = vp.item1;
                    Getbyname_TextFrame("Style_season", dup_group.TextFrames).Contents = vp.item2;
                    //241128改款
                    Getbyname_TextFrame("Rn_date", dup_group.TextFrames).Contents = vp.item3;
                    Getbyname_TextFrame("verder", dup_group.TextFrames).Contents = vp.item7;
                    Getbyname_TextFrame("Origin_duration", dup_group.TextFrames).Contents = vp.item4;
                    Getbyname_TextFrame("Sub_p", dup_group.TextFrames).Contents = vp.item5;
                }
                else
                {
                    Getbyname_TextFrame("No", dup_group.TextFrames).Contents = vp.no.ToString();
                    Getbyname_TextFrame("Sku", dup_group.TextFrames).Contents = vp.item1;
                    Getbyname_TextFrame("Style_season", dup_group.TextFrames).Contents = vp.item2;
                    Getbyname_TextFrame("Rn_verder", dup_group.TextFrames).Contents = vp.item3;
                    Getbyname_TextFrame("Origin_duration", dup_group.TextFrames).Contents = vp.item4;
                    Getbyname_TextFrame("Sub_p", dup_group.TextFrames).Contents = vp.item5;
                }
                
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

        public void REI_ID_forcheck()
        {
            Illustrator.Application app = new Illustrator.Application();
            Illustrator.Document doc_art = app.Open(vp_file);
            //取得母版圖稿
            Illustrator.Layers lrs_plan = doc_art.Layers;
            Illustrator.Layer lr_plan = Getbyname_layer("母版", lrs_plan);
            

            //取得母版圖層
            Illustrator.Document doc_vp_out = app.Open(System.Environment.CurrentDirectory + "\\母版樣板.ai");
            Illustrator.Layers vp_layers = doc_vp_out.Layers;

            Illustrator.Layer vp_oring = vp_layers[1];

            //取定位點
            double location_p_x = 20 * 2.835;
            double location_p_y = -20 * 2.835;
            int print_num = 1;
            //新增畫稿圖層把匯出資料放入
            vp_layers.Add();
            Illustrator.Layer vp_l = vp_layers[1];
            vp_l.Name = "畫稿";

            foreach (VP_data vp in vp_Datas)
            {
                //要改羽絨款
                string lay_na = "核稿";
                if (vp.item8 != null)
                    if (vp.item8.Contains("羽絨"))
                        lay_na = "羽絨用_核稿";
                Illustrator.GroupItem Inheritance_item = Getbyname_groupitem(lay_na, lr_plan.GroupItems);

                Inheritance_item.Duplicate(vp_l);

                Illustrator.GroupItem dup_group = vp_l.GroupItems[1];
                dup_group.Top = location_p_y;
                dup_group.Left = location_p_x;
                location_p_x += dup_group.Width + 10 * 2.835;

                if (print_num % 6 == 0)
                {
                    location_p_y -= dup_group.Height + 10 * 2.835;
                    location_p_x = 20 * 2.835;
                }
                string first_row = "#" + vp.no.ToString() + "," + vp.item1;
                Getbyname_TextFrame("序, ART#", dup_group.TextFrames).Contents = first_row;

                Illustrator.TextFrame po = Getbyname_TextFrame("PO", dup_group.TextFrames);
                po.Contents = vp.item2;
                if (po.Width > 20 * 2.835)
                    po.Width = 20 * 2.835;

                //Getbyname_TextFrame("PO", dup_group.TextFrames).Contents = vp.item2;
                if (!old_version)
                {
                    Getbyname_TextFrame("Sku", dup_group.TextFrames).Contents = vp.item3;
                    Getbyname_TextFrame("Style_season", dup_group.TextFrames).Contents = vp.item4;
                    Getbyname_TextFrame("Rn_date", dup_group.TextFrames).Contents = vp.item5;
                    //241128改款
                    Getbyname_TextFrame("verder", dup_group.TextFrames).Contents = vp.item9;
                    Getbyname_TextFrame("Origin_duration", dup_group.TextFrames).Contents = vp.item6;
                    Getbyname_TextFrame("Sub_p", dup_group.TextFrames).Contents = vp.item7;
                }
                else
                {
                    //Getbyname_TextFrame("PO", dup_group.TextFrames).Contents = vp.item2;
                    Getbyname_TextFrame("Sku", dup_group.TextFrames).Contents = vp.item3;
                    Getbyname_TextFrame("Style_season", dup_group.TextFrames).Contents = vp.item4;
                    Getbyname_TextFrame("Rn_verder", dup_group.TextFrames).Contents = vp.item5;
                    Getbyname_TextFrame("Origin_duration", dup_group.TextFrames).Contents = vp.item6;
                    Getbyname_TextFrame("Sub_p", dup_group.TextFrames).Contents = vp.item7;
                }
                
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
    }
}
