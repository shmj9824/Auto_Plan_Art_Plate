using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

using BarcodeStandard;
//using BarcodeLib;
//using Type = BarcodeStandard.Type;

namespace Auto_Plan_Art_Plate
{
    
    public partial class Main_page : Form
    {
        public Main_page()
        {
            InitializeComponent();
            Dataset_excel dse = new Dataset_excel();
            dse.Open_File(System.Environment.CurrentDirectory + "\\畫稿資料集.xlsx");
            pdis = dse.get_list_excel();
            foreach (Plate_data_item item in pdis)
            { 
                cb_dataset.Items.Add(item.name);
                cb_vp_dataset.Items.Add(item.name);
            }
        }
        List<Plate_data_item> pdis;
        string file_art_name = "";
        string file_plate_name = "";
        string file_exam_art_end = "";
        string file_reference = "";
        int plate_int = 0;
        int art_pic_num = 0;
        List<Dictionary<int, int>> Plate_data_list;
        

        private void bt_plan_start_Click(object sender, EventArgs e)
        {
            if (file_art_name == "" || file_plate_name == "")
                return;
            //241128泰維克標修改
            if (cb_dataset.Text.ToString() == "REI ID-5" && !CB_id5.Checked)
            {
                MessageBox.Show("已於241128泰維克標修改，暫時封鎖拼模，如需使用請開啟功能");
                return;
            }

            if (Plate_data_list.Count > 0)
            {
                int rotate_num = Convert.ToInt32(cb_Rotate.SelectedItem.ToString());
                /*
                Art_plate art_Plate = new Art_plate();
                if (art_Plate.check_doc_art(file_art_name, art_pic_num))
                {
                    //art_Plate.Tyvek_plate(ref rotate_num,
                    //file_art_name, file_plate_name,
                    //tb_art_plan_explan.Text, ref Plate_data_list, ref saveFileDialog1);
                    MessageBox.Show("OK");
                }
                else
                {
                    MessageBox.Show("畫稿圖層數量有誤");
                }
                */
                //int max_col = Convert.ToInt32(nud_plate_col.Value.ToString());
                int max_col = Convert.ToInt32(nud_plate_col.Value);
                int max_row = Convert.ToInt32(nud_plate_row.Value.ToString());
                string sel_dataset = cb_dataset.Text.ToString();
                DateTime tod = DateTime.Now;
                if (sel_dataset == "REI ID-1" || sel_dataset == "REI ID-5")
                    new Art_plate().Tyvek_plate(ref rotate_num, ref max_col, ref max_row,
                        file_art_name, file_plate_name,
                        tb_art_plan_explan.Text, ref Plate_data_list, ref saveFileDialog1);
                else if (sel_dataset == "ilina")
                    new Art_plate().iliea_plate(ref rotate_num, ref max_col, ref max_row,
                        file_art_name, file_plate_name,
                        tb_art_plan_explan.Text, ref Plate_data_list, ref saveFileDialog1);
                else if (sel_dataset == "BonBon")
                    new Art_plate().bonbon_plate(ref rotate_num, ref max_col, ref max_row,
                        file_art_name, file_plate_name,
                        tb_art_plan_explan.Text, ref Plate_data_list, ref saveFileDialog1);
                else if (sel_dataset == "BonBon-Basic")
                    new Art_plate().Basic_bonbon_plate(ref rotate_num, ref max_col, ref max_row,
                        file_art_name, file_plate_name,
                        tb_art_plan_explan.Text, ref Plate_data_list,ref tod, ref saveFileDialog1);
            }
            else
                MessageBox.Show("ERROR");
        }

        private void bt_sel_art_Click(object sender, EventArgs e)
        {
            string filename = "";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.FileName = "";
            openFileDialog1.ShowDialog();
            if (openFileDialog1.FileName != "")
            {
                filename = openFileDialog1.FileName;
                openFileDialog1.FileName = "";
                openFileDialog1.Dispose();
            }

            if (filename != "")
            { 
                file_art_name = filename;
                label_file_art.Text = file_art_name;
            }
        }

        private void bt_sel_plate_Click(object sender, EventArgs e)
        {
            string filename = "";
            openFileDialog2.FilterIndex = 1;
            openFileDialog2.FileName = "";
            openFileDialog2.ShowDialog();

            if (openFileDialog2.FileName != "")
            {
                filename = openFileDialog2.FileName;
                openFileDialog2.FileName = "";
                openFileDialog2.Dispose();
            }

            if (filename != "")
            { 
                file_plate_name = filename;
                label_file_plate.Text = file_plate_name;
            }
        }

        private void bt_sel_plan_excel_Click(object sender, EventArgs e)
        {
            string filename = "";
            openFileDialog3.FilterIndex = 1;
            openFileDialog3.FileName = "";
            openFileDialog3.ShowDialog();

            if (openFileDialog3.FileName != "")
            {
                filename = openFileDialog3.FileName;
                openFileDialog3.FileName = "";
                openFileDialog3.Dispose();
            }

            if (filename != "")
            {
                if (nud_plate_col.Value == 0 || nud_plate_row.Value == 0)
                {
                    MessageBox.Show("請輸入模板資料");
                    return;
                }

                plate_int = Convert.ToInt32(nud_plate_col.Value * nud_plate_row.Value);
                Call_plate_excel cpe = new Call_plate_excel(plate_int);
                cpe.Open_File(filename);

                string check_plate_str = cpe.mess;
                bt_plan_start.Enabled = true;
                /*
                try
                {
                    int in_excel_col = Convert.ToInt32(check_plate_str[0].ToString());
                    int in_excel_row = Convert.ToInt32(check_plate_str[2].ToString());

                    
                    if (nud_plate_col.Value != in_excel_col || nud_plate_row.Value != in_excel_row)
                    { 
                        bt_plan_start.Enabled = false;
                        MessageBox.Show("模板行列錯誤");
                    }
                }
                catch (Exception)
                {
                }
                */

                if (cpe.mess_plate_sum != "")
                { 
                    MessageBox.Show(cpe.mess_plate_sum);
                    bt_plan_start.Enabled = false;
                }
                if (bt_plan_start.Enabled)
                { 
                    Plate_data_list = cpe.Get_list_plate_data();
                    art_pic_num = cpe.data_count;
                    MessageBox.Show("成功");
                }
            }
        }

        private void cb_dataset_SelectedValueChanged(object sender, EventArgs e)
        {
            string checking_str = cb_dataset.Text;
            foreach (Plate_data_item pdi in pdis)
            {
                if (pdi.name == checking_str)
                {
                    nud_plate_row.Value = pdi.row;
                    nud_plate_col.Value = pdi.col;
                    int value_index = cb_Rotate.Items.IndexOf(Convert.ToString(pdi.rotate));
                    cb_Rotate.SelectedItem = cb_Rotate.Items[value_index];
                }
            }
        }

        private void bt_PV_table_Click(object sender, EventArgs e)
        {
            string filename = "";
            openFileDialog3.FilterIndex = 1;
            openFileDialog3.FileName = "";
            openFileDialog3.ShowDialog();

            if (openFileDialog3.FileName != "")
            {
                filename = openFileDialog3.FileName;
                openFileDialog3.FileName = "";
                openFileDialog3.Dispose();
            }
            if (filename != "")
            { 
                file_name_VP = filename;
                lab_dt.Text = filename;
            }
        }
        string file_name_VP = "";
        
        private void bt_VariPrint_Click(object sender, EventArgs e)
        {
            bool check_art = false;
            if (cb_print_way.Text == "核稿")
                check_art = true;
            bool bonbon_year_plus = false;
            if (file_name_VP != "" && file_exam_art_end != "" && cb_vp_dataset.Text != "" && cb_print_way.Text != "")
            {
                //撈取火併或備拼Excel檔案
                List<VP_data> vp_datas = new List<VP_data>();
                List<Ilina_vp_data> ivds = new List<Ilina_vp_data>();
                List<Bonbon_vp_data> bvds = new List<Bonbon_vp_data>();
                List<Bonbon_Basic_vpdata> bbvds = new List<Bonbon_Basic_vpdata>();
                Dictionary<string, (string, string)> dic_po_no = new Dictionary<string, (string, string)>();
                if (cb_vp_dataset.Text == "REI ID-1" || cb_vp_dataset.Text == "REI ID-5")
                {
                    if (check_art)
                    {
                        VP_Tyvek_Excel vP_Tyvek = new VP_Tyvek_Excel(VP_Mode.check);
                        vP_Tyvek.Set_old_version_value(CB_oldver.Checked);
                        vP_Tyvek.Open_File(file_name_VP);
                        vp_datas = vP_Tyvek.vp_Datas;
                    }
                    else
                    {
                        VP_Tyvek_Excel vP_Tyvek = new VP_Tyvek_Excel(VP_Mode.plate);
                        vP_Tyvek.Set_old_version_value(CB_oldver.Checked);
                        vP_Tyvek.Open_File(file_name_VP);
                        vp_datas = vP_Tyvek.vp_Datas;
                    }
                }
                else if (cb_vp_dataset.Text == "ilina")
                {
                    if (check_art)
                    {
                        VP_Ilina_Excel vP_Ilina = new VP_Ilina_Excel(VP_Mode.check);
                        vP_Ilina.Open_File(file_name_VP);
                        ivds = vP_Ilina.Get_Ilina_Vp_Datas();
                    }
                    else
                    {
                        VP_Ilina_Excel vP_Ilina = new VP_Ilina_Excel(VP_Mode.plate);
                        vP_Ilina.Open_File(file_name_VP);
                        vp_datas = vP_Ilina.vp_Datas;
                    }
                }
                else if (cb_vp_dataset.Text == "BonBon")
                {
                    if (check_art)
                    {
                        bonbon_year_plus = cb_year_p.Checked;

                        VP_Bonbon_Excel vP_Bonbon = new VP_Bonbon_Excel(VP_Mode.check);
                        vP_Bonbon.Open_File(file_name_VP);
                        bvds = vP_Bonbon.bonbon_Vps;
                        dic_po_no = vP_Bonbon.dic_po_no;
                        Bonbon_Reference_file bonbon_Reference_ = new Bonbon_Reference_file(bvds, "write");
                        bonbon_Reference_.Save_Exam_File(System.Environment.CurrentDirectory + "\\bonbon參考檔.xlsx", saveFileDialog2);
                    }
                    else
                    {
                        Bonbon_Reference_file bonbon_Reference = new Bonbon_Reference_file("read");
                        bonbon_Reference.Open_File(file_reference);

                        VP_Bonbon_Excel vP_Bonbon = new VP_Bonbon_Excel(VP_Mode.plate, bonbon_Reference.art_no_Ingredient);
                        vP_Bonbon.Open_File(file_name_VP);
                        bvds = vP_Bonbon.bonbon_Vps;
                    }
                }
                else if (cb_vp_dataset.Text == "BonBon-Basic")
                {
                    if (check_art)
                    {
                        bonbon_year_plus = cb_year_p.Checked;

                        VP_Bonbon_Basic_Excel vP_Bonbon_Basic = new VP_Bonbon_Basic_Excel(VP_Mode.check);
                        vP_Bonbon_Basic.Open_File(file_name_VP);
                        bbvds = vP_Bonbon_Basic.bonbon_Vps;
                        dic_po_no = vP_Bonbon_Basic.dic_po_no;
                        Bonbon_Basic_Reference_file bonbon_Reference_ = new Bonbon_Basic_Reference_file(bbvds, "write");
                        bonbon_Reference_.Save_Exam_File(System.Environment.CurrentDirectory + "\\bonbon_basic參考檔.xlsx", saveFileDialog2);
                    }
                    else
                    {
                        Bonbon_Basic_Reference_file bonbon_Reference = new Bonbon_Basic_Reference_file("read");
                        bonbon_Reference.Open_File(file_reference);

                        //
                        VP_Bonbon_Basic_Excel vP_Bonbon_Basic = new VP_Bonbon_Basic_Excel(VP_Mode.plate, bonbon_Reference.art_no_Ingredient);
                        vP_Bonbon_Basic.Open_File(file_name_VP);
                        bbvds = vP_Bonbon_Basic.bonbon_Vps;
                    }
                }
                
                //illustrator執行部分
                
                if (vp_datas.Count() > 0 || ivds.Count() > 0 || bvds.Count() > 0 || bbvds.Count() > 0)
                {
                    saveFileDialog1.FileName = "";
                    if (cb_vp_dataset.Text == "REI ID-1" || cb_vp_dataset.Text == "REI ID-5")
                    {

                        if (check_art)
                        {
                            Art_VP_Rei_Id art_VP_Rei_Id = new Art_VP_Rei_Id(file_exam_art_end, vp_datas, saveFileDialog1);
                            art_VP_Rei_Id.Set_old_version_value(CB_oldver.Checked);
                            art_VP_Rei_Id.REI_ID_forcheck();
                        }
                        else
                        {
                            Art_VP_Rei_Id art_VP_Rei_Id = new Art_VP_Rei_Id(file_exam_art_end, vp_datas, saveFileDialog1);
                            art_VP_Rei_Id.Set_old_version_value(CB_oldver.Checked);
                            art_VP_Rei_Id.REI_ID_forplate();
                        }
                    }
                    else if (cb_vp_dataset.Text == "ilina")
                    {
                        if (check_art)
                        {
                            new Art_VP_Ilina(file_exam_art_end, ivds, saveFileDialog1).Ilina_check_vp();
                        }
                        else
                            new Art_VP_File(file_exam_art_end, vp_datas, saveFileDialog1).Ilina_vp();
                    }
                    else if (cb_vp_dataset.Text == "BonBon")
                    {
                        if (check_art)
                            new Art_VP_Bonbon(file_exam_art_end, bvds, dic_po_no, bonbon_year_plus, saveFileDialog1).Bonbon_check_vp();
                        else
                            new Art_VP_Bonbon(file_exam_art_end, bvds, saveFileDialog1).Bonbon_plate_vp();
                    }
                    else if (cb_vp_dataset.Text == "BonBon-Basic")
                    {
                        if (check_art)
                        { 
                            new Art_VP_Bonbon_Basic(file_exam_art_end, bbvds, dic_po_no, bonbon_year_plus, saveFileDialog1).Bonbon_check_vp();
                        }
                        else
                        { 
                            new Art_VP_Bonbon_Basic(file_exam_art_end, bbvds, saveFileDialog1).Bonbon_plate_vp();
                        }
                    }
                }
                
            }
            else
                MessageBox.Show("請選擇好檔案");
        }

        private void bt_exam_art_Click(object sender, EventArgs e)
        {
            string filename = "";
            openFileDialog2.FilterIndex = 1;
            openFileDialog2.FileName = "";
            openFileDialog2.ShowDialog();

            if (openFileDialog2.FileName != "")
            {
                filename = openFileDialog2.FileName;
                openFileDialog2.FileName = "";
                openFileDialog2.Dispose();
            }

            if (filename != "")
            {
                file_exam_art_end = filename;
                lab_vp.Text = filename;
            }
        }

        
        private void cb_vp_dataset_SelectedChanged(object sender, EventArgs e)
        {
            if (cb_vp_dataset.Text == "BonBon" || cb_vp_dataset.Text == "BonBon-Basic")
            {
                bt_Reference.Enabled = true;
                cb_year_p.Enabled = true;
            }
            else
            { 
                bt_Reference.Enabled = false;
                cb_year_p.Enabled = false;
            }
        }

        private void bt_Reference_Click(object sender, EventArgs e)
        {
            string filename = "";
            openFileDialog2.FilterIndex = 1;
            openFileDialog2.FileName = "";
            openFileDialog2.ShowDialog();

            if (openFileDialog2.FileName != "")
            {
                filename = openFileDialog2.FileName;
                openFileDialog2.FileName = "";
                openFileDialog2.Dispose();
            }

            if (filename != "")
            {
                file_reference = filename;
            }
        }

        List<string> list_language;
        Dictionary<string, string> dic_Shellfabric;
        Dictionary<string, string> dic_Lining;
        Dictionary<string, string> dic_Padding;
        Dictionary<string, string> dic_Suspenders;
        List<Dictionary<string, string>> dic_list_ziener = new List<Dictionary<string, string>>();
        Dictionary<string, List<string>> dic_c;
        List<string> list_acronym;
        DataTable dt;
        List<string> eng_list_v;
        private void bt_tran_Click(object sender, EventArgs e)
        {
            if (list_language == null)
            {
                MessageBox.Show("請參考編碼檔");
                return;
            }
            if (tb_ziener_input.Text == "")
            {
                MessageBox.Show("請輸入資料");
                return;
            }
            if (cb_sel_ziener.Text == "洗語")
            {
                tb_care_num_z.Text = "";
                string tran_wash_str = tb_ziener_input.Text;
                tran_wash_str = tran_wash_str.Replace("\r", "");

                List<string> list_wash_list = tran_wash_str.Split('\n').ToList();
                /*
                list_wash_list.ForEach(wash => {
                    tb_ziener_out.Text += wash + System.Environment.NewLine;
                });
                */
                List<string> check_care_list = new List<string>();
                foreach (DataRow dr in dt.Rows)
                {
                    check_care_list.Add(dr["ENG"].ToString());
                }
                List<int> dt_index = new List<int>();
                list_wash_list.ForEach(wash => {
                    if (wash != "")
                    { 
                        string check = wash[0].ToString().ToUpper() + wash.Substring(1);
                        if (check_care_list.Contains(check))
                        { 
                            //tb_care_num.Text += check_care_list.IndexOf(check) + " ";
                            dt_index.Add(check_care_list.IndexOf(check));
                        }
                        else
                            tb_care_num_z.Text += "Not Found" + " ";
                    }
                });
                dt_index.Sort();
                dt_index.ForEach(index => {
                    //tb_ziener_out.Text += dt.Rows[index]["ENG"].ToString() + System.Environment.NewLine;
                    tb_care_num_z.Text += (index + 1) + " ";
                });
                string output_care = "";
                foreach (string str_lang in list_language)
                {
                    string cir_str = "";
                    cir_str += str_lang + " ";
                    dt_index.ForEach(index => {
                        if (cir_str.Length > 6)
                            cir_str += " / " + dt.Rows[index][str_lang].ToString();
                        else
                            cir_str += dt.Rows[index][str_lang].ToString();
                    });

                    output_care += cir_str + System.Environment.NewLine;
                }
                tb_ziener_out.Text = output_care;
            }
            
        }

        private void Cb_sel_ziener_SelectedValueChanged(object sender, EventArgs e)
        {
            
            if (cb_sel_ziener.Text == "成份")
            {
                tb_ziener_input.Enabled = false;
                tb_ziener_out.Enabled = false;
                bt_import_exam.Enabled = true;
            }
            else {
                tb_ziener_input.Enabled = true;
                tb_ziener_out.Enabled = true;
                bt_import_exam.Enabled = false;
            }

        }
        private void Cb_sel_vaude_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cb_sel_vaude.Text == "成份")
            {
                bt_vaude_refer_excel.Enabled = false;
            }
            else
                bt_vaude_refer_excel.Enabled = true;
        }
        public Dictionary<string, string> Get_temp_dic(string detail)
        {
            detail = detail.ToUpper();
            if (dic_Shellfabric["ENG"] == detail)
                return dic_Shellfabric;
            if (dic_Lining["ENG"] == detail)
                return dic_Lining;
            if (dic_Padding["ENG"] == detail)
                return dic_Padding;
            if (dic_Suspenders["ENG"] == detail)
                return dic_Suspenders;
            return null;
        }
        private void bt_ziener_excel_Click(object sender, EventArgs e)
        {
            string filename = "";
            openFileDialog2.FilterIndex = 1;
            openFileDialog2.FileName = "";
            openFileDialog2.ShowDialog();

            if (openFileDialog2.FileName != "")
            {
                filename = openFileDialog2.FileName;
                openFileDialog2.FileName = "";
                openFileDialog2.Dispose();
            }

            if (filename != "")
            {
                Ziener_refer_excel ziener_Excel = new Ziener_refer_excel();
                ziener_Excel.Open_File(filename);
                list_language = ziener_Excel.get_list();
                //分類
                dic_Shellfabric = ziener_Excel.get_dic_title(1);
                dic_Lining = ziener_Excel.get_dic_title(2);
                dic_Padding = ziener_Excel.get_dic_title(3);
                dic_Suspenders = ziener_Excel.get_dic_title(4);

                dic_list_ziener.Add(dic_Shellfabric);
                dic_list_ziener.Add(dic_Lining);
                dic_list_ziener.Add(dic_Padding);
                dic_list_ziener.Add(dic_Suspenders);

                list_acronym = ziener_Excel.Get_list_acronym();

                dic_c = ziener_Excel.get_dic_list();
                dt = ziener_Excel.get_dt();
                //tb_out.Text = dic_c[list_language[4]][2];
                bt_ziener_tran.Enabled = true;
                MessageBox.Show("讀檔完成");
                //dataGridView1.DataSource = dt;
            }
        }

        private void bt_import_exam_Click(object sender, EventArgs e)
        {
            if (list_language == null)
            {
                MessageBox.Show("請先使用讀檔");
                return;
            }
            if (cb_sel_ziener.Text == "")
            { 
                MessageBox.Show("請選擇項目");
                return;
            }
            string filename = "";
            openFileDialog2.FilterIndex = 1;
            openFileDialog2.FileName = "";
            openFileDialog2.ShowDialog();

            if (openFileDialog2.FileName != "")
            {
                filename = openFileDialog2.FileName;
                openFileDialog2.FileName = "";
                openFileDialog2.Dispose();
            }

            if (filename != "")
            {
                if (!cb_ziener_acronym.Checked)
                {
                    Ziener_exam_excel zee = new Ziener_exam_excel(list_language, dic_list_ziener, dic_c);
                    zee.Save_Exam_File(filename, saveFileDialog2);
                }
                else
                {
                    Ziener_acronym zan = new Ziener_acronym(list_language, dic_list_ziener,dic_c, list_acronym);
                    zan.Save_Exam_File(filename, saveFileDialog2);
                    tb_ziener_out.Text = zan.Get_comp();
                }
                //tb_ziener_out.Text = zee.Get_complate_str();
            }
        }

        private void bt_vaude_refer_excel_Click(object sender, EventArgs e)
        {
            string filename = "";
            openFileDialog2.FilterIndex = 1;
            openFileDialog2.FileName = "";
            openFileDialog2.ShowDialog();

            if (openFileDialog2.FileName != "")
            {
                filename = openFileDialog2.FileName;
                openFileDialog2.FileName = "";
                openFileDialog2.Dispose();
            }

            if (filename != "")
            {
                Import_Vaude_Care import_Vaude_Care = new Import_Vaude_Care();
                import_Vaude_Care.Open_File(filename);
                eng_list_v = import_Vaude_Care.Get_eng_list();
                MessageBox.Show("OK");
            }
        }

        private void bt_str_apply_Click(object sender, EventArgs e)
        {
            if (eng_list_v == null && cb_sel_vaude.Text == "洗語")
            {
                MessageBox.Show("請參考編碼檔");
                return;
            }
            if (tb_Vaude_input.Text == "")
            {
                MessageBox.Show("請輸入資料");
                return;
            }
            string need_cut_str = tb_Vaude_input.Text;
            tb_Vaude_out.Text = "";
            tb_care_num_v.Text = "";
            if (cb_sel_vaude.Text == "成份")
            {
                //need_cut_str = need_cut_str.Replace("\r", "");
                List<string> list_all = new List<string>();
                if (need_cut_str != "")
                {
                    int next_index = need_cut_str.LastIndexOf('[');
                    while (next_index >= 0)
                    {
                        string need_apply = need_cut_str.Substring(next_index, need_cut_str.Length - next_index);
                        need_apply.Replace("\r", "");
                        List<string> list_apply = need_apply.Split('\n').ToList();

                        need_apply = list_apply[0].Substring(0, list_apply[0].Length - 1);
                        list_apply.RemoveAt(0);
                        list_apply.RemoveAt(list_apply.Count() - 1);

                        bool still_unit = false;
                        foreach (string part in list_apply)
                        {
                            char c = part[0];
                            if (still_unit)
                                need_apply += " ; " + part.Substring(0, part.Length - 1);
                            else
                            {
                                if ((c >= '0') || (c <= '9'))
                                    need_apply += " : " + part.Substring(0, part.Length - 1);
                                else
                                    need_apply += " ; " + part.Substring(0, part.Length - 1);
                            }
                            if (part.Contains('%'))
                                still_unit = true;
                            else
                                still_unit = false;
                        }
                        need_apply += System.Environment.NewLine;

                        list_all.Add(need_apply);

                        need_cut_str = need_cut_str.Remove(next_index, need_cut_str.Length - next_index);
                        next_index = need_cut_str.LastIndexOf('[');
                    }
                    list_all.Reverse();

                    list_all.ForEach(item => {
                        tb_Vaude_out.Text += item;
                    });
                }

            }
            else if (cb_sel_vaude.Text == "洗語")
            {
                need_cut_str = need_cut_str.Replace("\r","");
                need_cut_str = need_cut_str.Replace("\n","\n ");

                if (!need_cut_str.Contains("["))
                {
                    MessageBox.Show("非Vaude格式");
                    return;
                }
                int next_index = need_cut_str.LastIndexOf('[');
                need_cut_str = need_cut_str.Replace("[", "+[");

                need_cut_str = need_cut_str.Replace("\n", "");

                List<string> list_all = need_cut_str.Split('+').ToList();
                list_all.RemoveAt(0);
                list_all.ForEach(x => {
                    tb_Vaude_out.Text += x + System.Environment.NewLine + System.Environment.NewLine;
                });

                //List<string> care_code_list = list_all[0].Split('/').ToList();
                List<string> care_code_list = list_all[0].Split(' ').ToList();

                care_code_list.RemoveAt(0);

                List<string> care_number_str = new List<string>();
                string temp_str = "";
                for (int i = 0; i < care_code_list.Count(); i++)
                {
                    string str_this = care_code_list[i];
                    
                    if (str_this != "")
                    {
                        if (str_this.Contains(".") || str_this[0] == '/')
                        {
                            if (temp_str != "")
                            {
                                if (str_this[0] != '/')
                                    temp_str += " " + care_code_list[i];
                                temp_str = temp_str.Remove(0, 1);
                                care_number_str.Add(temp_str);
                                temp_str = "";
                            }
                        }
                        else {
                            temp_str += " " + care_code_list[i];
                        }
                    }
                }
                temp_str = temp_str.Remove(0, 1);
                care_number_str.Add(temp_str);
                temp_str = "";

                if (eng_list_v != null)
                {
                    care_number_str.ForEach(x => {
                        //tb_Vaude_out.Text += x + System.Environment.NewLine;
                        if (eng_list_v.Contains(x))
                            tb_care_num_v.Text += (eng_list_v.IndexOf(x) + 1).ToString() + " ";
                        else
                            tb_care_num_v.Text += "NotFound ";
                    });
                    if (care_number_str.Count() > 15)
                        MessageBox.Show("洗語標籤已超過15個");
                }
            }
        }

    }
}
