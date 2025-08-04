using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace Auto_Plan_Art_Plate
{
    internal class Ziener_refer_excel : Call_Excel
    {
        //各國的英文縮寫
        List<string> abridge_con_name = new List<string>();
        //不同類別的項目 <縮寫,各語言的值>
        Dictionary<string, string> dic_Shellfabric = new Dictionary<string, string>();
        Dictionary<string, string> dic_Lining = new Dictionary<string, string>();
        Dictionary<string, string> dic_Padding = new Dictionary<string, string>();
        Dictionary<string, string> dic_Suspenders = new Dictionary<string, string>();
        //蒐集到的所有成分名稱 <縮寫,所有該語言的成分名>
        Dictionary<string, List<string>> dic_contect = new Dictionary<string, List<string>>();

        DataTable dt = new DataTable();
        List<string> list_acronym = new List<string>();
        public Ziener_refer_excel()
        {

        }

        public override void Load_File()
        {
            ws_excel = wb_excel.Worksheets["表裡布"];

            int name_col_num = 4;
            int end_col = 0;
            while (ws_excel.Cells[1, name_col_num].text != "不使用")
            {
                string con_name = ws_excel.Cells[1, name_col_num].text;

                con_name = con_name.Remove(0, con_name.IndexOf(' ') + 1);
                string[] sp_c_name = con_name.Split('\n');
                abridge_con_name.Add(sp_c_name[0]);
                name_col_num++;
            }
            end_col = name_col_num - 1;

            set_dic_cell(2, end_col, ref dic_Shellfabric);
            set_dic_cell(3, end_col, ref dic_Lining);
            set_dic_cell(4, end_col, ref dic_Padding);
            set_dic_cell(5, end_col, ref dic_Suspenders);

            ws_excel = wb_excel.Worksheets["成份"];

            for (int i = 4; i <= end_col; i++)
            {
                int row_count = 2;
                List<string> list = new List<string>();
                while (ws_excel.Cells[row_count, i].text != "")
                {
                    list.Add(ws_excel.Cells[row_count, i].text);
                    row_count++;
                }
                dic_contect.Add(abridge_con_name[i - 4], list);
            }

            //縮寫儲存
            int col_acronym = 1;
            while (ws_excel.Cells[1, col_acronym].text != "縮寫")
            {
                col_acronym++;
            }
            //col_acronym--;
            int per_cont = 2;
            while (ws_excel.Cells[per_cont, col_acronym].text != "")
            { 
                string acronym = ws_excel.Cells[per_cont, col_acronym].text;
                list_acronym.Add(acronym);
                per_cont++;
            }

            ws_excel = wb_excel.Worksheets["洗語"];


            dt.Columns.Add("no");
            foreach (string str in abridge_con_name)
            {
                dt.Columns.Add(str);
            }
            int row_end = 6;
            
            while (ws_excel.Cells[row_end, 6].text != "")
            {
                int col_init = 6;
                DataRow dr = dt.NewRow();
                dr["no"] = row_end - 5;
                foreach (string lang_str in abridge_con_name)
                {
                    dr[lang_str] = ws_excel.Cells[row_end, col_init].text;
                    col_init++;
                }
                dt.Rows.Add(dr);
                row_end++;
            }
            
        }
        public void set_dic_cell(int row_num, int end_col, ref Dictionary<string, string> dic_fun)
        {
            //int row_num = 2;
            for (int i = 4; i <= end_col; i++)
            {
                dic_fun.Add(abridge_con_name[i - 4], ws_excel.Cells[row_num, i].text);
            }
        }
        public List<string> get_list()
        {
            return abridge_con_name;
        }
        public Dictionary<string, string> get_dic_title(int sw)
        {
            switch (sw)
            {
                case 1:
                    return dic_Shellfabric;
                case 2:
                    return dic_Lining;
                case 3:
                    return dic_Padding;
                case 4:
                    return dic_Suspenders;
            }
            return null;
        }
        public Dictionary<string, List<string>> get_dic_list()
        {
            return dic_contect;
        }
        public DataTable get_dt()
        {
            return dt;
        }
        public List<string> Get_list_acronym()
        {
            return list_acronym;
        }
    }

    class Ziener_exam_excel : Call_Excel
    {
        List<string> list_language;
        Dictionary<string, string> dic_Shellfabric;
        Dictionary<string, string> dic_Lining;
        Dictionary<string, string> dic_Padding;
        Dictionary<string, string> dic_Suspenders;
        List<Dictionary<string, string>> dic_list_title;
        //key 語言縮寫 value 
        Dictionary<string, List<string>> dic_c;
        string complete_str = "";
        public Ziener_exam_excel(List<string> lang_list,List<Dictionary<string, string>> list_dic, Dictionary<string, List<string>> item1)
        {
            list_language = lang_list;
            dic_Shellfabric = list_dic[0];
            dic_Lining = list_dic[1];
            dic_Padding = list_dic[2];
            dic_Suspenders = list_dic[3];
            dic_c = item1;
            dic_list_title = list_dic;
        }

        public override void Load_File()
        {
            int sheet_count = wb_excel.Worksheets.Count;

            //多筆資料表做迴圈執行
            for (int i = 1; i <= sheet_count; i++)
            { 
                ws_excel = wb_excel.Worksheets[i];
                complete_str = "";
                List<string> list_output = new List<string>();
                //
                int row_index = 1;
                while (ws_excel.Cells[row_index, 3].text != "")
                {
                    string contect_text = ws_excel.Cells[row_index, 3].text;
                    if (contect_text.Length > 3)
                    {
                        //處理title的序號
                        string title_con = ws_excel.Cells[row_index, 1].text;
                        list_output.Add(title_con + "+" + contect_text);
                    }
                    row_index++;
                }

                foreach (string lang_str in list_language)
                {
                    complete_str += lang_str + " ";
                    foreach (string value in list_output)
                    {
                        //處理title的序號
                        string[] temp = value.Split('+');
                        string title = temp[0];
                        if (title.Contains(" "))
                        {
                            string[] tmp = title.Split(' ');
                            string title_lang = tmp[0].ToUpper();
                            title_lang = Get_title_dic_lan(title_lang, lang_str);
                            int tmp_num = Convert.ToInt32(tmp[1]);
                            if (tmp_num > 1)
                            {
                                complete_str += ", " + tmp_num + ". ";
                            }
                            else
                            {
                                if (complete_str.Count() > 5 && list_output.IndexOf(value) > 1)
                                    complete_str += " ; " + title_lang + " : " + tmp[1] + ". ";
                                else
                                    complete_str += title_lang + " : " + tmp[1] + ". ";
                            }
                        }
                        //處理趴數成份
                        string contect = temp[1];
                        List<string> contect_list = dic_c["ENG"];
                        int list_index;
                        if (contect_list.Contains(contect.ToUpper()))
                        {
                            list_index = contect_list.IndexOf(contect.ToUpper());
                            complete_str += dic_c[lang_str][list_index];
                        }
                        else
                        {
                            if (contect.Contains(" "))
                            {
                                contect = contect.Replace(" ", "");
                            }
                            if (contect.Contains(","))
                            {
                                List<string> cell_temp = contect.Split(',').ToList();

                                string cell_1 = cell_temp[0];
                                Set_str_space(ref cell_1);

                                string num = cell_1.Split(' ')[0];
                                string cell_1_con = cell_1.Split(' ')[1].ToUpper();
                                list_index = contect_list.IndexOf(cell_1_con);

                                complete_str += num + " " + dic_c[lang_str][list_index] + " / ";
                                //complete_str += cell_1.ToUpper() + " / ";
                                //第二個成分
                                string cell_2 = cell_temp[1];
                                Set_str_space(ref cell_2);

                                string cell_2_con = cell_2.Split(' ')[1].ToUpper();
                                list_index = contect_list.IndexOf(cell_2_con);

                                complete_str += cell_2.Split(' ')[0] + " " + dic_c[lang_str][list_index];
                                //complete_str += cell_2.ToUpper();
                            }
                            else
                            {
                                Set_str_space(ref contect);
                                string contect_sp = contect.Split(' ')[1].ToUpper();
                                list_index = contect_list.IndexOf(contect_sp);

                                complete_str += contect.Split(' ')[0] + " " + dic_c[lang_str][list_index];
                                //complete_str += contect.ToUpper();
                            }
                        }
                    }
                    complete_str += "\r";
                }
                ws_excel.Cells[20, 1].value = complete_str;
            }
        }

        public void Set_str_space(ref string str_fun)
        {
            if (!str_fun.Contains(' '))
                str_fun = str_fun.Insert(str_fun.IndexOf('%') + 1, " ");
        }
        public string Get_complate_str()
        { 
            return complete_str;
        }

        public string Get_title_dic_lan(string title_lang,string at_language)
        {
            Dictionary<string, string> dic_fun = new Dictionary<string, string>();
            foreach (Dictionary<string, string> dic in dic_list_title)
            {
                string lang_str = "ENG";
                if (dic[lang_str] == title_lang)
                { 
                    dic_fun = dic;
                    break;
                }
            }

            return dic_fun[at_language];
        }
    }

    class Ziener_acronym : Call_Excel
    {
        List<string> list_language;
        Dictionary<string, string> dic_Shellfabric;
        Dictionary<string, string> dic_Lining;
        Dictionary<string, string> dic_Padding;
        Dictionary<string, string> dic_Suspenders;
        List<Dictionary<string, string>> dic_list_title;
        //key 語言縮寫 value 
        Dictionary<string, List<string>> dic_c;
        string complete_str = "";
        List<string> list_acronym;
        public Ziener_acronym(List<string> lang_list, List<Dictionary<string, string>> list_dic, Dictionary<string, List<string>> item1, List<string> list_a)
        {
            list_language = lang_list;
            dic_Shellfabric = list_dic[0];
            dic_Lining = list_dic[1];
            dic_Padding = list_dic[2];
            dic_Suspenders = list_dic[3];
            dic_c = item1;
            dic_list_title = list_dic;
            this.list_acronym = list_a;
        }

        public override void Load_File()
        {
            ws_excel = wb_excel.Worksheets[1];
            List<int> row_init_list = new List<int>();

            int row_inde = 2;

            while (ws_excel.Cells[row_inde,1].text != "end")
            { 
                if (ws_excel.Cells[row_inde, 1].text != "")
                    row_init_list.Add(row_inde);
                row_inde++;
            }
            //取得核對序列
            List<string> eng_check = dic_c["ENG"];

            int col_out = 7;

            foreach (int row_each in row_init_list)
            { 
                //每列逐行處理
                List<string> list_row_per = new List<string>();
                string complete_str = "";
                int col_target = 1;
                string shell = ws_excel.Cells[row_each, col_target].text;
                //SHELLFABRIC
                //LINING
                //PADDING
                complete_str += dic_list_title[col_target - 1]["ENG"] + " : ";

                shell = shell.Trim();
                shell = shell.Replace("\r", "");
                shell = shell.Replace("\n", "");

                if (shell.Contains("("))
                {
                    if (list_acronym.Contains(shell))
                    {
                        complete_str += shell;
                    }
                    else {
                        complete_str += shell.Split(' ')[0].Trim() + " ";
                        string temp_con = shell.Split('%')[1].Trim();
                        if (list_acronym.Contains(temp_con))
                            complete_str += eng_check[list_acronym.IndexOf(temp_con)];
                        else
                            complete_str += "NotFound";
                    }
                }
                else if (shell.Contains("2."))
                {
                    shell = shell.Replace("2.", "+2.");
                    List<string> split_shell = shell.Split('+').ToList();
                    string ps_str = "";
                    
                    foreach (string split_part in split_shell)
                    {
                        if (ps_str != "")
                            ps_str += ", ";

                        if (split_part.Contains(","))
                        {
                            string[] split_1 = split_part.Trim().Split(',');

                            split_1[0] = split_1[0].Replace(" ", "");
                            split_1[0] = split_1[0].Replace(".",". ");
                            //趴數
                            ps_str += split_1[0].Split('%')[0] + "% ";
                            ps_str += eng_check[list_acronym.IndexOf(split_1[0].Split('%')[1])];

                            split_1[1] = split_1[1].Replace(" ", "");
                            split_1[1] = split_1[1].Replace(".", ". ");
                            ps_str += " / ";
                            split_1[1] = split_1[1].Trim();
                            ps_str += split_1[1].Split('%')[0] + "% ";
                            ps_str += eng_check[list_acronym.IndexOf(split_1[1].Split('%')[1])];
                        }
                        else {
                            string num_con;
                            string check_acronym;
                            string second_part = split_part.Trim();
                            if (second_part.Contains(" "))
                            {
                                num_con = split_part.Split(' ')[0];
                                check_acronym = split_part.Split(' ')[1].Trim();
                            }
                            else
                            {
                                num_con = split_part.Split('%')[0] + "%";
                                check_acronym = split_part.Split('%')[1];
                                check_acronym = check_acronym.Trim();
                            }

                            if (list_acronym.Contains(check_acronym))
                                ps_str += num_con + " " + eng_check[list_acronym.IndexOf(check_acronym)];
                            else
                                ps_str += " NotFound";
                        }
                    }
                    complete_str += ps_str;
                }
                else if (shell.Contains(","))
                {
                    string[] split_1 = shell.Split(',');

                    //趴數
                    complete_str += split_1[0].Split(' ')[0] + " ";
                    complete_str += eng_check[list_acronym.IndexOf(split_1[0].Split(' ')[1])];

                    complete_str += " / ";
                    split_1[1] = split_1[1].Trim();
                    complete_str += split_1[1].Split(' ')[0] + " ";
                    complete_str += eng_check[list_acronym.IndexOf(split_1[1].Split(' ')[1])];
                }
                else {
                    string num_con = shell.Split(' ')[0];
                    string check_acronym = shell.Split(' ')[1];

                    if (list_acronym.Contains(check_acronym))
                        complete_str += num_con + " " + eng_check[list_acronym.IndexOf(check_acronym)];
                    else
                        complete_str += " NotFound";
                }
                list_row_per.Add(shell);


                ws_excel.Cells[row_each, col_out].value = complete_str;
            }
        }

        public string Get_comp()
        { 
            return complete_str;
        }
    }
}
