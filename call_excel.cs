
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Auto_Plan_Art_Plate
{
    class Call_Excel
    {
        public Excel.Application app_excel = null;
        public Excel._Workbook wb_excel = null;
        public Excel._Worksheet ws_excel = null;
        public bool _save_file = false;
        public string _fileName;
        SaveFileDialog saveFileDialog_fun;
        public void Open_File(string file)
        {
            _fileName = file;
            main();
        }
        public void Save_Exam_File(string exam_file, SaveFileDialog saveFile, string file_type = "xlsx")
        {
            _fileName = exam_file;
            _save_file = true;
            saveFileDialog_fun = saveFile;
            if (file_type == "xlsx")
                saveFileDialog_fun.Filter = "Excel Worksheets|*.xlsx";
            else
                saveFileDialog_fun.Filter = "Excel Worksheets|*.xls";
            main();
        }
        protected void main()
        {
            try
            {
                app_excel = new Excel.Application
                {
                    DisplayAlerts = false        //關閉excel執行對話框
                };
                try
                {
                    wb_excel = app_excel.Workbooks.Open(_fileName);
                    Load_File();
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.ToString());
                    app_excel.Workbooks.Close();
                    ws_excel = null;
                    MessageBox.Show("找不到範本檔案");
                }
                //儲存檔案
                if (_save_file)
                {
                    //saveFileDialog_fun.Filter = "Excel Worksheets|*.xlsx";
                    saveFileDialog_fun.ShowDialog();
                    if (saveFileDialog_fun.FileName != "")
                    {
                        wb_excel.SaveAs(saveFileDialog_fun.FileName);
                        MessageBox.Show("輸出資料OK", "提醒");
                    }
                    saveFileDialog_fun.FileName = "";      //避免再次使用對話框是上一筆檔名
                }
                ws_excel = null;
                wb_excel.Close();
                wb_excel = null;
                app_excel.Quit();
                app_excel = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public virtual void Load_File()
        {

        }
    }
}
