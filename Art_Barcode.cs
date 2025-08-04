using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using BarcodeStandard;
using Type = BarcodeStandard.Type;

namespace Auto_Plan_Art_Plate
{
    internal class Art_Barcode
    {
        string barcode_str;
        Illustrator.GroupItem dup_group;
        Illustrator.PathItem position_path;
        double Width;
        double Height;
        public Art_Barcode(string barcode_str, double Width, double Height, Illustrator.GroupItem dup_group, Illustrator.PathItem position_path)
        {
            this.barcode_str = barcode_str;
            this.dup_group = dup_group;
            this.position_path = position_path;
            this.Width = Width;
            this.Height = Height;
        }
        public void Print_in_art_barcode()
        {
            string barcode_data = barcode_str;

            Barcode barcode = new Barcode();
            barcode.Encode(Type.Code128B, barcode_data).Encode();
            //barcode.Encode(BarcodeLib.TYPE.CODE128, barcode_data);
            string bar_data = "";
            try
            {
                bar_data = barcode.EncodedValue;
            }
            catch (Exception)
            {
            }

            Illustrator.GroupItem bar_group = dup_group.GroupItems.Add();
            bar_group.Name = "barcode";

            bar_group.Top = position_path.Top;
            bar_group.Left = position_path.Left;

            double point_x = position_path.Left;
            //每個條碼單位的寬度
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

                    position_path.Duplicate(bar_group);
                    Illustrator.PathItem pathItem = bar_group.PathItems[1];

                    pathItem.Top = position_path.Top;
                    pathItem.Left = point_x;
                    pathItem.Height = 12;
                    //pathItem.StrokeWidth = line_width;
                    pathItem.Width = line_width;
                    point_x += point_width * poetd;
                }
                else
                {
                    point_width = 0.25;
                    point_x += point_width;
                }
            }
            bar_group.Height = Height * 2.835;
            bar_group.Width = Width * 2.835;
            position_path.Delete();
        }
    }
}
