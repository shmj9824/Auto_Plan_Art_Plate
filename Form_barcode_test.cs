using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using BarcodeStandard;
//using BarcodeLib;
//using Type = BarcodeStandard.Type;

namespace Auto_Plan_Art_Plate
{
    public partial class Form_barcode_test : Form
    {
        public Form_barcode_test()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Barcode barcode = new Barcode();
            int tail = Convert.ToInt32(tb_tail.Text);
            int width = Convert.ToInt32(tb_width.Text);

            /*
            Image img_bar = Image.FromStream(barcode.Encode(Type.Code128B, "123456789ABCDEFG", width, tail).Encode().AsStream());
            pictureBox1.Image = img_bar;
            Span<byte> data = barcode.Encode(Type.Code128A, "BAB43342", width, tail).Encode().Span;
            SKData sKData = barcode.Encode(Type.Code128B, "123456789ABCDEFG", width, tail).EncodedData;

            textBox1.Text = barcode.EncodedValue;
            */
            /*
            Image img_bar = Image.FromStream(barcode.Encode(Type.Code128B, "pppppp", width, tail).Encode().AsStream());
            pictureBox1.Image = img_bar;
            textBox1.Text = barcode.EncodedValue;

            if (img_bar != null)
                img_bar.Save(System.Environment.CurrentDirectory + "\\barcode.bmp", System.Drawing.Imaging.ImageFormat.Bmp);
            */
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //new Art_VP_File(null, null, null, saveFileDialog1).Test_barcode_ai();
            new Art_VP_File(null, null, saveFileDialog1).test_print_barcode();
            /*
            Illustrator.Application app = new Illustrator.Application();
            int i = 0;
            foreach (string aatp in app.TracingPresetsList)
            {
                textBox1.Text += i++ + ":" + aatp + ",";
            }
            */
            //Illustrator.Application app = new Illustrator.Application();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            new Art_VP_File(null, null, saveFileDialog1).create_barcode_fun_test();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            new Art_VP_File(null, null, saveFileDialog1).Control_artboard();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            new Art_VP_File(null, null, saveFileDialog1).Test_barcode_ai();
        }
    }
}
