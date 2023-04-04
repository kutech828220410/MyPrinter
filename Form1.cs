using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MyOffice;
using Basic;
namespace MyPrinter
{
    public partial class Form1 : Form
    {
        MyPrinterlib.PrinterClass printerClass = new MyPrinterlib.PrinterClass();

        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            printerClass.Init();
            printerClass.PrintPageEvent += PrinterClass_PrintPageEvent;
        }

      

        private void button1_Click(object sender, EventArgs e)
        {

            string json = Basic.Net.WEBApiGet("http://127.0.0.1:444/api/test/excel");
            List<SheetClass> sheetClass = json.JsonDeserializet<List<SheetClass>>();
            if (printerClass.ShowPreviewDialog(sheetClass , MyPrinterlib.PrinterClass.PageSize.A4) == DialogResult.OK)
            {
          
            }
            sheetClass.NPOI_SaveFile(@"C:\Users\User\Desktop\123.xls");
        }
        private void PrinterClass_PrintPageEvent(object sender, Graphics g, int width, int height, int page_num)
        {
          
            //for (int i = 0; i < 50; i++)
            //{
            //    sheetClass.AddNewCell_Webapi(i + 3, i + 3, 0, 0, "123", "微軟正黑體", 14, false, NPOI_Color.BLACK, 430, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.None);
            //}

            Rectangle rectangle = new Rectangle(0, 0, width, height);
            using (Bitmap bitmap = printerClass.GetSheetClass(page_num).GetBitmap(width,height,0.7, H_Alignment.Center, V_Alignment.Top, 0, 50))
            {
                rectangle.Height = bitmap.Height;
                g.DrawImage(bitmap, rectangle);
            }
        }

    }
}
