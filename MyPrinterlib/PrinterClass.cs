using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using MyOffice;
namespace MyPrinterlib
{
    public class PrinterClass
    {
        public enum PageSize
        {
            A4,
            A4_reverse
        }

        public delegate void PrintPageEventHandler(object sender, Graphics g, int width, int height, int page_num);
        public event PrintPageEventHandler PrintPageEvent;
        private List<SheetClass> SheetClasses = new List<SheetClass>();

        private PrintDialog printDialog = new PrintDialog();
        private PrintDocument printDocument = new PrintDocument();
        private PrintPreviewDialog printPreviewDialog = new PrintPreviewDialog();
        private PageSetupDialog pageSetupDialog = new PageSetupDialog();
        public int numOfPage = 1;
        public int numOfPage_buf = 0;
        public void Init()
        {
            printDocument.PrintPage += PrintDocument_PrintPage;
            printDocument.BeginPrint += PrintDocument_BeginPrint;
        }

        public DialogResult ShowPrintDialog(List<SheetClass> sheetClasses, PageSize pageSize)
        {
            this.SheetClasses = sheetClasses;
            return this.ShowPrintDialog(pageSize, sheetClasses.Count);
        }
        public DialogResult ShowPrintDialog(PageSize pageSize)
        {
            return this.ShowPrintDialog(pageSize, 1);
        }
        public DialogResult ShowPrintDialog(PageSize pageSize, int numOfPage)
        {
            this.numOfPage = numOfPage;
            if (pageSize == PageSize.A4)
            {
                printDocument.DefaultPageSettings.PaperSize = new PaperSize("", 827, 1169);
            }
            else if (pageSize == PageSize.A4_reverse)
            {
                printDocument.DefaultPageSettings.PaperSize = new PaperSize("", 1169, 827);
            }
            printDialog.Document = printDocument;
            return printDialog.ShowDialog();
        }
        public DialogResult ShowPreviewDialog(List<SheetClass> sheetClasses, PageSize pageSize)
        {
            this.SheetClasses = sheetClasses;
            return this.ShowPreviewDialog(pageSize, sheetClasses.Count);
        }
        public DialogResult ShowPreviewDialog(PageSize pageSize)
        {
            return this.ShowPreviewDialog(pageSize, 1);
        }
        public DialogResult ShowPreviewDialog(PageSize pageSize , int numOfPage)
        {
            this.numOfPage = numOfPage;
            printPreviewDialog.ShowIcon = false;
            ((Form)printPreviewDialog).StartPosition = FormStartPosition.CenterScreen;
            printPreviewDialog.DesktopLocation = new System.Drawing.Point(0, 0);
            printPreviewDialog.Width = 1024;
            printPreviewDialog.Height = 768;

            if (pageSize == PageSize.A4)
            {
                printDocument.DefaultPageSettings.PaperSize = new PaperSize("", 827, 1169);
            }
            else if (pageSize == PageSize.A4_reverse)
            {
                printDocument.DefaultPageSettings.PaperSize = new PaperSize("", 1169, 827);
            }

            printPreviewDialog.Document = printDocument;

            return printPreviewDialog.ShowDialog();
        }
        public DialogResult ShowPageSetupDialog(PageSize pageSize)
        {
            return ShowPageSetupDialog(pageSize, 1);
        }
        public DialogResult ShowPageSetupDialog(PageSize pageSize , int numOfPage)
        {
            this.numOfPage = numOfPage;
            if (pageSize == PageSize.A4)
            {
                printDocument.DefaultPageSettings.PaperSize = new PaperSize("", 827, 1169);
            }
            else if (pageSize == PageSize.A4_reverse)
            {
                printDocument.DefaultPageSettings.PaperSize = new PaperSize("", 1169, 827);
            }
            pageSetupDialog.Document = printDocument;
            return pageSetupDialog.ShowDialog();
        }
        public void Print(PageSize pageSize)
        {
            Print(pageSize, 1);
        }
        public void Print(PageSize pageSize, int numOfPage)
        {
            this.numOfPage = numOfPage;
            if (pageSize == PageSize.A4)
            {
                printDocument.DefaultPageSettings.PaperSize = new PaperSize("", 827, 1169);
            }
            else if (pageSize == PageSize.A4_reverse)
            {
                printDocument.DefaultPageSettings.PaperSize = new PaperSize("", 1169, 827);
            }
            printDocument.Print();
        }
        public SheetClass GetSheetClass(int pageNum)
        {
            return this.SheetClasses[pageNum - 1];
        }

        #region Event
        private void PrintDocument_BeginPrint(object sender, PrintEventArgs e)
        {
            numOfPage_buf = 0;
        }
        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {        
            numOfPage_buf++;
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.HighQuality; //使繪圖質量最高，即消除鋸齒
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.CompositingQuality = CompositingQuality.HighQuality;
            g.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;

            int canvas_width = e.PageBounds.Width;
            int canvas_height = e.PageBounds.Height;
            string 頁碼 = $"{numOfPage_buf}/{numOfPage}";
            SizeF size_頁碼 = g.MeasureString(頁碼,new Font("微軟正黑體",12, FontStyle.Bold), new Size(canvas_width, canvas_height), StringFormat.GenericDefault);
            PointF point_頁碼 = new PointF((canvas_width - size_頁碼.Width) / 2, (canvas_height - size_頁碼.Height - 5));
            g.DrawString(頁碼, new Font("微軟正黑體", 12, FontStyle.Bold), new SolidBrush(Color.Black), point_頁碼.X, point_頁碼.Y);


            if (PrintPageEvent != null)
            {
                PrintPageEvent(sender, g, canvas_width, canvas_height, numOfPage_buf);
            }
            if (numOfPage_buf < numOfPage)
            {
                e.HasMorePages = true;
            }
        }
        #endregion
    }
}
