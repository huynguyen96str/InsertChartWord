using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using XlAxisGroup = Microsoft.Office.Interop.Word.XlAxisGroup;
using XlAxisType = Microsoft.Office.Core.XlAxisType;

namespace InsertChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowser_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter  = @"Docx Files|*.docx;*.doc;";
            if(openFileDialog.ShowDialog()== DialogResult.OK)
            {
                string fn = openFileDialog.FileName;


                Microsoft.Office.Interop.Word.Application wordApplication = new Microsoft.Office.Interop.Word.Application();

                object oMissing = System.Reflection.Missing.Value;
                Object oFalse = false;
                Object filename = (Object)(fn);

                Microsoft.Office.Interop.Word.Document doc = wordApplication.Documents.Open(ref filename, ref oMissing,
                    ref oFalse, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                doc.Activate();
                // code Region
                // Bước 1: tạo table
                Table dataTable = doc.Tables.Add(doc.Range(), 4, 3);
                dataTable.Cell(1, 1).Range.Text = "Tháng";
                dataTable.Cell(1, 2).Range.Text = "Doanh thu";
                dataTable.Cell(1, 3).Range.Text = "Lợi nhuận";
                dataTable.Cell(2, 1).Range.Text = "Tháng 1";
                dataTable.Cell(2, 2).Range.Text = "100";
                dataTable.Cell(2, 3).Range.Text = "20";
                dataTable.Cell(3, 1).Range.Text = "Tháng 2";
                dataTable.Cell(3, 2).Range.Text = "150";
                dataTable.Cell(3, 3).Range.Text = "30";
                dataTable.Cell(4, 1).Range.Text = "Tháng 3";
                dataTable.Cell(4, 2).Range.Text = "200";
                dataTable.Cell(4, 3).Range.Text = "40";
                // https://blog.conholdate.com/total/create-charts-in-word-documents-using-csharp/
                // Tạo một Chart
                Range range = doc.Range(dataTable.Range.End, dataTable.Range.End);
                Chart chart = doc.InlineShapes.AddChart(XlChartType.xlCombo, range).Chart;

                
                // Thêm các series vào biểu đồ
                Series series1 = (Series)chart.SeriesCollection(1);
                series1.Name = "Doanh thu";
                series1.ChartType = XlChartType.xlColumnClustered;
                series1.Values = new double[] { 2.7, 3.2, 0.8 }; // .value
                series1.AxisGroup = XlAxisGroup.xlPrimary;

                Series series2 = (Series)chart.SeriesCollection(2);
                series2.Name = "Lợi nhuận";
                series2.ChartType = XlChartType.xlLineMarkers;
                series2.Values = new double[] { 1, 25, 5 }; //.value
                series2.AxisGroup = XlAxisGroup.xlSecondary;

                // Đổi tên trục X và trục Y
                Axis xAxis = (Axis)chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                xAxis.HasTitle = true;
                xAxis.AxisTitle.Text = "Tháng";

                Axis yAxis1 = (Axis)chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                yAxis1.HasTitle = true;
                yAxis1.AxisTitle.Text = "Doanh thu";

                Axis yAxis2 = (Axis)chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary);
                yAxis2.HasTitle = true;
                yAxis2.AxisTitle.Text = "Lợi nhuận";

                // Đổi màu và kiểu của series
                //series1.Format.Fill.ForeColor.RGB = (int)XlRgbColor.rgbRed;
                //series2.Format.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlue;

                var pathfilename = @"C:\Users\HUY NGUYEN\Desktop\export\thuyetminh.pdf";
                Object filename2 = (Object)pathfilename;

                doc.SaveAs(ref filename2, WdSaveFormat.wdFormatPDF,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);


                // close word doc and word app.
                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                ((_Application)wordApplication).Quit(ref oMissing, ref oMissing, ref oMissing);
                wordApplication = null;
                doc = null;
            }
        }
    }
}
