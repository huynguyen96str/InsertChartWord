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
using System.Diagnostics;
using Points = Microsoft.Office.Core.Points;

namespace InsertChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            pgb.Visible = false;
        }
        Microsoft.Office.Interop.Word.Application wordApplication;
        Microsoft.Office.Interop.Word.Document doc;
        object oMissing;
        private async void btnBrowser_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = @"Docx Files|*.docx;*.doc;";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                pgb.Visible = true;
                btnBrowser.Enabled = false;
                wordApplication = new Microsoft.Office.Interop.Word.Application();
                oMissing = System.Reflection.Missing.Value;
                Object oFalse = false;
                Object filename = (Object)(openFileDialog.FileName);

                doc = wordApplication.Documents.Open(ref filename, ref oMissing,
                    ref oFalse, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                doc.Activate();

                try
                {
                    for (int i = 1; i < doc.Tables.Count; i++)
                    {
                        SetTableComb(doc.Tables[i]);
                        pgb.Value = (int)(i * 100.0 / doc.Tables.Count);
                    }
                    MessageBox.Show("Finished!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    Release();
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                btnBrowser.Enabled = true;
            }
            pgb.Visible = false;
            await System.Threading.Tasks.Task.Delay(1);
        }
        private void Release()
        {
            //var pathfilename = @"C:\Users\PC\Desktop\Export\thuyetminh.pdf";
            //Object filename2 = (Object)pathfilename;
            //doc.SaveAs(ref filename2, WdSaveFormat.wdFormatPDF,
            //    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            //    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            //    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            // close word doc and word app.
            object saveChanges = WdSaveOptions.wdSaveChanges;
            ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
            ((_Application)wordApplication).Quit(ref oMissing, ref oMissing, ref oMissing);
            wordApplication = null;
            doc = null;
        }
        private void SetTableComb(Table tb)
        {
            if (!(tb.Columns.Count >= 4)) return;
            if (!(tb.Rows.Count > 1)) return;
            int rowstartIndex = 2;
            int rowsEndIndex = tb.Rows.Count - 1;
            // create Excel application object
            Range range = doc.Range(tb.Range.End, tb.Range.End);
            //Chart chart = doc.InlineShapes.AddChart(XlChartType.xlCombo, range).Chart;
            Chart chart = doc.InlineShapes.AddChart2(-1, XlChartType.xlCombo, range, Type.Missing).Chart;
            string title = tb.Cell(1, tb.Columns.Count - 2).Range.Text.Trim().Replace("\r", "").Replace("\a", "");
            chart.ChartTitle.Text = title;
            chart.HasTitle = title.Length > 0;
            int count = rowsEndIndex - rowstartIndex + 1;
            double[] seri1 = new double[count];
            double[] seri2 = new double[count];
            string[] lb = new string[count];
            int index = 0;
            string tt = "";
            for (int r = rowstartIndex; r <= rowsEndIndex; r++)
            {
                string v1 = tb.Cell(r, tb.Columns.Count - 1).Range.Text.Trim().Replace(",", ".").Replace("\r", "").Replace("\a", "");
                string v2 = tb.Cell(r, tb.Columns.Count).Range.Text.Trim().Replace(",", ".").Replace("\r", "").Replace("\a", "");
                seri1[index] = Convert.ToDouble(v1);
                seri2[index] = Convert.ToDouble(v2);
                tt = tb.Cell(r, tb.Columns.Count - 2).Range.Text.Trim().Replace("\r", "").Replace("\a", "");
                if (tt.Length > 60) tt = tt.Substring(0, 30) + "...";
                lb[index] = tt;
                index++;
            }
            if (lb.Count() <= 4)
            {
                ChartGroup cg1 = (ChartGroup)chart.ChartGroups(1);
                cg1.GapWidth = 500;
            }
            // Thêm các series vào biểu đồ
            Series series1 = (Series)chart.SeriesCollection(1);
            series1.Name = tb.Cell(1, tb.Columns.Count - 1).Range.Text;
            series1.ChartType = XlChartType.xlColumnClustered;
            series1.Values = seri1; // .value
            series1.AxisGroup = XlAxisGroup.xlPrimary;
            series1.ApplyDataLabels(Microsoft.Office.Interop.Word.XlDataLabelsType.xlDataLabelsShowValue);
            DataLabels dlb1 = (DataLabels)series1.DataLabels();
            series1.Format.Fill.ForeColor.RGB = (int)XlRgbColor.xlGreen;
            dlb1.Font.Color = ColorTranslator.ToOle(Color.White);
            dlb1.Position = Microsoft.Office.Interop.Word.XlDataLabelPosition.xlLabelPositionCenter;

            Series series2 = (Series)chart.SeriesCollection(2);
            series2.Name = tb.Cell(1, tb.Columns.Count).Range.Text;
            series2.ChartType = XlChartType.xlLineMarkers;
            series2.Values = seri2; //.value
            series2.AxisGroup = XlAxisGroup.xlSecondary;
            series2.ApplyDataLabels(Microsoft.Office.Interop.Word.XlDataLabelsType.xlDataLabelsShowValue);
            DataLabels dlb2 = (DataLabels)series2.DataLabels();
            series2.Format.Fill.ForeColor.RGB = (int)XlRgbColor.xlOrange;
            dlb2.Font.Color = ColorTranslator.ToOle(Color.Orange);
            dlb2.NumberFormat =  "General\\%"; ;
            dlb2.Position = Microsoft.Office.Interop.Word.XlDataLabelPosition.xlLabelPositionAbove;
            Series series3 = ((Series)chart.SeriesCollection(3));
            series3.Delete();

            chart.SeriesCollection(1).XValues = lb;

            // Đổi tên trục X và trục Y
            //Axis xAxis = (Axis)chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
            //xAxis.HasTitle = true;
            //xAxis.AxisTitle.Text = "Tháng";

            Axis yAxis1 = (Axis)chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
            yAxis1.HasTitle = true;
            yAxis1.AxisTitle.Text = tb.Cell(1, tb.Columns.Count - 1).Range.Text;

            Axis yAxis2 = (Axis)chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary);
            yAxis2.HasTitle = true;
            yAxis2.AxisTitle.Text = tb.Cell(1, tb.Columns.Count).Range.Text;
            chart.ChartData.Workbook.close();
        }
    }
}
