using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.Axes;
using OxyPlot.WindowsForms;
using System.Reflection;
using ExcelObj = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Demonstration
{
    public partial class Form1 : Form
    {
        PlotView Plot = new PlotView();

        public Form1()
        {
            InitializeComponent();
            Controls.Add(Plot);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            int x, y, w, z;
            System.Threading.ThreadPool.GetMinThreads(out y, out x);
            System.Threading.ThreadPool.SetMinThreads(8000, x);
            System.Threading.ThreadPool.GetMaxThreads(out w, out z);
            System.Threading.ThreadPool.SetMaxThreads(16000, z);

            Dictionary<double, double> dataExcel = new Dictionary<double, double>();
            string directory = "D:\\Downloads\\Practice2017\\data1.xlsx";

            Maket maket = new Maket(new ExcelBinding(dataExcel, directory), new GraphicsOxiPlot(dataExcel));

            maket.Output();
            //Task ts = new Task(() => maket.Output());
            //ts.Start();
            //ts.Wait();
        }

        interface IDataBinding
        {
            void Binding();
        }

        class ExcelBinding : IDataBinding
        {
            private Dictionary<double, double> _dic;
            private string _dir;

            public ExcelBinding(Dictionary<double, double> dic, string dir)
            {
                _dic = dic;
                _dir = dir;
            }

            public void Binding()
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.DefaultExt = "*.xls;*.xlsx";
                ofd.Filter = "Excel Sheet(*.xlsx)|*.xlsx";
                ofd.Title = "Выберите документ для загрузки данных";
                ExcelObj.Application app = new ExcelObj.Application();
                ExcelObj.Workbook workbook;
                ExcelObj.Worksheet NwSheet;
                ExcelObj.Range ShtRange;
                DataTable dt = new DataTable();

                
                ofd.FileName = _dir;
                //if(ofd.ShowDialog() == DialogResult.OK)

                workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value);

                NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                ShtRange = NwSheet.UsedRange;
                for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                {
                    dt.Columns.Add(
                       new DataColumn((ShtRange.Cells[1, Cnum] as ExcelObj.Range).Value2.ToString()));
                }
                dt.AcceptChanges();

                string[] columnNames = new String[dt.Columns.Count];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    columnNames[0] = dt.Columns[i].ColumnName;
                }

                for (int Rnum = 2; Rnum <= ShtRange.Rows.Count; Rnum++)
                {
                    DataRow dr = dt.NewRow();
                    for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                    {
                        if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                        {
                            dr[Cnum - 1] =
                (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }

                for (int m = 0; m < dt.Rows.Count; m++)
                    _dic.Add(Convert.ToDouble(dt.Rows[m][0]), Convert.ToDouble(dt.Rows[m][1]));

                app.Quit();
            }
        }

        interface IGraphics
        {
            void OutputGraphics();
        }

        class GraphicsOxiPlot : Form1, IGraphics
        {
            private LineSeries s1;
            private LineSeries s2;
            private LineSeries s3;

            //values from binding (P(t))
            private Dictionary<double, double> _pt;

            //values of trend
            private Dictionary<double, double> _ft;

            //values stationary process
            private Dictionary<double, double> _et;

            //estimated values (E^(t))
            private Dictionary<double, double> _eet;

            public GraphicsOxiPlot(Dictionary<double, double> dic)
            {
                _pt = dic;
            }

            public void OutputGraphics()
            {
                //count measurement
                int k = 5;

                //critical value
                double N = -0.01;

                //max measurement error
                double b = 0.025;

                //start and finish sequence
                int ts = 1;
                int tf = k;

                //--------------------- Building ----------------------------
                
                Plot.Model = new PlotModel { Title = "plot3 Score" }; 
                Plot.Dock = DockStyle.Fill;
                

                Plot.Model.PlotType = PlotType.XY;
                Plot.Model.Background = OxyColor.FromRgb(255, 255, 255);
                Plot.Model.TextColor = OxyColor.FromRgb(0, 0, 0);

                s1 = new LineSeries { Title = "P(t)", StrokeThickness = 1 };
                s1.Color = OxyColor.FromRgb(255, 0, 0);//red

                //foreach (var item in _pt)
                //{
                //    s1.Points.Add(new DataPoint(item.Key, item.Value));
                //}

                Plot.Model.Series.Add(s1);

                s2 = new LineSeries { Title = "f(t)", StrokeThickness = 1 };
                s2.Color = OxyColor.FromRgb(0, 255, 0);//green

                s2.Points.Add(new DataPoint(42796.4, 50.2));
                s2.Points.Add(new DataPoint(42796.5, 51.00));
                s2.Points.Add(new DataPoint(42797.6, 51.400));

                //foreach (var item in _ft)
                //{
                //    s1.Points.Add(new DataPoint(item.Key, item.Value));
                //}

                Plot.Model.Series.Add(s2);

                s3 = new LineSeries { Title = "E(t)", StrokeThickness = 1 };
                s3.Color = OxyColor.FromRgb(0, 0, 255);//blue

                s3.Points.Add(new DataPoint(42796.4, 50.2));
                s3.Points.Add(new DataPoint(42796.5, 50.30));
                s3.Points.Add(new DataPoint(42797.6, 50.500));

                //foreach (var item in _et)
                //{
                //    s1.Points.Add(new DataPoint(item.Key, item.Value));
                //}

                Plot.Model.Series.Add(s3);
            }
        }

        class Maket
        {
            private IDataBinding _db;
            private IGraphics _gr;

            public Maket(IDataBinding db, IGraphics gr)
            {
                _db = db;
                _gr = gr;
            }

            public void Output()
            {
                _db.Binding();
                _gr.OutputGraphics();
            }
        }
    }
}