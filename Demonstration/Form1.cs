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
using System.Threading;

namespace Demonstration
{
    public partial class Form1 : Form
    {
        static PlotView Plot = new PlotView();

        public Form1()
        {
            InitializeComponent();
            Controls.Add(Plot);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            int x, y, w, z;
            ThreadPool.GetMinThreads(out y, out x);
            ThreadPool.SetMinThreads(100000, x);
            ThreadPool.GetMaxThreads(out w, out z);
            ThreadPool.SetMaxThreads(300000, z);

            Dictionary<double, double> dataExcel = new Dictionary<double, double>();
            string directory = "D:\\Downloads\\Practice2017\\data2.xlsx";

            Maket maket = new Maket(new ExcelBinding(dataExcel, directory), new GraphicsOxiPlot(dataExcel));

            maket.Output();
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
                //count measurement
                int k = 20;

                //critical value
                double N = 0.02;

                //max measurement error
                double b = 0.025;

                double sum = 0;
                double max = 0;
                double min = 0;
                double[] average = new double[0];
                double averageA = 0;
                double[] x = new double[k];
                double[] y = new double[1];
                y[y.Length - 1] = 0;

                double estKsi = 0.0;
                double ksi = 0.0;

                //upper and lower control limits
                double[] upper = new double[1];
                upper[0] = 0;
                double[] lower = new double[1];
                lower[0] = 0;
                double sigma = 0.0;
                double slack = 0.0;
                double target = 0.0;
                double[] r = new double[0];
                double averageR = 0.0;
                double[] cum = new double[1];
                cum[0] = 0;
                double[] s = new double[0];
                double allsum = 0;
                double[] cumPlus = new double[1];
                cumPlus[0] = 0;
                double[] cumMinus = new double[1];
                cumMinus[0] = 0;

                //v-mask
                double h = 0.0;
                double f = 0;
                double sigmaE = 0.0;
                double H = 0;
                double F = 0;

                //valuation standart deriviation
                double sigma0 = 0.0;

                Plot.Model = new PlotModel();
                Plot.Dock = DockStyle.Fill;

                Plot.Model.PlotType = PlotType.XY;
                Plot.Model.Background = OxyColor.FromRgb(255, 255, 255);
                Plot.Model.TextColor = OxyColor.FromRgb(0, 0, 0);

                GraphicsOxiPlot.s1 = new LineSeries { Title = "P(t)", StrokeThickness = 1 };
                GraphicsOxiPlot.s1.Color = OxyColor.FromRgb(255, 0, 0);//red
                Plot.Model.Series.Add(GraphicsOxiPlot.s1);

                GraphicsOxiPlot.s2 = new LineSeries { Title = "f(t)", StrokeThickness = 1 };
                GraphicsOxiPlot.s2.Color = OxyColor.FromRgb(0, 255, 0);//green
                Plot.Model.Series.Add(GraphicsOxiPlot.s2);

                GraphicsOxiPlot.s3 = new LineSeries { Title = "", StrokeThickness = 1 };
                GraphicsOxiPlot.s3.Color = OxyColor.FromRgb(0, 0, 255);//blue
                Plot.Model.Series.Add(GraphicsOxiPlot.s3);

                GraphicsOxiPlot.s4 = new LineSeries { Title = "", StrokeThickness = 1 };
                GraphicsOxiPlot.s4.Color = OxyColor.FromRgb(255, 125, 255);//pink
                Plot.Model.Series.Add(GraphicsOxiPlot.s4);

                GraphicsOxiPlot.s5 = new LineSeries { Title = "", StrokeThickness = 1 };
                GraphicsOxiPlot.s5.Color = OxyColor.FromRgb(200, 125, 255);//?
                Plot.Model.Series.Add(GraphicsOxiPlot.s5);

                GraphicsOxiPlot.s6 = new LineSeries { Title = "", StrokeThickness = 1 };
                GraphicsOxiPlot.s6.Color = OxyColor.FromRgb(255, 125, 200);//?
                Plot.Model.Series.Add(GraphicsOxiPlot.s6);

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

                Task task = new Task(() =>
                {
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

                    double temp1 = 0;
                    double temp2 = 0;

                    for (int Rnum = 2; Rnum <= ShtRange.Rows.Count; Rnum++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                        {
                            if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                            {
                                dr[Cnum - 1] = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();

                                if (temp1 == 0)
                                    temp1 = Convert.ToDouble(dr[Cnum - 1]);
                                else
                                {
                                    temp2 = Convert.ToDouble(dr[Cnum - 1]);
                                }
                            }
                        }
                        dt.Rows.Add(dr);
                        dt.AcceptChanges();

                        Plot.InvalidatePlot(true);
                        GraphicsOxiPlot.s1.Points.Add(new DataPoint(temp1, temp2));
                        _dic.Add(temp1, temp2);

                        //at first X = {P(ts), P(ts+1), ... P(tf)}
                        if (_dic.Count == k)
                        {
                            int ind = 0;
                            foreach (var item in _dic)
                            {
                                if (ind < x.Length)
                                {
                                    x[ind] = _dic[item.Key];
                                    ind++;
                                }
                            }
                        }

                        //trend, CUSUM
                        if (_dic.Count > k && Rnum % 1 == 0)
                        {
                            //average(TREND) from set X                       
                            Array.Resize(ref average, average.Length + 1);
                            average[average.Length - 1] = x.Average();

                            //average average
                            averageA = average.Average();

                            //target
                            target = averageA;

                            //v-mask
                            H = h * sigmaE;
                            F = f * sigmaE;

                            allsum = 0;
                            //sum of all _dic
                            foreach (var item in _dic)
                                allsum += item.Value - target;

                            //partial sums of averages
                            Array.Resize(ref s, s.Length + 1);
                            s[s.Length - 1] = allsum;

                            //cum
                            Array.Resize(ref cum, cum.Length + 1);
                            cum[cum.Length - 1] = s[s.Length - 1];

                            //range
                            max = x.Max();
                            min = x.Min();
                            Array.Resize(ref r, r.Length + 1);
                            r[r.Length - 1] = max - min;

                            //average range
                            averageR = r.Average();
                            sigma0 = averageR / 1.128;
                            sigmaE = sigma0 / Math.Sqrt(x.Length);

                            //remove trend E^(tf + 1) = P(tf + 1) - f(tf + 1)                            
                            estKsi = _dic[temp1] - average[average.Length - 1];
                            
                            //find E[tf + 1] = E^[tf + 1] + b
                            ksi = estKsi + b;

                            //CUSUM -- y[tf + 1] = y[tf] + ksi
                            Array.Resize(ref y, y.Length + 1);
                            y[y.Length - 1] = y[y.Length - 2] + ksi;

                            //standart deviation for limits
                            sigma = average[average.Length - 1] / 1.128;
                            slack = 0.5 * sigma;
                            h = 4 * sigma;

                            //cumPlus
                            Array.Resize(ref cumPlus, cumPlus.Length + 1);
                            cumPlus[cumPlus.Length - 1] = Math.Max(0, temp2 - target - slack + cumPlus[cumPlus.Length - 2]);

                            //cumMinus
                            Array.Resize(ref cumMinus, cumMinus.Length + 1);
                            cumMinus[cumMinus.Length - 1] = Math.Max(0, target - slack - temp2 + cumMinus[cumMinus.Length - 2]);

                            //upper control limit
                            Array.Resize(ref upper, upper.Length + 1);
                            upper[upper.Length - 1] = Math.Max(0, upper[upper.Length - 2] + temp2 - target - slack);

                            //lower control limit
                            Array.Resize(ref lower, lower.Length + 1);
                            lower[lower.Length - 1] = Math.Min(0, lower[lower.Length - 2] + temp2 - target + slack);


                            //check on d(y[t]) = I(y[t] < N), if d(y[t]) = 1 then STOP
                            if (y[y.Length - 1] < N)
                            {
                                MessageBox.Show("Значение y("+ temp1 + ") = " + y[y.Length - 1] + " меньше, чем N = " + N + ".\nЗначение P(" + ((temp1 % 1) * 24) + ") = " + temp2, "Момент разладки");
                                break;
                            }
                            else
                            {
                                //replace X = {P(ts), P(ts+1), ... P(tf)}
                                int ind = 0;
                                while(ind < x.Length - 1)
                                {
                                    x[ind] = x[ind + 1];
                                    ind++;
                                }
                                if (ind == k - 1)
                                    x[ind] = temp2;
                            }                            
                        }
                        if (_dic.Count > k && _dic.Count % 1 == 0)
                        {
                            GraphicsOxiPlot.s2.Points.Add(new DataPoint(temp1, average[average.Length - 1]));
                            //GraphicsOxiPlot.s3.Points.Add(new DataPoint(temp1, cum[cum.Length - 1]));
                            //GraphicsOxiPlot.s4.Points.Add(new DataPoint(temp1, y[y.Length - 1]));
                            //GraphicsOxiPlot.s5.Points.Add(new DataPoint(temp1, cumPlus[cumPlus.Length - 1]));
                            //GraphicsOxiPlot.s6.Points.Add(new DataPoint(temp1, cumMinus[cumMinus.Length - 1]));
                        }

                        temp1 = 0;
                        temp2 = 0;
                    }

                    app.Quit();
                });
                task.Start();
            }
        }

        interface IGraphics
        {
            void OutputGraphics();
        }
        
        public class GraphicsOxiPlot : IGraphics
        {
            static public LineSeries s1;
            static public LineSeries s2;
            static public LineSeries s3;
            static public LineSeries s4;
            static public LineSeries s5;
            static public LineSeries s6;

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
            #region CUSUM
                
            #endregion
            //--------------------- Building ----------------------------                               
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