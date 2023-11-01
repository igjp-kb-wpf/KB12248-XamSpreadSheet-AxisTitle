using Infragistics.Documents.Excel;
using Infragistics.Documents.Excel.Charts;
using System.Windows;

namespace WpfApp_XamSpreadSheet_AxisTitle
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Workbook workbook = new Workbook(WorkbookFormat.Excel2007);
            Worksheet sheet = workbook.Worksheets.Add("Sheet1");

            // X axis labels
            sheet.GetCell("A1").Value = "January";
            sheet.GetCell("B1").Value = "February";
            sheet.GetCell("C1").Value = "March";
            sheet.GetCell("D1").Value = "April";

            // Data
            sheet.GetCell("A3").Value = 10;
            sheet.GetCell("B3").Value = 20;
            sheet.GetCell("C3").Value = 30;
            sheet.GetCell("D3").Value = 40;

            sheet.GetCell("A5").Value = 15;
            sheet.GetCell("B5").Value = 25;
            sheet.GetCell("C5").Value = 23;
            sheet.GetCell("D5").Value = 45;

            sheet.GetCell("A7").Value = 13;
            sheet.GetCell("B7").Value = 23;
            sheet.GetCell("C7").Value = 39;
            sheet.GetCell("D7").Value = 11;

            WorksheetCell cell1 = sheet.GetCell("E7");
            WorksheetCell cell2 = sheet.GetCell("M30");

            WorksheetChart chart1 = sheet.Shapes.AddChart(ChartType.ColumnClustered, cell1, new Point(0, 0), cell2, new Point(100, 100));
            chart1.SetSourceData("A1:D1,A3:D7", true);
            ChartTitle chartTitle = new ChartTitle();
            chartTitle.Text = new FormattedString("Title Text");
            chartTitle.Text.GetFont(0).Height = 500;
            chart1.ChartTitle = chartTitle;

            // 以下より Axis タイトルの設定です。
            // チャートから AxisCollection を取得する
            AxisCollection axisCollection = chart1.AxisCollection;

            // 以下より XAxis の設定です。
            // XAxis のタイトル設定用に ChartTitle を生成する
            ChartTitle xAxisChartTitle = new ChartTitle();
            // Text にタイトルを設定する
            xAxisChartTitle.Text = new FormattedString("XAxis Title Text");
            // テキストのサイズ（Height）を設定する
            xAxisChartTitle.Text.GetFont(0).Height = 300;
            // AxisCollection から XAxis を取得する
            Axis xAxis = axisCollection[AxisType.Category, AxisGroup.Primary];
            // XAxis のタイトルに、XAxis のタイトル設定用に生成した ChartTitle を設定する
            xAxis.AxisTitle = xAxisChartTitle;


            // 以下より YAxis の設定です。
            // XAxis のタイトル設定用に ChartTitle を生成する
            ChartTitle yAxisChartTitle = new ChartTitle();
            // Text にタイトルを設定する
            yAxisChartTitle.Text = new FormattedString("YAxis Title Text");
            // テキストのサイズ（Height）を設定する
            yAxisChartTitle.Text.GetFont(0).Height = 300;
            // テキストの向きを調整する場合は Rotation を設定する
            //yAxisChartTitle.Rotation = 90;
            // AxisCollection から XAxis を取得する
            Axis yAxis = axisCollection[AxisType.Value, AxisGroup.Primary];
            // XAxis のタイトルに、XAxis のタイトル設定用に生成した ChartTitle を設定する
            yAxis.AxisTitle = yAxisChartTitle;


            xamSpreadsheet1.Workbook = workbook;

        }
    }
}
