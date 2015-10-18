using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Drawing;
using System.IO;
using OfficeOpenXml.Style;
using Image = System.Drawing.Image;

namespace ImageToExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string FilePath { get; set; }
        private string FileName { get; set; }
        private string OptionsFormatter { get { return "Column Width:{0} | Row Height:{1}"; } }

        private readonly BackgroundWorker _backgroundWorker = new BackgroundWorker();

        private double ColumnWidth { get; set; }
        private double RowHeight { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            _backgroundWorker.WorkerReportsProgress = true;
            _backgroundWorker.ProgressChanged += ProgressChanged;
            _backgroundWorker.DoWork += DoWork;
            _backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
        }

        private void loadButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                DefaultExt = ".png",
                Filter =
                    "PNG Files (*.png)|*.png|JPG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif|BMP Files (*.bmp)|*.bmp|JPEG Files (*.jpeg)|*.jpeg"
            };


            var result = dlg.ShowDialog();
            if (result != true)
            {
                stateLabel.Content = "You have to select image file...";
                FilePath = "";
                startButton.IsEnabled = false;
                return;
            }

            FilePath = dlg.FileName;
            FileName = dlg.SafeFileName.Split('.')[0];
            stateLabel.Content = dlg.SafeFileName;
            startButton.IsEnabled = true;

            var prvImg = new BitmapImage(new Uri(FilePath));
            previewImage.Source = prvImg;

        }

        private void GenerateExcelFromImage(string imagePath)
        {
            using (var p = new ExcelPackage())
            {
                var totalProgress = 0f;
                var currentProgress = 0;
                var progress = 0f;

                p.Workbook.Properties.Author = "Patryk Daru";
                p.Workbook.Properties.Title = "ExcelToImage Converted";
                p.Workbook.Properties.Company = "DaruHQ";
                p.Workbook.Properties.Comments = FileName;

                var ws = p.Workbook.Worksheets.Add(FileName);

                var image = new Bitmap(imagePath);
                totalProgress = image.Width * image.Height;

                _backgroundWorker.ReportProgress(currentProgress, "Generating xlsx...");
                
                for (var x = 1; x < image.Height; x++)
                {
                    for (var y = 1; y < image.Width; y++)
                    {
                        ws.Column(y).Width = 0.3;
                        ws.Cells[x, y].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[x, y].Style.Fill.BackgroundColor.SetColor(image.GetPixel(y, x));

                        progress++;
                        currentProgress = (int)(progress / totalProgress * 100);
                        _backgroundWorker.ReportProgress(currentProgress);
                    }
                    ws.Row(x).Height = 2;
                }

                _backgroundWorker.ReportProgress(currentProgress, "Saving xlsx...");
                var bin = p.GetAsByteArray();
                File.WriteAllBytes(string.Format("{0}.xlsx", FileName), bin);
            }
        }

        private void startButton_Click(object sender, RoutedEventArgs e)
        {
            startButton.IsEnabled = false;
            loadButton.IsEnabled = false;
            optionsGroupBox.IsEnabled = false;
            _backgroundWorker.RunWorkerAsync();   
        }

        private void DoWork(object sender, DoWorkEventArgs e)
        {
            GenerateExcelFromImage(FilePath);
        }

        private void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;

            var labelText = e.UserState as string;
            if (!string.IsNullOrEmpty(labelText))
            {
                stateLabel.Content = labelText;
            }
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar.Value = progressBar.Maximum;
            stateLabel.Content = "Done";
            optionsGroupBox.IsEnabled = true;
            loadButton.IsEnabled = true;
        }

        private void columnSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            ColumnWidth = Math.Round(e.NewValue, 2);
            optionsStatus.Content = string.Format(OptionsFormatter, ColumnWidth, RowHeight);
        }

        private void rowSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            RowHeight = Math.Round(e.NewValue, 2);
            optionsStatus.Content = string.Format(OptionsFormatter, ColumnWidth, RowHeight);
        }

    }
}
