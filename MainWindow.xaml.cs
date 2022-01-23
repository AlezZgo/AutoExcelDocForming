using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;

namespace AutoExcel
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string currentDir = Environment.CurrentDirectory;
            string folderName = @"\signals\";
            DirectoryInfo signalsFolder = new DirectoryInfo(currentDir + folderName);
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("MySheet");

                var someCells = sheet.Cells["A1:C1"];
                someCells.Style.Font.Bold = true;
                someCells.Style.Font.Color.SetColor(Color.Ivory);
                someCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                someCells.Style.Fill.BackgroundColor.SetColor(Color.Navy);

                ExcelRange headerFirstCell = sheet.Cells[1, 1];
                headerFirstCell.Value = "Название сигнала";
                ExcelRange headerSecondCell = sheet.Cells[1, 2];
                headerSecondCell.Value = "СЛС";
                ExcelRange headerThirdCell = sheet.Cells[1, 3];
                headerThirdCell.Value = "ТА";
                try
                {
                    for (int i = 0; i < signalsFolder.GetDirectories().Length; i++)
                    {
                        string sls_puth = signalsFolder + signalsFolder.GetDirectories()[i].ToString() + @"\SLS.txt";
                        string ta_puth = signalsFolder + signalsFolder.GetDirectories()[i].ToString() + @"\TA.txt";

                        ExcelRange firstCell = sheet.Cells[i + 2, 1];
                        firstCell.Value = signalsFolder.GetDirectories()[i].Name;

                        try
                        {
                            using (StreamReader sr = new StreamReader(File.Open(sls_puth, FileMode.Open), System.Text.Encoding.Default))
                            {
                                string sls_txt = sr.ReadToEnd();
                                ExcelRange secondCell = sheet.Cells[i + 2, 2];
                                
                                secondCell.Value = sls_txt;
                                
                            }
                            using (StreamReader sr = new StreamReader(File.Open(ta_puth, FileMode.Open), System.Text.Encoding.Default))
                            {
                                string ta_txt = sr.ReadToEnd();
                                ExcelRange thirdCell = sheet.Cells[i + 2, 3];
                                thirdCell.Value = ta_txt;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }

                        sheet.Cells.AutoFitColumns();

                        package.SaveAs(new FileInfo(@"Результаты.xlsx"));

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\n Вероятно нарушена структура каталогов, обратитесь к оператору ТА");
                }
            }

        }
    }
}
