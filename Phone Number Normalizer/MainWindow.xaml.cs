using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace Phone_Number_Normalizer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
        }

        List<string> numbers = new List<string>();
        ExcelWorkbook workbook;
        ExcelWorksheet sheet;



        public int NumberWithHyphenCount
        {
            get { return (int)GetValue(NumberWithHyphenCountProperty); }
            set { SetValue(NumberWithHyphenCountProperty, value); }
        }

        // Using a DependencyProperty as the backing store for NumberWithHyphenCount.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty NumberWithHyphenCountProperty =
            DependencyProperty.Register("NumberWithHyphenCount", typeof(int), typeof(MainWindow), new PropertyMetadata(0));


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var _columnLetterPosition = GetColumnName(dataColumn.Position);

                var _address = $"{_columnLetterPosition}:{_columnLetterPosition}";

                var _cellRange = sheet.Cells[_address];

                //System.Diagnostics.Debug.WriteLine($"Total row count: {_phoneNumberColumn.Count()}");

                foreach (var cell in _cellRange)
                {
                    string _holder = "";
                    string _midHolder = "";

                    var _sourceString = cell.Value.ToString();
                    var _sb = Regex.Replace(_sourceString, @"\s+", "");

                    if (_sb.StartsWith("(+60)"))
                    {
                        _holder = _sb.Replace("(+60)", "0");
                    }
                    else if (_sb.StartsWith("+60"))
                    {
                        _holder = _sb.Replace("+60", "0");
                    }
                    else
                    {
                        _holder = _sb;
                    }

                    if (chkBox_removeHyphen.IsChecked == true && Regex.IsMatch(_holder, @"\b.+-.+"))
                    {                        
                        _midHolder = Regex.Replace(_holder, @"-", "");
                    }
                    else
                    {
                        NumberWithHyphenCount++;
                        _midHolder = _holder;
                    }

                    numbers.Add(_midHolder);
                }

                listbox.ItemsSource = numbers;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnTestSingleBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(txtBox.Text))
            {
                if (Regex.IsMatch(txtBox.Text, @"\b.+-.+"))
                {
                    txtBox.Text = Regex.Replace(txtBox.Text, @"-", "");

                    MessageBox.Show("Successfully removed");
                }
                else
                {
                    MessageBox.Show("Failed to remove");
                }

            }
        }

        
        ExcelPackage package;

        private async void btn_loadFile_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openfileDialog1 = new OpenFileDialog();
            if (openfileDialog1.ShowDialog() == true)
            {
                FileInfo file = new FileInfo(openfileDialog1.FileName);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                package = new ExcelPackage();

                await package.LoadAsync(file);
                workbook = package.Workbook;

                txtBlock_filename.Text = openfileDialog1.SafeFileName;

                foreach (var item in package.Workbook.Worksheets)
                {
                    cmbBox_sheetSelector.Items.Add(item);
                }
            }

            IsSheetSelectorEnabled = !string.IsNullOrEmpty(openfileDialog1.FileName);

        }

        private void cmbBox_sheetSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbBox_sheetSelector.SelectedItem is ExcelWorksheet _ws)
            {
                sheet = _ws;

                var _columns = _ws.Tables.FirstOrDefault().Columns;
                foreach (var item in _columns)
                {
                    cmbBox_columnSelector.Items.Add(item);
                }

                IsColumnSelectorEnabled = true;
            }
            else
            {
                IsColumnSelectorEnabled = false;
            }
        }

        static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }

        ExcelTableColumn dataColumn;
        private void cmbBox_columnSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbBox_columnSelector.SelectedItem is ExcelTableColumn dc)
            {                                
                dataColumn = dc;
                IsManipulationButtonsEnabled = true;
            }
            else
            {
                dataColumn = null;
                IsManipulationButtonsEnabled = false;
            }
        }



        public bool IsSheetSelectorEnabled
        {
            get { return (bool)GetValue(IsSheetSelectorEnabledProperty); }
            set { SetValue(IsSheetSelectorEnabledProperty, value); }
        }

        // Using a DependencyProperty as the backing store for IsSheetSelectorEnabled.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsSheetSelectorEnabledProperty =
            DependencyProperty.Register("IsSheetSelectorEnabled", typeof(bool), typeof(MainWindow), new PropertyMetadata(false));



        public bool IsColumnSelectorEnabled
        {
            get { return (bool)GetValue(IsColumnSelectorEnabledProperty); }
            set { SetValue(IsColumnSelectorEnabledProperty, value); }
        }

        // Using a DependencyProperty as the backing store for IsColumnSelectorEnabled.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsColumnSelectorEnabledProperty =
            DependencyProperty.Register("IsColumnSelectorEnabled", typeof(bool), typeof(MainWindow), new PropertyMetadata(false));



        public bool IsManipulationButtonsEnabled
        {
            get { return (bool)GetValue(IsManipulationButtonsEnabledProperty); }
            set { SetValue(IsManipulationButtonsEnabledProperty, value); }
        }

        // Using a DependencyProperty as the backing store for IsManipulationButtonsEnabled.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsManipulationButtonsEnabledProperty =
            DependencyProperty.Register("IsManipulationButtonsEnabled", typeof(bool), typeof(MainWindow), new PropertyMetadata(false));


    }
}
