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
using System.Diagnostics;

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



        public int NumberWithWhitespaceCount
        {
            get { return (int)GetValue(NumberWithWhitespaceCountProperty); }
            set { SetValue(NumberWithWhitespaceCountProperty, value); }
        }

        // Using a DependencyProperty as the backing store for NumberWithWhitespaceCount.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty NumberWithWhitespaceCountProperty =
            DependencyProperty.Register("NumberWithWhitespaceCount", typeof(int), typeof(MainWindow), new PropertyMetadata(0));



        public int CountryCodeCounter
        {
            get { return (int)GetValue(CountryCodeCounterProperty); }
            set { SetValue(CountryCodeCounterProperty, value); }
        }

        // Using a DependencyProperty as the backing store for CountryCodeCounter.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty CountryCodeCounterProperty =
            DependencyProperty.Register("CountryCodeCounter", typeof(int), typeof(MainWindow), new PropertyMetadata(0));



        public int OtherCountryCodeCounter
        {
            get { return (int)GetValue(OtherCountryCodeCounterProperty); }
            set { SetValue(OtherCountryCodeCounterProperty, value); }
        }

        // Using a DependencyProperty as the backing store for OtherCountryCodeCounter.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty OtherCountryCodeCounterProperty =
            DependencyProperty.Register("OtherCountryCodeCounter", typeof(int), typeof(MainWindow), new PropertyMetadata(0));




        public int MultipleNumberCounter
        {
            get { return (int)GetValue(MultipleNumberCounterProperty); }
            set { SetValue(MultipleNumberCounterProperty, value); }
        }

        // Using a DependencyProperty as the backing store for MultipleNumberCounter.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty MultipleNumberCounterProperty =
            DependencyProperty.Register("MultipleNumberCounter", typeof(int), typeof(MainWindow), new PropertyMetadata(0));




        public int NumberWithAlphabetCounter
        {
            get { return (int)GetValue(NumberWithAlphabetCounterProperty); }
            set { SetValue(NumberWithAlphabetCounterProperty, value); }
        }

        // Using a DependencyProperty as the backing store for NumberWithAlphabetCounter.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty NumberWithAlphabetCounterProperty =
            DependencyProperty.Register("NumberWithAlphabetCounter", typeof(int), typeof(MainWindow), new PropertyMetadata(0));




        public int LandlineCounter
        {
            get { return (int)GetValue(LandlineCounterProperty); }
            set { SetValue(LandlineCounterProperty, value); }
        }

        // Using a DependencyProperty as the backing store for LandlineCounter.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty LandlineCounterProperty =
            DependencyProperty.Register("LandlineCounter", typeof(int), typeof(MainWindow), new PropertyMetadata(0));




        private void Button_Click(object sender, RoutedEventArgs e)
        {
            listbox.ItemsSource = null;
            try
            {
                var _columnLetterPosition = GetColumnName(dataColumn.Position);

                var _address = $"{_columnLetterPosition}:{_columnLetterPosition}";

                var _cellRange = sheet.Cells[_address];

                //System.Diagnostics.Debug.WriteLine($"Total row count: {_phoneNumberColumn.Count()}");

                foreach (var cell in _cellRange)
                {
                    //string _holder = "";
                    //string _midHolder = "";

                    var _sourceString = cell.Value.ToString();
                    
                    if (Regex.IsMatch(_sourceString, @"\s+"))
                    {
                        NumberWithWhitespaceCount++;

                        if(chkBox_removeWhiteSpace.IsChecked == true)
                            _sourceString = Regex.Replace(_sourceString, @"\s+", "");
                    }

                    if (Regex.IsMatch(_sourceString, @"@/&"))
                    {
                        MultipleNumberCounter++;
                    }

                    if (Regex.IsMatch(_sourceString, @""))
                    {

                    }

                    if (_sourceString.StartsWith("(+60)"))
                    {
                        CountryCodeCounter++;
                        _sourceString = _sourceString.Replace("(+60)", "0");
                    }
                    else if (_sourceString.StartsWith("+60"))
                    {
                        CountryCodeCounter++;
                        _sourceString = _sourceString.Replace("+60", "0");
                    }

                    if (Regex.IsMatch(_sourceString, @"\b.+-.+"))
                    {
                        NumberWithHyphenCount++;
                        if (chkBox_removeHyphen.IsChecked == true)
                        {
                            _sourceString = Regex.Replace(_sourceString, @"-", "");
                        }                       
                    }

                    numbers.Add(_sourceString);
                }

                listbox.ItemsSource = numbers;
                IsExportButtonEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        
        ExcelPackage package;

        private async void btn_loadFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openfileDialog1 = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
            };

            if (openfileDialog1.ShowDialog() == true)
            {
                FileInfo file = new FileInfo(openfileDialog1.FileName);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                package = new ExcelPackage();

                await package.LoadAsync(file);
                workbook = package.Workbook;

                txtBlock_filename.Content = openfileDialog1.SafeFileName;

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

                if (_ws.Tables.Count > 0)
                {
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
                    MessageBox.Show("No table(s) detected in this sheet. please create a table by pressing Ctrl + T to proceed");
                }                
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



        public bool IsExportButtonEnabled
        {
            get { return (bool)GetValue(IsExportButtonEnabledProperty); }
            set { SetValue(IsExportButtonEnabledProperty, value); }
        }

        // Using a DependencyProperty as the backing store for IsExportButtonEnabled.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsExportButtonEnabledProperty =
            DependencyProperty.Register("IsExportButtonEnabled", typeof(bool), typeof(MainWindow), new PropertyMetadata(false));



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

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (Clipboard.ContainsText(TextDataFormat.Text))
            {
                string _sourceString = Clipboard.GetText(TextDataFormat.Text);
                // Do whatever you need to do with clipboardText

                if (!string.IsNullOrEmpty(_sourceString))
                {
                    try
                    {
                        if (Regex.IsMatch(_sourceString, @"\s+"))
                        {
                            _sourceString = Regex.Replace(_sourceString, @"\s+", "");                               
                        }


                        if (_sourceString.StartsWith("(+60)"))
                        {
                            _sourceString = _sourceString.Replace("(+60)", "0");
                        }
                        else if (_sourceString.StartsWith("+60"))
                        {
                            _sourceString = _sourceString.Replace("+60", "0");
                        }

                        if (Regex.IsMatch(_sourceString, @"\b.+-.+"))
                        {
                            _sourceString = Regex.Replace(_sourceString, @"-", "");
                        }

                        if (!_sourceString.StartsWith("6"))
                        {
                            _sourceString = $"6{_sourceString}";
                        }

                        if (chkBox_useNativeWhatsapp.IsChecked == true)
                        {
                            Process.Start(new ProcessStartInfo($"whatsapp://send?phone={_sourceString}&text=message") { UseShellExecute = true });
                        }
                        else
                        {
                            Process.Start(new ProcessStartInfo($"https://api.whatsapp.com/send?phone={_sourceString}&text=message") { UseShellExecute = true });
                        }
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }
    }
}
