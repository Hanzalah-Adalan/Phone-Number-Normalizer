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
using System.Net.Http;
using System.Diagnostics;
using HtmlAgilityPack;
using RestSharp;
using MaterialDesignExtensions.Model;

namespace Phone_Number_Normalizer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private LowLevelKeyboardListener _listener;
        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            _listener = new LowLevelKeyboardListener();
            _listener.OnKeyPressed += _listener_OnKeyPressed;

            _listener.HookKeyboard();
        }

        public Visibility IsFetchingUpline
        {
            get { return (Visibility)GetValue(IsFetchingUplineProperty); }
            set { SetValue(IsFetchingUplineProperty, value); }
        }
        // Using a DependencyProperty as the backing store for IsFetchingUpline.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsFetchingUplineProperty =
            DependencyProperty.Register("IsFetchingUpline", typeof(Visibility), typeof(MainWindow), new PropertyMetadata(Visibility.Collapsed));



        public bool IsFindUplineEnabled
        {
            get { return (bool)GetValue(IsFindUplineEnabledProperty); }
            set { SetValue(IsFindUplineEnabledProperty, value); }
        }

        // Using a DependencyProperty as the backing store for IsFindUplineEnabled.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsFindUplineEnabledProperty =
            DependencyProperty.Register("IsFindUplineEnabled", typeof(bool), typeof(MainWindow), new PropertyMetadata(true));



        void _listener_OnKeyPressed(object sender, KeyPressedArgs e)
        {
            if (Keyboard.IsKeyDown(Key.LeftCtrl) && Keyboard.IsKeyDown(Key.LeftShift) && e.KeyPressed == Key.U)
            {
                if (this.WindowState == WindowState.Minimized)
                    this.WindowState = WindowState.Normal;

                this.Activate();

                if (Clipboard.ContainsText(TextDataFormat.Text))
                {
                    string _sourceString = Clipboard.GetText(TextDataFormat.Text);

                    if (!string.IsNullOrEmpty(_sourceString))
                    {
                        try
                        {
                            if (Regex.IsMatch(_sourceString, @"\s+"))//remove all whitespaces
                            {
                                _sourceString = Regex.Replace(_sourceString, @"\s+", "");
                            }

                            customerID.Text = _sourceString;
                            btn_findUpline_Click(null, null);

                            //if (chkBox_useNativeWhatsapp.IsChecked == true)
                            //{
                            //    Process.Start(new ProcessStartInfo($"whatsapp://send?phone={_sourceString}&text=message") { UseShellExecute = true });
                            //}
                            //else
                            //{
                            //    Process.Start(new ProcessStartInfo($"https://api.whatsapp.com/send?phone={_sourceString}&text=message") { UseShellExecute = true });
                            //}

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _listener.UnHookKeyboard();
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

        class UplineInfo
        {
            public string Name { get; set; }
            public string Email { get; set; }
        }
        enum AgentGroup
        {
            Dropship,
            PartnerCircle,
            PremierPartner
        }
        class AgentInfo
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string Group { get; set; }
            public string EmailAddress { get; set; }
            public string ContactNumber { get; set; }
            public UplineInfo Upline { get; set; }
        }

        string GetGroup(HtmlDocument doc)
        {
            try
            {
                var _groupXP = "//*[@id='defaultSelected']/td[4]";
                var _group = doc.DocumentNode.SelectNodes(_groupXP).FirstOrDefault();
                var _a = _group.ChildNodes.FirstOrDefault();
                return _a.InnerText;
            }
            catch (Exception)
            {
                return null;
            }
        }

        RestClient client;
        RestRequest request;
        IRestResponse response;
        private async void btn_findUpline_Click(object sender, RoutedEventArgs e)
        {
            IsFetchingUpline = Visibility.Visible;
            IsFindUplineEnabled = false;

            stepper.Steps.Clear();
            var _custID = customerID.Text.Trim();
             
            try
            {
                client = new RestClient($"https://www.sabella.com.my/secure/orders.php?startdate=&enddate=&payment_method=&shipping_method=&status=&search_field={_custID}&platform=all&search=Search+Order")
                {
                    Timeout = -1
                };
                request = new RestRequest(Method.GET);
                request.AddHeader("Upgrade-Insecure-Requests", "1");
                client.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.67 Safari/537.36 Edg/87.0.664.47";
                request.AddHeader("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9");
                request.AddHeader("Sec-Fetch-Site", "same-origin");
                request.AddHeader("Sec-Fetch-Mode", "navigate");
                request.AddHeader("Sec-Fetch-User", "?1");
                request.AddHeader("Sec-Fetch-Dest", "document");
                request.AddHeader("Referer", $"https://www.sabella.com.my/secure/orders.php?startdate=&enddate=&payment_method=&shipping_method=&status=&search_field={_custID}&platform=all&search=Search+Order");
                request.AddHeader("Accept-Encoding", "gzip, deflate, br");
                request.AddHeader("Accept-Language", "en-US,en;q=0.9");
                request.AddHeader("Cookie", "_ga=GA1.3.1550589700.1605787753; _fbp=fb.2.1606570879165.709323232; _gid=GA1.3.295114435.1606570879; sabellaAID=iav0l0l5gvu4q9dq18u6ukv2c1; tokenid=0cpo4mcfl23hlakhn0q1s14g85; newuser=Yes");
                response = await client.ExecuteAsync(request);

                var doc = new HtmlDocument();
                doc.LoadHtml(response.Content);


                var _xp = "//*[@id='defaultSelected']/td[3]";
                var _priceXP = "//*[@id='defaultSelected']/td[6]";
                var _dateXP = "//*[@id='defaultSelected']/td[7]";
                var _statusXP = "//*[@id='defaultSelected']/td[12]/span";

                var _startingEmail = "";
                var value = doc.DocumentNode.SelectNodes(_xp).FirstOrDefault();
                if (value != null)
                {
                    var _a = value.ChildNodes.FirstOrDefault();
                    if (_a != null)
                    {
                        _startingEmail = _a.LastChild.InnerText;//crucial line

                        txtBox_endCustName.Text = _a.FirstChild.InnerText;
                    }
                }

                var _price = doc.DocumentNode.SelectNodes(_priceXP).FirstOrDefault();
                if (_price != null)
                {
                    var _a = _price.ChildNodes.FirstOrDefault();
                    if (_a != null)
                    {
                        txtBox_endCustPay.Text = _a.InnerText;
                    }
                }

                var _date = doc.DocumentNode.SelectNodes(_dateXP).FirstOrDefault();
                if (_date != null)
                {
                    var _a = _date.ChildNodes.FirstOrDefault();
                    if (_a != null)
                    {
                        txtBox_endCustDate.Text = _a.InnerText;
                    }
                }

                var _status = doc.DocumentNode.SelectNodes(_statusXP).FirstOrDefault();
                if (_status != null)
                {
                    var _a = _status.ChildNodes.FirstOrDefault();
                    if (_a != null)
                    {
                        txtBox_endCustStatus.Text = _a.InnerText;
                    }
                }

                AgentInfo _f1 = await GetAgentAsync(_startingEmail);

                stepper.Steps.Add(new Step
                {
                    Header = new StepTitleHeader { FirstLevelTitle = _f1.Group, SecondLevelTitle = $"{_f1.FirstName} {_f1.LastName}" },
                    Content = GetStackPanel(_f1)
                });

                while (_f1.Group != "Premier Partner")
                {
                    _f1 = await GetAgentAsync(_f1.Upline.Email);
                    stepper.Steps.Add(new Step
                    {
                        Header = new StepTitleHeader { FirstLevelTitle = _f1.Group, SecondLevelTitle = $"{_f1.FirstName} {_f1.LastName}" },
                        Content = GetStackPanel(_f1)
                    });
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            IsFindUplineEnabled = true;
            IsFetchingUpline = Visibility.Collapsed;
        }

        StackPanel GetStackPanel(AgentInfo _f1)
        {
            var _stkPanel = new StackPanel();

            var _txtBox = new TextBox
            {
                Text = $"Contact number: {_f1.ContactNumber}{Environment.NewLine}Email: {_f1.EmailAddress}",
                IsReadOnly = true,
                BorderThickness = new Thickness(0),
                Background = new SolidColorBrush(Colors.Transparent),
                TextWrapping = TextWrapping.Wrap
            };
            var _btn = new Button
            {
                Content = "Chat in Whatsapp",
                Background = new SolidColorBrush(Colors.LimeGreen),
                Margin = new Thickness(0, 10, 0, 0),
                BorderThickness = new Thickness(0)
            };
            _btn.Click += (s, e) =>
            {
                OpenChatInWhatsapp(_f1.ContactNumber);
            };

            _stkPanel.Children.Add(_txtBox);
            _stkPanel.Children.Add(_btn);

            return _stkPanel;
        }

        async Task<AgentInfo> GetAgentAsync(string email)
        {
            client = new RestClient($"https://www.sabella.com.my/secure/customers.php?search={email}&customers_agent_id=&agent_status=")
            {
                Timeout = -1
            };

            request = new RestRequest(Method.GET);
            request.AddHeader("Upgrade-Insecure-Requests", "1");
            request.AddHeader("Referer", "https://www.sabella.com.my/secure/customers.php");
            request.AddHeader("Cookie", "_ga=GA1.3.1550589700.1605787753; _fbp=fb.2.1606570879165.709323232; _gid=GA1.3.295114435.1606570879; sabellaAID=iav0l0l5gvu4q9dq18u6ukv2c1; tokenid=0cpo4mcfl23hlakhn0q1s14g85");
            response = await client.ExecuteAsync(request);

            var doc = new HtmlDocument();
            doc.LoadHtml(response.Content);

            var _agentEmailXP = "//*[@id='defaultSelected']/td[5]/a";
            var _phoneNumberXP = "//*[@id='defaultSelected']/td[7]";
            var _uplineNameAndEmailXP = "//*[@id='defaultSelected']/td[6]";
            var _firstNameXP = "//*[@id='defaultSelected']/td[2]";
            var _lastNameXP = "//*[@id='defaultSelected']/td[1]";
            
            var _agent = new AgentInfo();
            var _upline = new UplineInfo();

            var _agentFName = doc.DocumentNode.SelectNodes(_firstNameXP).FirstOrDefault();
            if (_agentFName != null)
            {
                var _a = _agentFName.ChildNodes.FirstOrDefault();
                if (_a != null)
                {
                    _agent.FirstName = _a.InnerText;
                }
            }

            var _agentLName = doc.DocumentNode.SelectNodes(_lastNameXP).FirstOrDefault();
            if (_agentLName != null)
            {
                var _a = _agentLName.ChildNodes.FirstOrDefault();
                if (_a != null)
                {
                    _agent.LastName = _a.InnerText;
                }
            }

            var _agentEmail = doc.DocumentNode.SelectNodes(_agentEmailXP).FirstOrDefault();
            if (_agentEmail != null)
            {
                var _a = _agentEmail.ChildNodes.FirstOrDefault();
                if (_a != null)
                {
                    _agent.EmailAddress = _a.InnerText;
                }
            }

            var nameAndEmailUpline = doc.DocumentNode.SelectNodes(_uplineNameAndEmailXP).FirstOrDefault();
            if (nameAndEmailUpline != null)
            {
                var _a = nameAndEmailUpline.ChildNodes.FirstOrDefault();
                if (_a != null)
                {
                    _upline.Name = _a.InnerText.Trim();
                    _upline.Email = _a.NextSibling.NextSibling.InnerText.Trim();

                    _agent.Upline = _upline;
                }
            }

            var _phoneNo = doc.DocumentNode.SelectNodes(_phoneNumberXP).FirstOrDefault();
            if (_phoneNo != null)
            {
                _agent.ContactNumber = _phoneNo.FirstChild.InnerText;
            }

            _agent.Group = GetGroup(doc);

            return _agent;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

        }

        public void OpenChatInWhatsapp(string _sourceString)
        {
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

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (Clipboard.ContainsText(TextDataFormat.Text))
            {
                string _sourceString = Clipboard.GetText(TextDataFormat.Text);
                // Do whatever you need to do with clipboardText

                OpenChatInWhatsapp(_sourceString);
            }
        }

        private void btn_openOrderDetailsPage_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var _url = $"https://www.sabella.com.my/secure/orders.php?page=1&oID={customerID.Text.Trim()}&action=edit";

                Process.Start(new ProcessStartInfo(_url) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
