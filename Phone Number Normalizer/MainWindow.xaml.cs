﻿using Microsoft.Win32;
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
using System.Net.Http;
using System.Diagnostics;
using HtmlAgilityPack;
using RestSharp;
using MaterialDesignExtensions.Model;
using System.Net;
using OfficeOpenXml.Table;
using OfficeOpenXml;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using System.Threading;
using Google.Apis.Util.Store;
using Google.Apis.Sheets.v4.Data;
using System.Diagnostics.CodeAnalysis;
using Phone_Number_Normalizer.Models;
using Phone_Number_Normalizer.Controls;

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
            this.Loaded += MainWindow_Loaded;
        }

        #region Contacts normalizer
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
            numbers.Clear();
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

                        if (chkBox_removeWhiteSpace.IsChecked == true)
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

                cmbBox_sheetSelector.Items.Clear();
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
                cmbBox_columnSelector.Items.Clear();
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
        #endregion



        public string DHLTrackingID
        {
            get { return (string)GetValue(DHLTrackingIDProperty); }
            set { SetValue(DHLTrackingIDProperty, value); }
        }

        // Using a DependencyProperty as the backing store for DHLTrackingID.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty DHLTrackingIDProperty =
            DependencyProperty.Register("DHLTrackingID", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));



        public string LastUpdateStatusCode
        {
            get { return (string)GetValue(LastUpdateStatusCodeProperty); }
            set { SetValue(LastUpdateStatusCodeProperty, value); }
        }

        // Using a DependencyProperty as the backing store for LastUpdateStatusCode.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty LastUpdateStatusCodeProperty =
            DependencyProperty.Register("LastUpdateStatusCode", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));


        public string LastUpdateDate
        {
            get { return (string)GetValue(LastUpdateDateProperty); }
            set { SetValue(LastUpdateDateProperty, value); }
        }

        // Using a DependencyProperty as the backing store for LastUpdateDate.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty LastUpdateDateProperty =
            DependencyProperty.Register("LastUpdateDate", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));



        public string LastUpdateLocation
        {
            get { return (string)GetValue(LastUpdateLocationProperty); }
            set { SetValue(LastUpdateLocationProperty, value); }
        }

        // Using a DependencyProperty as the backing store for LastUpdateLocation.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty LastUpdateLocationProperty =
            DependencyProperty.Register("LastUpdateLocation", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));



        public string AddressMock
        {
            get { return (string)GetValue(AddressMockProperty); }
            set { SetValue(AddressMockProperty, value); }
        }

        // Using a DependencyProperty as the backing store for AddressMock.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty AddressMockProperty =
            DependencyProperty.Register("AddressMock", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));



        public string DestinationAddress
        {
            get { return (string)GetValue(DestinationAddressProperty); }
            set { SetValue(DestinationAddressProperty, value); }
        }

        // Using a DependencyProperty as the backing store for DestinationAddress.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty DestinationAddressProperty =
            DependencyProperty.Register("DestinationAddress", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));



        public string OriginAddress
        {
            get { return (string)GetValue(OriginAddressProperty); }
            set { SetValue(OriginAddressProperty, value); }
        }

        // Using a DependencyProperty as the backing store for OriginAddress.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty OriginAddressProperty =
            DependencyProperty.Register("OriginAddress", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));




        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            var _dtr = $@"SABELLA HOLDINGS SDN BHD{Environment.NewLine}LOT 19,{Environment.NewLine}JALAN PERUSAHAAN 2,{Environment.NewLine}KAW PERINDUSTRIAN BATU CAVES{Environment.NewLine}BATU CAVES{Environment.NewLine}SEL{Environment.NewLine}MALAYSIA{Environment.NewLine}68100{Environment.NewLine}6036177190{Environment.NewLine}SABELLA HQ@MAIL.COM{Environment.NewLine}";

            AddressMock = _dtr;
            
            txtBox_cookie.Text = Properties.Settings.Default.HCookie;
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
            if (Keyboard.IsKeyDown(Key.LeftCtrl) && Keyboard.IsKeyDown(Key.LeftShift) && e.KeyPressed == Key.O)
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

            var _theCookie = Properties.Settings.Default.HCookie;

            try
            {
                //var _cfCookie = "_ga=GA1.3.1550589700.1605787753; _fbp=fb.2.1606570879165.709323232; _gid=GA1.3.295114435.1606570879; sabellaAID=iav0l0l5gvu4q9dq18u6ukv2c1; __cfduid=dbd139d25619f02bb6e68f30ab8fcb0681606870835; newuser=Yes; tokenid=b9gla0an7qg041clg94c9j3bi3";
                ////dont find anis folder
                //client = new RestClient($"https://www.sabella.com.my/secure/orders.php?startdate=&enddate=&payment_method=&shipping_method=&status=&search_field={_custID}&platform=all&search=Search+Order")
                //{
                //    Timeout = -1
                //};
                //request = new RestRequest(Method.GET);
                //request.AddHeader("Upgrade-Insecure-Requests", "1");
                //client.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.67 Safari/537.36 Edg/87.0.664.47";
                //request.AddHeader("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9");
                //request.AddHeader("Sec-Fetch-Site", "same-origin");
                //request.AddHeader("Sec-Fetch-Mode", "navigate");
                //request.AddHeader("Sec-Fetch-User", "?1");
                //request.AddHeader("Sec-Fetch-Dest", "document");
                //request.AddHeader("Referer", $"https://www.sabella.com.my/secure/orders.php?startdate=&enddate=&payment_method=&shipping_method=&status=&search_field={_custID}&platform=all&search=Search+Order");
                //request.AddHeader("Accept-Encoding", "gzip, deflate, br");
                //request.AddHeader("Accept-Language", "en-US,en;q=0.9");
                //request.AddHeader("Cookie", _cfCookie);
                //response = await client.ExecuteAsync(request);

                //var doc = new HtmlDocument();
                //doc.LoadHtml(response.Content);


                //var _xp = "//*[@id='defaultSelected']/td[3]";
                //var _priceXP = "//*[@id='defaultSelected']/td[6]";
                //var _dateXP = "//*[@id='defaultSelected']/td[7]";
                //var _statusXP = "//*[@id='defaultSelected']/td[12]/span";

                var _startingEmail = "paepaezah81@yahoo.com";
                //var value = doc.DocumentNode.SelectNodes(_xp).FirstOrDefault();
                //if (value != null)
                //{
                //    var _a = value.ChildNodes.FirstOrDefault();
                //    if (_a != null)
                //    {
                //        _startingEmail = _a.LastChild.InnerText;//crucial line

                //        txtBox_endCustName.Text = $"Name: {_a.FirstChild.InnerText}";
                //    }
                //}

                //var _price = doc.DocumentNode.SelectNodes(_priceXP).FirstOrDefault();
                //if (_price != null)
                //{
                //    var _a = _price.ChildNodes.FirstOrDefault();
                //    if (_a != null)
                //    {
                //        txtBox_endCustPay.Text = $"Total: {_a.InnerText}";
                //    }
                //}

                //var _date = doc.DocumentNode.SelectNodes(_dateXP).FirstOrDefault();
                //if (_date != null)
                //{
                //    var _a = _date.ChildNodes.FirstOrDefault();
                //    if (_a != null)
                //    {
                //        txtBox_endCustDate.Text = $"Order Date: {_a.InnerText}";
                //    }
                //}

                //var _status = doc.DocumentNode.SelectNodes(_statusXP).FirstOrDefault();
                //if (_status != null)
                //{
                //    var _a = _status.ChildNodes.FirstOrDefault();
                //    if (_a != null)
                //    {
                //        txtBox_endCustStatus.Text = $"Status: {_a.InnerText}";
                //    }
                //}

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

            var _newKumkies = "_ga=GA1.3.1550589700.1605787753; _fbp=fb.2.1606570879165.709323232; _gid=GA1.3.295114435.1606570879; sabellaAID=iav0l0l5gvu4q9dq18u6ukv2c1; __cfduid=dbd139d25619f02bb6e68f30ab8fcb0681606870835; tokenid=b9gla0an7qg041clg94c9j3bi3";

            request = new RestRequest(Method.GET);
            request.AddHeader("Upgrade-Insecure-Requests", "1");
            request.AddHeader("Referer", "https://www.sabella.com.my/secure/customers.php?search=paepaezah81%40yahoo.com&customers_agent_id=&agent_status=");
            request.AddHeader("Cookie", _newKumkies);
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

        private void btn_register_cookie_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(txtBox_cookie.Text))
            {
                Properties.Settings.Default.HCookie = txtBox_cookie.Text;
                Properties.Settings.Default.Save();
            }
        }

        public string ProductName
        {
            get { return (string)GetValue(ProductNameProperty); }
            set { SetValue(ProductNameProperty, value); }
        }

        // Using a DependencyProperty as the backing store for ProductName.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty ProductNameProperty =
            DependencyProperty.Register("ProductName", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));



        public string ReferenceNumber
        {
            get { return (string)GetValue(ReferenceNumberProperty); }
            set { SetValue(ReferenceNumberProperty, value); }
        }

        // Using a DependencyProperty as the backing store for ReferenceNumber.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty ReferenceNumberProperty =
            DependencyProperty.Register("ReferenceNumber", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));


        public string ReferenceType
        {
            get { return (string)GetValue(ReferenceTypeProperty); }
            set { SetValue(ReferenceTypeProperty, value); }
        }

        // Using a DependencyProperty as the backing store for ReferenceType.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty ReferenceTypeProperty =
            DependencyProperty.Register("ReferenceType", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));

        public string DHLService
        {
            get { return (string)GetValue(DHLServiceProperty); }
            set { SetValue(DHLServiceProperty, value); }
        }

        // Using a DependencyProperty as the backing store for DHLService.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty DHLServiceProperty =
            DependencyProperty.Register("DHLService", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));

        private async void btn_trackDHLShipment_Click(object sender, RoutedEventArgs e)
        {
            var client = new RestClient($"https://api-eu.dhl.com/track/shipments?trackingNumber={txtBox_DHLTrackingID.Text}&language=en&limit=5")
            {
                Timeout = -1
            };
            var request = new RestRequest(Method.GET);
            request.AddHeader("accept", "application/json");
            request.AddHeader("DHL-API-Key", "n6PrQ9GcfQVXapm1MoCOhkXWeztA302m");
            IRestResponse response = await client.ExecuteAsync(request);

            var _shipments = DHL.DhlShipments.FromJson(response.Content);

            string ProcAddress(DHL.Address _address)
            {
                return $"{_address.AddressLocality} {_address.CountryCode} {_address.PostalCode}";
            }

            if (response.StatusCode == HttpStatusCode.OK)
            {
                foreach (var item in _shipments.Shipments)
                {
                    DHLTrackingID = item.Id;

                    var _destination = ProcAddress(item.Destination.Address);
                    var _origin = ProcAddress(item.Origin.Address);

                    OriginAddress = _origin;
                    DestinationAddress = _destination;

                    
                    
                    if (item.Details.Product != null)
                    {
                        var _details_product = $"ProductName: {item.Details.Product.ProductName}";
                        ProductName = _details_product;
                    }

                    if (item.Details.References != null)
                    {
                        var _details_reference = $"Reference Number: {item.Details.References.FirstOrDefault().Number}{Environment.NewLine}Reference Type: {item.Details.References.FirstOrDefault().Type}";

                        ReferenceNumber = item.Details.References.FirstOrDefault().Number;
                        ReferenceType = item.Details.References.FirstOrDefault().Type; 
                    }

                    var _service = $"{item.Service}";
                    DHLService = _service;

                    var _status = $"{item.Status.StatusCode}";

                    stepper.Steps.Clear();
                    foreach (var ev in item.Events)
                    {
                        stepper_tracking.Steps.Add(new Step
                        {
                            Header = new StepTitleHeader
                            {
                                FirstLevelTitle = ev.StatusCode,
                                SecondLevelTitle = ev.Timestamp.ToString()
                            },
                            Content = new TextBlock { Text = ev.StatusStatus, TextWrapping = TextWrapping.Wrap }
                        });
                    }
                }
            }
            else if (response.StatusCode == HttpStatusCode.NotFound)
            {
                MessageBox.Show("Can't find the shipment with tracking ID");
            }


        }

        private async void btn_openSheet_Click(object sender, RoutedEventArgs e)
        {
            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = await GetCredential(),
                ApplicationName = "Google-SheetsSample/0.1",
            });

            // The ID of the spreadsheet to retrieve data from.
            string spreadsheetId = "1miSN4WiyK_2YxhAZx7kpGj0dspQ04HFOF1C7QWbYxUY";  // TODO: Update placeholder value.

            // The A1 notation of the values to retrieve.
            string range = "A:A";  // TODO: Update placeholder value.

            // How values should be represented in the output.
            // The default render option is ValueRenderOption.FORMATTED_VALUE.
            SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum valueRenderOption = (SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum)0;  // TODO: Update placeholder value.

            // How dates, times, and durations should be represented in the output.
            // This is ignored if value_render_option is
            // FORMATTED_VALUE.
            // The default dateTime render option is [DateTimeRenderOption.SERIAL_NUMBER].
            SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum dateTimeRenderOption = (SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum)0;  // TODO: Update placeholder value.

            SpreadsheetsResource.ValuesResource.GetRequest request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);
            //request.Key = "AIzaSyA5--mfHvxooRryQqagNNXOsX-JWyq5k60";
            request.OauthToken = credential.Token.AccessToken;
            request.ValueRenderOption = valueRenderOption;
            request.DateTimeRenderOption = dateTimeRenderOption;

            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            Google.Apis.Sheets.v4.Data.ValueRange response = await request.ExecuteAsync();
            // Data.ValueRange response = await request.ExecuteAsync();

            foreach (var item in response.Values)
            {
                foreach (var _c in item)
                {
                    listView_scanIT.Items.Add(_c.ToString());
                }
            }
        }


        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes =  { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        static UserCredential credential;

        public static async Task<UserCredential> GetCredential()
        {
            // TODO: Change placeholder below to generate authentication credentials. See:
            // https://developers.google.com/sheets/quickstart/dotnet#step_3_set_up_the_sample
            //
            // Authorize using one of the following scopes:
            //     "https://www.googleapis.com/auth/drive"
            //     "https://www.googleapis.com/auth/drive.file"
            //     "https://www.googleapis.com/auth/drive.readonly"
            //     "https://www.googleapis.com/auth/spreadsheets"
            //     "https://www.googleapis.com/auth/spreadsheets.readonly"

            
            
            

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            String range = "Class Data!A2:E";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            ValueRange response = await request.ExecuteAsync();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}", row[0], row[4]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Console.Read();


            return null;
        }



        class PlaceEqualityComparer : IEqualityComparer<Place>
        {
            public bool Equals(Place b1, Place b2)
            {
                if (b2 == null && b1 == null)
                    return true;
                else if (b1 == null || b2 == null)
                    return false;
                else if (b1.GroupKey == b2.GroupKey)
                    return true;
                else
                    return false;
            }

            public int GetHashCode([DisallowNull] Place obj)
            {
                return obj.GetHashCode();
            }
        }

        enum CommonDistrictName
        {
            Sungai,
            Kuala,
            Kota,
            Batu,
            Bandar,
            Pasir,
            Tanah,
            Teluk,
            Alor,
            Bukit,
            Tanjung
        }


        class MyClass : IEqualityComparer<ExcelRangeBase>
        {
            public bool Equals(ExcelRangeBase b1, ExcelRangeBase b2)
            {
                if (b2 == null && b1 == null)
                    return true;
                else if (b1 == null || b2 == null)
                    return false;
                else if (b1.Value.ToString().Trim().ToLower() == b2.Value.ToString().Trim().ToLower())
                    return true;
                else
                    return false;
            }

            public int GetHashCode([DisallowNull] ExcelRangeBase obj)
            {
                var _rdm = obj.Value.ToString().Count() ^ obj.Address.Length;
                return _rdm.GetHashCode();
            }
        }

        private void btn_standardizeRegionName_Click(object sender, RoutedEventArgs e)
        {
            listbox.ItemsSource = null;
            var _thePlaces = new List<Place>();

            try
            {
                var _columnLetterPosition = GetColumnName(dataColumn.Position);

                var _address = $"{_columnLetterPosition}:{_columnLetterPosition}";

                var _cellRange = sheet.Cells[_address];



                var _groupedByValue = _cellRange.GroupBy(x => x.Value)
                                      .Where(g => g.Count() > 0)
                                      .Select(y => new { Key = y.Key, Items = y });


                foreach (var _grup in _groupedByValue)
                {
                    var _disName = _grup.Key.ToString().Trim().ToUpper();

                    var _parentPlace = new Place
                    {
                        District = _disName,
                        State = "Johor",
                        DuplicateCount = _grup.Items.Count(),
                    };
                    

                    if (_parentPlace.District.Length == 3)
                    {
                        _parentPlace.GroupKey = _disName.Substring(0, 3);
                    }
                    else if (_parentPlace.District.Length == 4)
                    {
                        _parentPlace.GroupKey = _disName.Substring(0, 4);
                    }
                    else if (_parentPlace.District.Length >= 5)
                    {
                        _parentPlace.GroupKey = _disName.Substring(0, 5);
                    }


                    foreach (var _it in _grup.Items)
                    {

                        var _childPlace = new Place
                        {
                            District = _disName,
                            State = "Johor",
                            Address = _it.Address
                        };

                        _parentPlace.Children.Add(_childPlace);                       
                    }

                    _thePlaces.Add(_parentPlace);
                }

                var _groupedData = _thePlaces.GroupBy(p => p.GroupKey).OrderBy(g => g.Key).Where(q => q.Count() > 1);

                var _childStyle = this.Resources["MaterialDesignTreeViewItem"] as Style;

                foreach (var gd in _groupedData)
                {
                    var innerTVI = new TVIDistrictResolver(gd.Key.ToString());
                    innerTVI.DeleteListViewRequested += (s, e) =>
                    {
                        var _t = treeview_regionNameFixer.Items.OfType<TreeViewItem>().FirstOrDefault(tvi => tvi.Header.ToString() == gd.Key.ToString());
                        treeview_regionNameFixer.Items.Remove(_t);
                    };

                    gd.OrderBy(d => d.District).ToList().ForEach(x => innerTVI.Places.Add(x));

                    var _acc = "Cells:";
                    foreach (var pp in gd)
                    {
                        foreach (var ll in pp.Children)
                        {
                            _acc += $" {ll.Address}";
                        }
                    }

                    ToolTipService.SetToolTip(innerTVI, _acc);

                    var _tvi = new TreeViewItem { Header = gd.Key, ItemContainerStyle = _childStyle };
                    _tvi.Items.Add(innerTVI);

                    treeview_regionNameFixer.Items.Add(_tvi);
                }

                


                IsExportButtonEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            ScrollViewer scv = (ScrollViewer)sender;
            scv.ScrollToVerticalOffset(scv.VerticalOffset - e.Delta);
            e.Handled = true;
        }

        private void treeview_regionNameFixer_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            
        }

        private void expander_mainExcelCleaner_Expanded(object sender, RoutedEventArgs e)
        {
            expander_mainExcelCleaner.Header = "Hide Initials Control elements";
        }

        private void expander_mainExcelCleaner_Collapsed(object sender, RoutedEventArgs e)
        {
            expander_mainExcelCleaner.Header = "Show Initials Control elements";
        }
    }
}
