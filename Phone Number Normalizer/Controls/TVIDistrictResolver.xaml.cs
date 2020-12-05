using Phone_Number_Normalizer.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

namespace Phone_Number_Normalizer.Controls
{
    /// <summary>
    /// Interaction logic for TVIDistrictResolver.xaml
    /// </summary>
    public partial class TVIDistrictResolver : UserControl, INotifyPropertyChanged
    {
        public TVIDistrictResolver(string key)
        {
            InitializeComponent();
            DataContext = this;

            Key = key;
        }

        public string Key { get; }
        public ObservableCollection<Place> Places { get; set; } = new ObservableCollection<Place>();

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string propName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));



        public event EventHandler<string> ListViewItemSelectedCellsRequested;



        private void ListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string _acc = "Cells: ";
            foreach (var item in listview.SelectedItems.OfType<Place>())
            {
                foreach (var innerItem in item.Children)
                {
                    _acc += $" {innerItem.Address},";
                }
            }
            ListViewItemSelectedCellsRequested.Invoke(this, _acc);

            foreach (var ai in e.AddedItems)
            {
                cmbBox_potentialCandidates.Items.Add(ai);
            }

            foreach (var ri in e.RemovedItems)
            {
                cmbBox_potentialCandidates.Items.Remove(ri);
            }

            if (cmbBox_potentialCandidates.Items.Count > 0)
            {
                cmbBox_potentialCandidates.SelectedIndex = 0;
                btn_resolve.IsEnabled = true;
            }
            else
            {
                cmbBox_potentialCandidates.SelectedIndex = -1;
                btn_resolve.IsEnabled = false;
            }
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbBox_potentialCandidates.SelectedItem is Place _place)
            {
                btn_resolve.Content = $"Resolve as {_place.District}";
            }
        }

        private void cmbBox_potentialCandidates_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(cmbBox_potentialCandidates.Text))
            {
                btn_resolve.Content = $"Resolve as {cmbBox_potentialCandidates.Text}";
            }
        }

        public class TheEventArgs: EventArgs
        {
            public string ResolvedString { get; set; }
            public List<string> Addresses { get; set; }
        }

        public event EventHandler<string> DeleteListViewRequested;
        public event EventHandler<TheEventArgs> ResolveRequested;

        List<Place> places;
        private void btn_resolve_Click(object sender, RoutedEventArgs e)
        {
            var _resolvedString = cmbBox_potentialCandidates.Text;
            if (!string.IsNullOrEmpty(_resolvedString))
            {
                places = new List<Place>();
                foreach (var item in listview.SelectedItems.OfType<Place>())
                {
                    places.Add(item);
                }

                foreach (var nn in places)
                {
                    Places.Remove(nn);
                }
                OnPropertyChanged(nameof(Places));

                if (listview.Items.Count == 0)
                {
                    DeleteListViewRequested.Invoke(this, Key);
                }
            }

            List<Place> _outerItems = new List<Place>();
            foreach (var item in places)
            {
                _outerItems.Add(item);
                foreach (var innerItem in item.Children)
                {
                    _outerItems.Add(innerItem);
                }
            }

            var _strOnly = _outerItems.Select(s => s.Address);

            ResolveRequested.Invoke(this, new TheEventArgs { ResolvedString = _resolvedString, Addresses = _strOnly.ToList() });
        }
    }
}
