using System.Collections.Generic;
using System.Windows;

namespace WpfApp15
{
    public partial class TariffSelectionWindow : Window
    {
        public Window3.Service SelectedService { get; private set; }

        public TariffSelectionWindow(List<Window3.Service> services)
        {
            InitializeComponent();
            TariffsListView.ItemsSource = services;
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            if (TariffsListView.SelectedItem is Window3.Service selected)
            {
                SelectedService = selected;
                DialogResult = true;
            }
            else
            {
                MessageBox.Show("Выберите тариф из списка", "Внимание",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}