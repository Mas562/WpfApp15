using LiveCharts.Wpf;
using LiveCharts;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ClosedXML.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace WpfApp15
{
    public partial class WindowChart : Window
    {
        public SeriesCollection Payments { get; set; }
        public List<string> Months { get; set; }

        private string connectionString = ConnectionString.Path;
        private List<(DateTime date, decimal amount)> rawData;

        public WindowChart()
        {
            InitializeComponent();
            Payments = new SeriesCollection(); // Инициализация Payments
            Months = new List<string>(); // Инициализация Months
            DataContext = this;

            // Заполняем список годов и загружаем данные
            PopulateYears();
            LoadRawData();
        }

        private void PopulateYears()
        {
            try
            {
                using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();
                    string query = "SELECT DISTINCT EXTRACT(YEAR FROM date) AS year FROM Charges ORDER BY year;";
                    using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                    using (NpgsqlDataReader reader = command.ExecuteReader())
                    {
                        YearComboBox.Items.Add("Все");
                        while (reader.Read())
                        {
                            YearComboBox.Items.Add(reader.GetInt32(0).ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки годов: " + ex.Message);
            }
        }

        private void LoadRawData()
        {
            if (rawData != null && rawData.Count > 0)
            {
                return;
            }

            rawData = new List<(DateTime date, decimal amount)>();
            try
            {
                using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();
                    string query = "SELECT ch.date, p.amount FROM Charges ch JOIN Payments p ON ch.charge_id = p.payment_id;";
                    using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                    using (NpgsqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            rawData.Add((reader.GetDateTime(0), reader.GetDecimal(1)));
                        }
                    }
                }

                if (rawData.Count == 0)
                {
                    MessageBox.Show("Данные отсутствуют. Проверьте базу данных.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки данных: " + ex.Message);
            }
        }

        private void FilterChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilter();
        }

        private void ApplyFilter_Click(object sender, RoutedEventArgs e)
        {
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            string selectedYear = YearComboBox.SelectedItem?.ToString();
            string selectedMonth = MonthComboBox.SelectedItem is ComboBoxItem monthItem && monthItem.Content.ToString() != "Все"
                ? monthItem.Content.ToString()
                : null;

            LoadChartData(selectedYear, selectedMonth);
        }

        private void LoadChartData(string year, string month)
        {
            if (rawData == null || rawData.Count == 0)
            {
                LoadRawData();
                if (rawData.Count == 0)
                {
                    return;
                }
            }

            try
            {
                var filteredData = rawData.Where(d =>
                    (string.IsNullOrEmpty(year) || year == "Все" || d.date.Year.ToString() == year) &&
                    (string.IsNullOrEmpty(month) || d.date.Month == MonthToNumber(month)));

                if (!filteredData.Any())
                {
                    MessageBox.Show("Нет данных для отображения. Попробуйте изменить фильтры.");
                    return;
                }

                var groupedData = filteredData
                    .GroupBy(d => d.date.Month)
                    .Select(g => new { Month = g.Key, TotalAmount = g.Sum(d => d.amount) })
                    .OrderBy(g => g.Month);

                if (Months == null)
                {
                    Months = new List<string>();
                }
                Months.Clear();

                ChartValues<double> values = new ChartValues<double>();

                foreach (var item in groupedData)
                {
                    Months.Add(MonthNumberToName(item.Month));
                    values.Add((double)item.TotalAmount);
                }

                if (Payments == null)
                {
                    Payments = new SeriesCollection();
                }
                Payments.Clear();

                Payments.Add(new ColumnSeries
                {
                    Title = "Сумма платежей",
                    Values = values,
                    Fill = System.Windows.Media.Brushes.SkyBlue
                });

                if (cartesianChart == null)
                {
                   // MessageBox.Show("Ошибка: элемент графика не найден. Проверьте XAML-файл.");
                    return;
                }

                cartesianChart.AxisX.Clear();
                cartesianChart.AxisX.Add(new Axis
                {
                    Title = "Месяц",
                    Labels = Months
                });

                cartesianChart.AxisY.Clear();
                cartesianChart.AxisY.Add(new Axis
                {
                    Title = "Сумма (руб.)",
                    LabelFormatter = value => value.ToString("N2")
                });

                cartesianChart.Series = Payments;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        private void GoBack_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ExportChartWithExcelChart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Months == null || Months.Count == 0 || Payments == null || Payments.Count == 0)
                {
                    MessageBox.Show("Нет данных для экспорта.");
                    return;
                }

                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Сохранить график",
                    FileName = "График_платежей.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    // Создаем приложение Excel
                    var excelApp = new Excel.Application();
                    var workbook = excelApp.Workbooks.Add();
                    var worksheet = workbook.ActiveSheet as Excel.Worksheet;

                    // Заполнение таблицы данными
                    worksheet.Cells[1, 1] = "Месяц";
                    worksheet.Cells[1, 2] = "Сумма платежей";

                    for (int i = 0; i < Months.Count; i++)
                    {
                        worksheet.Cells[i + 2, 1] = Months[i];
                        worksheet.Cells[i + 2, 2] = Convert.ToDouble(((ColumnSeries)Payments[0]).Values[i]);
                    }

                    // Добавление диаграммы
                    var charts = worksheet.ChartObjects() as Excel.ChartObjects;
                    var chartObject = charts.Add(300, 50, 500, 300) as Excel.ChartObject; // Координаты и размеры диаграммы
                    var chart = chartObject.Chart;

                    // Настройка диапазона данных для диаграммы
                    var dataRange = worksheet.Range[
                        worksheet.Cells[1, 1], // Заголовок "Месяц"
                        worksheet.Cells[Months.Count + 1, 2] // Последний ряд данных "Сумма платежей"
                    ];
                    chart.SetSourceData(dataRange);

                    // Установка типа диаграммы
                    chart.ChartType = Excel.XlChartType.xlColumnClustered; // Столбчатая диаграмма

                    // Настройка заголовка диаграммы и легенды
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = "График платежей";
                    chart.HasLegend = true;
                    chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;

                    // Сохранение файла
                    workbook.SaveAs(saveFileDialog.FileName);
                    workbook.Close();
                    excelApp.Quit();

                    MessageBox.Show("График успешно экспортирован в Excel.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при экспорте графика: " + ex.Message);
            }
        }



        private void ExportChart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Months == null || Months.Count == 0 || Payments == null || Payments.Count == 0)
                {
                    MessageBox.Show("Нет данных для экспорта.");
                    return;
                }

                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Сохранить график",
                    FileName = "График_платежей.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    using (var workbook = new ClosedXML.Excel.XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("График платежей");

                        // Заголовки
                        worksheet.Cell(1, 1).Value = "Месяц";
                        worksheet.Cell(1, 2).Value = "Сумма платежей";

                        // Данные
                        for (int i = 0; i < Months.Count; i++)
                        {
                            worksheet.Cell(i + 2, 1).Value = Months[i]; // Месяц в текстовом формате
                            worksheet.Cell(i + 2, 2).Value = Convert.ToDouble(((ColumnSeries)Payments[0]).Values[i]); // Значение в формате double
                        }

                        workbook.SaveAs(saveFileDialog.FileName);
                        MessageBox.Show("График успешно экспортирован.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при экспорте графика: " + ex.Message);
            }
        }


        private int MonthToNumber(string monthName)
        {
            var months = new[] {
                "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
            };
            return Array.IndexOf(months, monthName) + 1;
        }

        private string MonthNumberToName(int month)
        {
            var months = new[] {
                "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
            };
            return months[month - 1];
        }
    }
}
