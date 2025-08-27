using System;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using Npgsql;
using System.Net.NetworkInformation;
using System.Threading.Tasks;

namespace WpfApp15
{
    public partial class Window1 : Window
    {
        private string connectionString = ConnectionString.Path;
        private DataTable originalDataTable;
        private string currentTable = "Clients";
        private const string ServerIp = "109.163.242.150";

        public Window1()
        {
            InitializeComponent();
            LoadClients();
            InitializeSearchFields();
        }

        private void InitializeSearchFields()
        {
            SearchField1.Items.Clear();
            SearchField2.Items.Clear();

            if (originalDataTable != null && originalDataTable.Columns.Count > 0)
            {
                foreach (DataColumn column in originalDataTable.Columns)
                {
                    SearchField1.Items.Add(column.ColumnName);
                    SearchField2.Items.Add(column.ColumnName);
                }

                if (SearchField1.Items.Count > 0)
                    SearchField1.SelectedIndex = 0;

                if (SearchField2.Items.Count > 0)
                    SearchField2.SelectedIndex = Math.Min(1, SearchField2.Items.Count - 1);
            }
        }

        private void LoadClients()
        {
            currentTable = "Clients";
            LoadData(
                @"SELECT 
                    client_id AS ""Идентификатор клиента"",
                    last_name AS ""Фамилия"",
                    first_name AS ""Имя"",
                    middle_name AS ""Отчество"",
                    phone_number AS ""Номер телефона"",
                    contract_id AS ""Номер договора"",
                    address AS ""Адрес""
                  FROM Clients",
                "Клиенты"
            );
        }

        private void LoadClients_Click(object sender, RoutedEventArgs e)
        {
            LoadClients();
        }

        private void LoadServiceCategories_Click(object sender, RoutedEventArgs e)
        {
            currentTable = "ServiceCategories";
            LoadData("SELECT service_id AS \"ID\", service_name AS \"Название услуги\", service_price AS \"Цена\", service_description AS \"Описание\" FROM Services", "Все услуги");
        }

        private void LoadAvailableEquipment_Click(object sender, RoutedEventArgs e)
        {
            currentTable = "AvailableEquipment";
            LoadData("SELECT equipment_id AS \"ID\", equipment_name AS \"Название оборудования\", equipment_price AS \"Цена\", quantity AS \"Количество в наличии\" FROM Equipment WHERE quantity > 0", "Доступное оборудование");
        }

        private void LoadServiceContracts_Click(object sender, RoutedEventArgs e)
        {
            currentTable = "ClientServices";
            string query = @"
                SELECT 
                    c.client_id AS ""ID клиента"",
                    CONCAT(c.last_name, ' ', c.first_name, ' ', c.middle_name) AS ""ФИО клиента"",
                    c.phone_number AS ""Телефон"",
                    COALESCE(s.service_name, 'Нет услуг') AS ""Название услуги"",
                    COALESCE(s.service_price::text, 'N/A') AS ""Цена услуги"",
                    COALESCE(s.service_description, 'N/A') AS ""Описание услуги"",
                    COALESCE(cc.contract_id::text, 'Нет договора') AS ""Номер договора"",
                    COALESCE(TO_CHAR(cc.contract_date, 'DD.MM.YYYY'), 'N/A') AS ""Дата договора"",
                    CASE 
                        WHEN cc.contract_id IS NULL THEN 'Нет договора'
                        WHEN cc.termination_date IS NULL THEN 'Активен'
                        WHEN cc.termination_date > CURRENT_DATE THEN 'Активен'
                        ELSE 'Завершен'
                    END AS ""Статус договора""
                FROM Clients c
                LEFT JOIN ClientServices cs ON c.client_id = cs.client_id
                LEFT JOIN Services s ON cs.service_id = s.service_id
                LEFT JOIN ClientContracts cc ON c.client_id = cc.client_id
                ORDER BY c.last_name, c.first_name, s.service_name";
            LoadData(query, "Услуги клиентов");
        }

        private void LoadServiceRequests_Click(object sender, RoutedEventArgs e)
        {
            currentTable = "ServiceRequests";
            LoadData(
                @"SELECT 
                    request_id AS ""Номер заявки"",
                    CONCAT(last_name, ' ', first_name, ' ', middle_name) AS ""ФИО клиента"",
                    phone_number AS ""Телефон"",
                    address AS ""Адрес"",
                    reason AS ""Причина"",
                    TO_CHAR(creation_date, 'DD.MM.YYYY HH24:MI') AS ""Дата создания"",
                    status AS ""Статус""
                  FROM ServiceRequests sr
                  JOIN Clients c ON sr.client_id = c.client_id
                  ORDER BY creation_date DESC",
                "Заявки на обслуживание"
            );
        }

        private void LoadData(string query, string title)
        {
            using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                    {
                        NpgsqlDataAdapter dataAdapter = new NpgsqlDataAdapter(command);
                        originalDataTable = new DataTable();
                        dataAdapter.Fill(originalDataTable);
                        dataGrid.ItemsSource = originalDataTable.DefaultView;

                        InitializeSearchFields();
                        UpdateSummary(originalDataTable);
                        DataGridTitle.Text = title;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка загрузки данных: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ShowContractStatus_Click(object sender, RoutedEventArgs e)
        {
            if (currentTable != "Clients")
            {
                ResultTextBlock.Text = "Пожалуйста, выберите таблицу клиентов.";
                return;
            }

            if (!(dataGrid.SelectedItem is DataRowView rowView) || originalDataTable.Rows.Count == 0)
            {
                ResultTextBlock.Text = "Нет данных о клиентах или не выбран клиент.";
                return;
            }

            if (!originalDataTable.Columns.Contains("Номер договора"))
            {
                ResultTextBlock.Text = "Столбец 'Номер договора' отсутствует в данных.";
                return;
            }

            if (rowView["Номер договора"] == DBNull.Value || string.IsNullOrEmpty(rowView["Номер договора"].ToString()))
            {
                ResultTextBlock.Text = "У клиента отсутствует договор.";
                return;
            }

            if (!int.TryParse(rowView["Номер договора"].ToString(), out int contractId))
            {
                ResultTextBlock.Text = "Некорректный номер договора.";
                return;
            }

            using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = @"
                        SELECT contract_id, contract_date, termination_date
                        FROM ClientContracts
                        WHERE contract_id = @contractId";
                    using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("contractId", contractId);
                        using (NpgsqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                string contractDate = reader["contract_date"] != DBNull.Value ? reader["contract_date"].ToString() : "Не указана";
                                string terminationDate = reader["termination_date"] != DBNull.Value ? reader["termination_date"].ToString() : "Не указан";
                                string status = "Активен";

                                if (reader["termination_date"] != DBNull.Value && DateTime.TryParse(terminationDate, out DateTime termDate))
                                {
                                    status = termDate < DateTime.Now ? "Завершен" : "Активен";
                                }

                                ResultTextBlock.Text = $"Договор №{contractId}\n" +
                                                       $"Дата заключения: {contractDate}\n" +
                                                       $"Дата окончания: {terminationDate}\n" +
                                                       $"Статус: {status}";
                            }
                            else
                            {
                                ResultTextBlock.Text = "Договор с указанным номером не найден.";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при получении статуса договора: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ShowClientEquipment_Click(object sender, RoutedEventArgs e)
        {
            if (currentTable != "Clients")
            {
                ResultTextBlock.Text = "Пожалуйста, выберите таблицу клиентов.";
                return;
            }

            if (!(dataGrid.SelectedItem is DataRowView rowView) || originalDataTable.Rows.Count == 0)
            {
                ResultTextBlock.Text = "Нет данных о клиентах или не выбран клиент.";
                return;
            }

            if (!originalDataTable.Columns.Contains("Идентификатор клиента"))
            {
                ResultTextBlock.Text = "Столбец 'Идентификатор клиента' отсутствует в данных.";
                return;
            }

            string clientIdStr = rowView["Идентификатор клиента"].ToString();
            if (!int.TryParse(clientIdStr, out int clientId))
            {
                ResultTextBlock.Text = "Некорректный идентификатор клиента.";
                return;
            }

            using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = @"
                        SELECT e.equipment_name, e.equipment_price
                        FROM Equipment e
                        JOIN Clients c ON e.equipment_id = c.equipment_id
                        WHERE c.client_id = @clientId";
                    using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("clientId", clientId);
                        using (NpgsqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                string equipmentName = reader["equipment_name"] != DBNull.Value ? reader["equipment_name"].ToString() : "Не указано";
                                string equipmentPrice = reader["equipment_price"] != DBNull.Value ? reader["equipment_price"].ToString() : "Не указано";

                                ResultTextBlock.Text = $"Оборудование клиента:\n" +
                                                       $"Наименование: {equipmentName}\n" +
                                                       $"Цена: {equipmentPrice} руб.";
                            }
                            else
                            {
                                ResultTextBlock.Text = "Оборудование для данного клиента не найдено.";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при получении данных об оборудовании: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async void CheckPing_Click(object sender, RoutedEventArgs e)
        {
            ResultTextBlock.Text = "Проверка ping...";
            try
            {
                using (Ping ping = new Ping())
                {
                    int successfulPings = 0;
                    int totalPings = 4;
                    long totalTime = 0;

                    for (int i = 0; i < totalPings; i++)
                    {
                        PingReply reply = await ping.SendPingAsync(ServerIp, 1000);
                        if (reply.Status == IPStatus.Success)
                        {
                            successfulPings++;
                            totalTime += reply.RoundtripTime;
                        }
                    }

                    double packetLoss = ((totalPings - successfulPings) / (double)totalPings) * 100;
                    double avgTime = successfulPings > 0 ? totalTime / successfulPings : 0;

                    ResultTextBlock.Text = $"Ping {ServerIp}:\n" +
                                           $"Отправлено: {totalPings}, Получено: {successfulPings}\n" +
                                           $"Потери: {packetLoss:F0}%\n" +
                                           $"Среднее время: {avgTime:F0} мс";
                }
            }
            catch (Exception ex)
            {
                ResultTextBlock.Text = $"Ошибка ping: {ex.Message}";
            }
        }

        private void ChangeRequestStatus_Click(object sender, RoutedEventArgs e)
        {
            if (currentTable != "ServiceRequests")
            {
                ResultTextBlock.Text = "Пожалуйста, выберите таблицу заявок на обслуживание.";
                return;
            }

            if (!(dataGrid.SelectedItem is DataRowView rowView))
            {
                ResultTextBlock.Text = "Выберите заявку для изменения статуса.";
                return;
            }

            try
            {
                int requestId = Convert.ToInt32(rowView["Номер заявки"]);
                string currentStatus = rowView["Статус"].ToString();
                string clientName = rowView["ФИО клиента"].ToString();

                // Создаем InputBox с ComboBox для выбора статуса
                InputBox statusBox = new InputBox($"Изменить статус заявки №{requestId} для {clientName}:", currentStatus);
                ComboBox statusCombo = new ComboBox
                {
                    ItemsSource = new[] { "Открыта", "В работе", "Закрыта" },
                    SelectedItem = currentStatus,
                    Margin = new Thickness(5),
                    FontSize = 14,
                    Width = 200
                };

                // Заменяем содержимое InputBox
                StackPanel panel = new StackPanel { Margin = new Thickness(10) };
                panel.Children.Add(new TextBlock { Text = statusBox.MessageTextBlock.Text, Margin = new Thickness(0, 0, 0, 10), FontSize = 16 });
                panel.Children.Add(statusCombo);
                panel.Children.Add(new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Margin = new Thickness(0, 10, 0, 0),
                    Children =
            {
                new Button
                {
                    Content = "OK",
                    Width = 75,
                    Margin = new Thickness(5),
                    Command = new RelayCommand(() =>
                    {
                        statusBox.DialogResult = true;
                        statusBox.Close();
                    })
                },
                new Button
                {
                    Content = "Отмена",
                    Width = 75,
                    Margin = new Thickness(5),
                    Command = new RelayCommand(() => statusBox.Close())
                }
            }
                });
                statusBox.Content = panel;

                if (statusBox.ShowDialog() != true || statusCombo.SelectedItem == null || statusCombo.SelectedItem.ToString() == currentStatus)
                {
                    ResultTextBlock.Text = statusCombo.SelectedItem?.ToString() == currentStatus
                        ? "Новый статус совпадает с текущим."
                        : "Изменение статуса отменено.";
                    return;
                }

                string newStatus = statusCombo.SelectedItem.ToString();

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    string query = "UPDATE ServiceRequests SET status = @status WHERE request_id = @requestId";
                    using (var cmd = new NpgsqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("status", newStatus);
                        cmd.Parameters.AddWithValue("requestId", requestId);
                        cmd.ExecuteNonQuery();
                    }

                    string auditQuery = "INSERT INTO audit_log (operation_type, table_name, changed_by, changed_at) VALUES (@operation, @table, @user, CURRENT_TIMESTAMP)";
                    using (var auditCmd = new NpgsqlCommand(auditQuery, conn))
                    {
                        auditCmd.Parameters.AddWithValue("operation", "UPDATE_STATUS");
                        auditCmd.Parameters.AddWithValue("table", "ServiceRequests");
                        auditCmd.Parameters.AddWithValue("user", Environment.UserName);
                        auditCmd.ExecuteNonQuery();
                    }
                }

                ResultTextBlock.Text = $"Статус заявки №{requestId} изменен на '{newStatus}'.";
                LoadServiceRequests_Click(sender, e);
            }
            catch (Exception ex)
            {
                ResultTextBlock.Text = $"Ошибка при изменении статуса: {ex.Message}";
            }
        }

        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            if (originalDataTable == null || originalDataTable.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для поиска.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string field1 = SearchField1.SelectedItem?.ToString();
            string value1 = SearchValue1.Text?.Trim();
            string field2 = SearchField2.SelectedItem?.ToString();
            string value2 = SearchValue2.Text?.Trim();

            DataView dv = originalDataTable.DefaultView;
            string filter = "";

            if (!string.IsNullOrEmpty(field1) && !string.IsNullOrEmpty(value1))
            {
                string fieldType = GetColumnType(originalDataTable, field1);
                if (fieldType == "System.String")
                {
                    filter = $"[{field1}] LIKE '%{value1}%'";
                }
                else if (fieldType == "System.Int32" || fieldType == "System.Int64" || fieldType == "System.Decimal")
                {
                    if (int.TryParse(value1, out _))
                    {
                        filter = $"[{field1}] = {value1}";
                    }
                    else
                    {
                        MessageBox.Show($"Значение для поля {field1} должно быть числом.", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }
                else if (fieldType == "System.DateTime")
                {
                    if (DateTime.TryParse(value1, out _))
                    {
                        filter = $"[{field1}] = '{value1}'";
                    }
                    else
                    {
                        MessageBox.Show($"Значение для поля {field1} должно быть датой.", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }
            }

            if (!string.IsNullOrEmpty(field2) && !string.IsNullOrEmpty(value2))
            {
                string fieldType = GetColumnType(originalDataTable, field2);
                string filterPart = "";

                if (fieldType == "System.String")
                {
                    filterPart = $"[{field2}] LIKE '%{value2}%'";
                }
                else if (fieldType == "System.Int32" || fieldType == "System.Int64" || fieldType == "System.Decimal")
                {
                    if (int.TryParse(value2, out _))
                    {
                        filterPart = $"[{field2}] = {value2}";
                    }
                    else
                    {
                        MessageBox.Show($"Значение для поля {field2} должно быть числом.", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }
                else if (fieldType == "System.DateTime")
                {
                    if (DateTime.TryParse(value2, out _))
                    {
                        filterPart = $"[{field2}] = '{value2}'";
                    }
                    else
                    {
                        MessageBox.Show($"Значение для поля {field2} должно быть датой.", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }

                if (!string.IsNullOrEmpty(filter))
                {
                    filter += " AND " + filterPart;
                }
                else
                {
                    filter = filterPart;
                }
            }

            try
            {
                dv.RowFilter = filter;
                dataGrid.ItemsSource = dv;

                DataTable filteredTable = dv.ToTable();
                UpdateSummary(filteredTable);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при фильтрации данных: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                dv.RowFilter = "";
                dataGrid.ItemsSource = dv;
                UpdateSummary(originalDataTable);
            }
        }

        private string GetColumnType(DataTable table, string columnName)
        {
            if (table.Columns.Contains(columnName))
            {
                return table.Columns[columnName].DataType.ToString();
            }
            return "System.String";
        }

        private void ResetSearch_Click(object sender, RoutedEventArgs e)
        {
            SearchValue1.Text = "";
            SearchValue2.Text = "";

            if (originalDataTable != null)
            {
                DataView dv = originalDataTable.DefaultView;
                dv.RowFilter = "";
                dataGrid.ItemsSource = dv;
                UpdateSummary(originalDataTable);
            }
        }

        private void UpdateSummary(DataTable dataTable)
        {
            RecordCountText.Text = dataTable.Rows.Count.ToString();

            DataColumn numericColumn = null;
            foreach (DataColumn column in dataTable.Columns)
            {
                Type columnType = column.DataType;
                if (columnType == typeof(long) ||
                    columnType == typeof(decimal) || columnType == typeof(double))
                {
                    numericColumn = column;
                    break;
                }
            }

            if (numericColumn != null)
            {
                decimal sum = 0;
                int count = 0;

                foreach (DataRow row in dataTable.Rows)
                {
                    if (!row.IsNull(numericColumn))
                    {
                        sum += Convert.ToDecimal(row[numericColumn]);
                        count++;
                    }
                }

                SummaryLabel1.Text = $"Сумма ({numericColumn.ColumnName}):";
                SummaryValue1.Text = sum.ToString("N2");

                SummaryLabel2.Text = $"Среднее ({numericColumn.ColumnName}):";
                SummaryValue2.Text = count > 0 ? (sum / count).ToString("N2") : "0";
            }
            else
            {
                SummaryLabel1.Text = "Статистика:";
                SummaryValue1.Text = "Нет числовых полей";

                SummaryLabel2.Text = "Дата обновления:";
                SummaryValue2.Text = DateTime.Now.ToString("dd.MM.yyyy HH:mm");
            }
        }

        private void ExportClients_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel файлы (*.xlsx)|*.xlsx",
                FileName = "Данные.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    DataView view = (DataView)dataGrid.ItemsSource;
                    DataTable table = view.ToTable();

                    using (var workbook = new ClosedXML.Excel.XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Данные");

                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            worksheet.Cell(1, i + 1).Value = table.Columns[i].ColumnName;
                        }

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                worksheet.Cell(i + 2, j + 1).Value = table.Rows[i][j]?.ToString() ?? string.Empty;
                            }
                        }

                        workbook.SaveAs(saveFileDialog.FileName);
                    }

                    MessageBox.Show("Данные успешно экспортированы в Excel.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка экспорта данных: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }
    }
}