using System;
using System.Data;
using System.Windows;
using Microsoft.Win32;
using Npgsql;
using System.IO;
using Xceed.Words.NET;
using System.Linq;
using System.Collections.Generic;

namespace WpfApp15
{
    public partial class Window2 : Window
    {
        private string connectionString = ConnectionString.Path;
        private DataTable originalDataTable;

        public Window2()
        {
            InitializeComponent();
            LoadClientDetails();
            InitializeSearchFields();
            DataGridHeader.Text = "Информация о клиентах";
        }

        private void LoadClientDetails()
        {
            DataGridHeader.Text = "Информация о клиентах";
            string sql = @"
                SELECT 
                    client_id AS ""ID"",
                    last_name AS ""Фамилия"",
                    first_name AS ""Имя"",
                    middle_name AS ""Отчество"",
                    phone_number AS ""Телефон"",
                    address AS ""Адрес""
                FROM Clients
                ORDER BY last_name, first_name";
            FillGrid(sql);
        }

        private void LoadClientDetails_Click(object sender, RoutedEventArgs e)
        {
            LoadClientDetails();
        }

        private void LoadServiceContracts_Click(object sender, RoutedEventArgs e)
        {
            DataGridHeader.Text = "Договоры клиентов";
            string sql = @"
        SELECT 
            cc.contract_id AS ""Номер договора"",
            TO_CHAR(cc.contract_date, 'DD.MM.YYYY') AS ""Дата договора"",
            CONCAT(c.last_name, ' ', c.first_name, ' ', c.middle_name) AS ""ФИО клиента"",
            c.address AS ""Адрес клиента"",
            CASE 
                WHEN cc.termination_date IS NULL THEN 'Активен'
                WHEN cc.termination_date > CURRENT_DATE THEN 'Активен'
                ELSE 'Завершен'
            END AS ""Статус"",
            CASE 
                WHEN cc.termination_date IS NULL THEN 'Не указана'
                ELSE TO_CHAR(cc.termination_date, 'DD.MM.YYYY')
            END AS ""Дата завершения""
        FROM ClientContracts cc
        JOIN Clients c ON cc.client_id = c.client_id
        ORDER BY cc.contract_date DESC";
            FillGrid(sql);
        }

        private void LoadClientPayments_Click(object sender, RoutedEventArgs e)
        {
            DataGridHeader.Text = "Платежи клиентов";
            DateTime startDate = new DateTime(2023, 1, 1);
            DateTime endDate = DateTime.Now;

            string sql = @"
                SELECT 
                    CONCAT(c.last_name, ' ', c.first_name, ' ', c.middle_name) AS ""ФИО"",
                    c.address AS ""Адрес"",
                    gp.amount AS ""Сумма"",
                    TO_CHAR(gp.payment_date, 'DD.MM.YYYY') AS ""Дата""
                FROM get_payments_in_period(@startDate::DATE, @endDate::DATE) gp
                JOIN Clients c ON gp.client_id = c.client_id
                ORDER BY gp.payment_date DESC";

            try
            {
                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("startDate", startDate);
                        cmd.Parameters.AddWithValue("endDate", endDate);

                        var adapter = new NpgsqlDataAdapter(cmd);
                        var dt = new DataTable();
                        adapter.Fill(dt);
                        originalDataTable = dt;
                        dataGrid.ItemsSource = dt.DefaultView;
                        InitializeSearchFields();
                        UpdateSummary(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке платежей: " + ex.Message,
                                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void FillGrid(string sql)
        {
            try
            {
                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand(sql, conn))
                    {
                        var adapter = new NpgsqlDataAdapter(cmd);
                        var dt = new DataTable();
                        adapter.Fill(dt);
                        originalDataTable = dt;
                        dataGrid.ItemsSource = dt.DefaultView;
                        InitializeSearchFields();
                        UpdateSummary(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке данных: " + ex.Message,
                                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            if (originalDataTable == null || originalDataTable.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для поиска.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string f1 = SearchField1.SelectedItem as string;
            string v1 = SearchValue1.Text.Trim();
            string f2 = SearchField2.SelectedItem as string;
            string v2 = SearchValue2.Text.Trim();

            var dv = originalDataTable.DefaultView;
            string filter = BuildFilter(originalDataTable, f1, v1);
            string filter2 = BuildFilter(originalDataTable, f2, v2);

            if (!string.IsNullOrEmpty(filter) && !string.IsNullOrEmpty(filter2))
                filter += " AND " + filter2;
            else if (string.IsNullOrEmpty(filter))
                filter = filter2;

            try
            {
                dv.RowFilter = filter;
                dataGrid.ItemsSource = dv;
                UpdateSummary(dv.ToTable());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при фильтрации: " + ex.Message,
                                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                dv.RowFilter = "";
                dataGrid.ItemsSource = dv;
                UpdateSummary(originalDataTable);
            }
        }

        private string BuildFilter(DataTable table, string column, string value)
        {
            if (string.IsNullOrEmpty(column) || string.IsNullOrEmpty(value))
                return "";

            var type = table.Columns[column].DataType;
            if (type == typeof(string))
                return $"[{column}] LIKE '%{value}%'";
            if (type == typeof(int) || type == typeof(long) || type == typeof(decimal))
                return int.TryParse(value, out _) ? $"[{column}] = {value}"
                                               : "";
            if (type == typeof(DateTime))
                return DateTime.TryParse(value, out _)
                    ? $"[{column}] = '{value}'"
                    : "";
            return "";
        }

        private void ResetSearch_Click(object sender, RoutedEventArgs e)
        {
            SearchValue1.Text = "";
            SearchValue2.Text = "";
            if (originalDataTable != null)
            {
                var dv = originalDataTable.DefaultView;
                dv.RowFilter = "";
                dataGrid.ItemsSource = dv;
                UpdateSummary(originalDataTable);
            }
        }

        private void InitializeSearchFields()
        {
            SearchField1.Items.Clear();
            SearchField2.Items.Clear();
            if (originalDataTable == null) return;

            foreach (DataColumn col in originalDataTable.Columns)
            {
                SearchField1.Items.Add(col.ColumnName);
                SearchField2.Items.Add(col.ColumnName);
            }

            if (SearchField1.Items.Count > 0) SearchField1.SelectedIndex = 0;
            if (SearchField2.Items.Count > 1) SearchField2.SelectedIndex = 1;
        }

        private void UpdateSummary(DataTable dt)
        {
            RecordCountText.Text = dt.Rows.Count.ToString();

            DataColumn numCol = null;
            foreach (DataColumn c in dt.Columns)
            {
                if (c.DataType == typeof(decimal) ||
                    c.DataType == typeof(double) ||
                    c.DataType == typeof(long))
                {
                    numCol = c;
                    break;
                }
            }

            if (numCol != null)
            {
                decimal sum = 0;
                int count = 0;
                foreach (DataRow r in dt.Rows)
                {
                    if (!r.IsNull(numCol))
                    {
                        sum += Convert.ToDecimal(r[numCol]);
                        count++;
                    }
                }
                SummaryLabel1.Text = $"Сумма ({numCol.ColumnName}):";
                SummaryValue1.Text = sum.ToString("N2");
                SummaryLabel2.Text = $"Среднее ({numCol.ColumnName}):";
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

        private void AddClient_Click(object sender, RoutedEventArgs e)
        {
            string last = LastNameTextBox.Text.Trim();
            string first = FirstNameTextBox.Text.Trim();
            string mid = MiddleNameTextBox.Text.Trim();
            string phone = PhoneNumberTextBox.Text.Trim();
            string addr = AddressTextBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(last) ||
                string.IsNullOrWhiteSpace(first) ||
                string.IsNullOrWhiteSpace(phone))
            {
                MessageBox.Show("ФИО и телефон обязательны.");
                return;
            }

            try
            {
                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    string ins = @"
                        INSERT INTO Clients 
                            (last_name, first_name, middle_name, phone_number, address)
                        VALUES
                            (@ln, @fn, @mn, @ph, @ad)";
                    using (var cmd = new NpgsqlCommand(ins, conn))
                    {
                        cmd.Parameters.AddWithValue("ln", last);
                        cmd.Parameters.AddWithValue("fn", first);
                        cmd.Parameters.AddWithValue("mn", mid);
                        cmd.Parameters.AddWithValue("ph", phone);
                        cmd.Parameters.AddWithValue("ad", addr);
                        cmd.ExecuteNonQuery();
                    }
                    MessageBox.Show("Клиент добавлен.");
                    ClearFields();
                    LoadClientDetails();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при добавлении: " + ex.Message);
            }
        }

        private void ClearFields()
        {
            LastNameTextBox.Text = "";
            FirstNameTextBox.Text = "";
            MiddleNameTextBox.Text = "";
            PhoneNumberTextBox.Text = "";
            AddressTextBox.Text = "";
        }

        private void AddContract_Click(object sender, RoutedEventArgs e)
        {
            var selectWindow = new SelectClientEmployeeWindow();
            selectWindow.Owner = this;
            if (selectWindow.ShowDialog() == true)
            {
                MessageBox.Show("Договор успешно создан.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                LoadClientDetails(); // Обновляем список клиентов
            }
        }

        private void EditClient_Click(object sender, RoutedEventArgs e)
        {
            if (!(dataGrid.SelectedItem is DataRowView rowView))
            {
                MessageBox.Show("Выберите клиента для редактирования.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string clientId = rowView["ID"].ToString();
            string newLastName = PromptInput("Введите новую фамилию:", rowView["Фамилия"].ToString());
            string newFirstName = PromptInput("Введите новое имя:", rowView["Имя"].ToString());
            string newPhone = PromptInput("Введите новый номер телефона:", rowView["Телефон"].ToString());

            using (var conn = new NpgsqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    string query = "UPDATE Clients SET last_name = @lastName, first_name = @firstName, phone_number = @phone WHERE client_id = @clientId";
                    using (var cmd = new NpgsqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("lastName", newLastName);
                        cmd.Parameters.AddWithValue("firstName", newFirstName);
                        cmd.Parameters.AddWithValue("phone", newPhone);
                        cmd.Parameters.AddWithValue("clientId", int.Parse(clientId));
                        cmd.ExecuteNonQuery();
                    }
                    MessageBox.Show("Клиент успешно обновлен.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    LoadClientDetails();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка редактирования клиента: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void DeleteClient_Click(object sender, RoutedEventArgs e)
        {
            if (!(dataGrid.SelectedItem is DataRowView rowView))
            {
                MessageBox.Show("Выберите клиента для удаления.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string clientId = rowView["ID"].ToString();

            if (MessageBox.Show("Вы уверены, что хотите удалить клиента?", "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                using (var conn = new NpgsqlConnection(connectionString))
                {
                    try
                    {
                        conn.Open();
                        string query = "DELETE FROM Clients WHERE client_id = @clientId";
                        using (var cmd = new NpgsqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("clientId", int.Parse(clientId));
                            cmd.ExecuteNonQuery();
                        }
                        MessageBox.Show("Клиент успешно удален.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                        LoadClientDetails();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка удаления клиента: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void ExportClients_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel файлы (*.xlsx)|*.xlsx",
                FileName = "Клиенты.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    DataView view = (DataView)dataGrid.ItemsSource;
                    DataTable table = view.ToTable();

                    using (var workbook = new ClosedXML.Excel.XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Клиенты");

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

        private void ShowChartButton_Click(object sender, RoutedEventArgs e)
        {
            WindowChart chartWindow = new WindowChart();
            chartWindow.Show();
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            new MainWindow().Show();
            Close();
        }
        private void CreateServiceRequest_Click(object sender, RoutedEventArgs e)
        {
            if (!(dataGrid.SelectedItem is DataRowView rowView))
            {
                MessageBox.Show("Выберите клиента для создания заявки.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                int clientId = Convert.ToInt32(rowView["ID"]);
                string fullName = $"{rowView["Фамилия"]} {rowView["Имя"]} {rowView["Отчество"]}".Trim();

                // Используем ваш InputBox для ввода причины
                InputBox inputBox = new InputBox("Введите причину заявки:", "");
                if (inputBox.ShowDialog() != true || string.IsNullOrWhiteSpace(inputBox.Result))
                {
                    MessageBox.Show("Причина заявки не указана.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string reason = inputBox.Result;

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    string query = @"
                INSERT INTO ServiceRequests (client_id, reason, creation_date, status)
                VALUES (@clientId, @reason, CURRENT_TIMESTAMP, 'Открыта')
                RETURNING request_id";

                    using (var cmd = new NpgsqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("clientId", clientId);
                        cmd.Parameters.AddWithValue("reason", reason);
                        int requestId = (int)cmd.ExecuteScalar();

                        MessageBox.Show($"Заявка №{requestId} для клиента {fullName} успешно создана.",
                                       "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }

                LoadClientDetails(); // Обновляем данные
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании заявки: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private string PromptInput(string message, string defaultValue)
        {
            InputBox inputBox = new InputBox(message, defaultValue);
            return inputBox.ShowDialog() == true ? inputBox.Result : defaultValue;
        }

        // Формирование договора в Word
        private void GenerateContract_Click(object sender, RoutedEventArgs e)
        {
            // Проверка, выбран ли клиент
            if (!(dataGrid.SelectedItem is DataRowView rowView))
            {
                MessageBox.Show("Выберите клиента для формирования договора.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                // Извлечение данных клиента из выбранной строки
                int clientId = Convert.ToInt32(rowView["ID"]);
                string lastName = rowView["Фамилия"].ToString();
                string firstName = rowView["Имя"].ToString();
                string middleName = rowView["Отчество"].ToString();
                string phone = rowView["Телефон"].ToString();
                string address = rowView["Адрес"].ToString();
                string fullName = $"{lastName} {firstName} {middleName}".Trim();

                // Диалог для выбора места сохранения договора
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word документы (*.docx)|*.docx",
                    FileName = $"Договор_{lastName}_{firstName}.docx"
                };

                if (saveFileDialog.ShowDialog() != true)
                    return;

                // Формирование договора
                GenerateWordContract(clientId, fullName, phone, address, saveFileDialog.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при формировании договора: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GenerateWordContract(int clientId, string fullName, string phone, string address, string outputPath)
        {
            try
            {
                // Путь к шаблону договора
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "ContractTemplate.docx");

                // Проверка наличия шаблона
                if (!File.Exists(templatePath))
                {
                    MessageBox.Show("Шаблон договора не найден. Проверьте наличие файла по пути: " + templatePath,
                                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Данные договора и клиента
                string contractNumber = "";
                string contractDate = DateTime.Now.ToString("dd.MM.yyyy");
                string terminationDate = "Не указана";
                string city = "Москва"; // Статическое значение, можно сделать настраиваемым
                string executorName = "ООО 'ТехноПровайдер'";
                string executorInnKpp = "1234567890/123456789";
                string executorOgrn = "1234567890123";
                string executorAddress = "г. Москва, ул. Примерная, д. 1";
                string executorPhone = "+7 (999) 123-45-67";
                string executorEmail = "info@technoprovider.ru";
                string executorRepresentative = "Генерального директора Иванова И.И.";
                string executorBasis = "Устава";
                string clientInn = "Не указан"; // Нет в базе, можно запросить у пользователя
                string clientEmail = "client@example.com"; // Нет в базе, можно запросить
                string internetSpeed = "100"; // Можно добавить в Services.service_description
                string additionalServices = "Нет";
                string serviceLocation = address; // Используем адрес клиента
                string connectionDays = "5"; // Статическое значение
                string monthlyFee = "0";
                string connectionFee = "0"; // Статическое значение
                string paymentDay = "10"; // Статическое значение
                string penaltyRate = "0.1"; // Статическое значение

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();

                    // Получение номера договора и дат
                    string query = @"
                SELECT 
                    contract_id,
                    TO_CHAR(contract_date, 'DD.MM.YYYY') AS contract_date,
                    TO_CHAR(termination_date, 'DD.MM.YYYY') AS termination_date
                FROM ClientContracts
                WHERE client_id = @clientId
                ORDER BY contract_date DESC
                LIMIT 1";

                    using (var cmd = new NpgsqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("clientId", clientId);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                contractNumber = reader["contract_id"].ToString();
                                contractDate = reader["contract_date"].ToString();
                                if (!reader.IsDBNull(reader.GetOrdinal("termination_date")))
                                    terminationDate = reader["termination_date"].ToString();
                            }
                            else
                            {
                                // Создание нового договора
                                reader.Close();
                                var rnd = new Random();
                                int num = rnd.Next(100000, 999999);
                                string insQuery = @"
                            INSERT INTO ClientContracts
                                (client_id, contract_date, contract_id)
                            VALUES
                                (@cid, CURRENT_DATE, @num)
                            RETURNING contract_id, TO_CHAR(CURRENT_DATE, 'DD.MM.YYYY') AS contract_date";

                                using (var cmdIns = new NpgsqlCommand(insQuery, conn))
                                {
                                    cmdIns.Parameters.AddWithValue("cid", clientId);
                                    cmdIns.Parameters.AddWithValue("num", num);
                                    using (var readerIns = cmdIns.ExecuteReader())
                                    {
                                        if (readerIns.Read())
                                        {
                                            contractNumber = readerIns["contract_id"].ToString();
                                            contractDate = readerIns["contract_date"].ToString();
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Получение списка услуг клиента
                    List<string> servicesList = new List<string>();
                    string queryServices = @"
                SELECT s.service_name, s.service_price, s.service_description
                FROM ClientServices cs
                JOIN Services s ON cs.service_id = s.service_id
                WHERE cs.client_id = @clientId";

                    using (var cmd = new NpgsqlCommand(queryServices, conn))
                    {
                        cmd.Parameters.AddWithValue("clientId", clientId);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (!reader.HasRows)
                            {
                                servicesList.Add("Услуги не указаны");
                            }
                            else
                            {
                                while (reader.Read())
                                {
                                    string serviceName = reader["service_name"].ToString();
                                    decimal servicePrice = reader.GetDecimal(reader.GetOrdinal("service_price"));
                                    servicesList.Add($"{serviceName} - {servicePrice:N2} руб.");
                                    // Проверяем service_description на наличие скорости интернета
                                    string description = reader["service_description"].ToString();
                                    if (description.Contains("Мбит/с"))
                                    {
                                        var match = System.Text.RegularExpressions.Regex.Match(description, @"\d+");
                                        if (match.Success) internetSpeed = match.Value;
                                    }
                                }
                            }
                        }
                    }

                    // Рассчет абонентской платы
                    decimal totalPrice = servicesList
                        .Where(s => s != "Услуги не указаны")
                        .Select(s => decimal.Parse(s.Split('-')[1].Replace(" руб.", "").Trim()))
                        .Sum();
                    monthlyFee = totalPrice.ToString("N2");
                    additionalServices = string.Join("; ", servicesList);
                }

                // Копирование шаблона
                File.Copy(templatePath, outputPath, true);

                // Заполнение документа
                using (var doc = DocX.Load(outputPath))
                {
                    // Замена placeholders
                    doc.ReplaceText("{{CONTRACT_NUMBER}}", contractNumber);
                    doc.ReplaceText("{{CITY}}", city);
                    doc.ReplaceText("{{CONTRACT_DATE}}", contractDate);
                    doc.ReplaceText("{{EXECUTOR_NAME}}", executorName);
                    doc.ReplaceText("{{EXECUTOR_INN_KPP}}", executorInnKpp);
                    doc.ReplaceText("{{EXECUTOR_OGRN}}", executorOgrn);
                    doc.ReplaceText("{{EXECUTOR_ADDRESS}}", executorAddress);
                    doc.ReplaceText("{{EXECUTOR_PHONE}}", executorPhone);
                    doc.ReplaceText("{{EXECUTOR_EMAIL}}", executorEmail);
                    doc.ReplaceText("{{EXECUTOR_REPRESENTATIVE}}", executorRepresentative);
                    doc.ReplaceText("{{EXECUTOR_BASIS}}", executorBasis);
                    doc.ReplaceText("{{CLIENT_NAME}}", fullName);
                    doc.ReplaceText("{{CLIENT_INN}}", clientInn);
                    doc.ReplaceText("{{CLIENT_ADDRESS}}", address);
                    doc.ReplaceText("{{CLIENT_PHONE}}", phone);
                    doc.ReplaceText("{{CLIENT_EMAIL}}", clientEmail);
                    doc.ReplaceText("{{INTERNET_SPEED}}", internetSpeed);
                    doc.ReplaceText("{{ADDITIONAL_SERVICES}}", additionalServices);
                    doc.ReplaceText("{{SERVICE_LOCATION}}", serviceLocation);
                    doc.ReplaceText("{{CONNECTION_DAYS}}", connectionDays);
                    doc.ReplaceText("{{MONTHLY_FEE}}", monthlyFee);
                    doc.ReplaceText("{{CONNECTION_FEE}}", connectionFee);
                    doc.ReplaceText("{{PAYMENT_DAY}}", paymentDay);
                    doc.ReplaceText("{{PENALTY_RATE}}", penaltyRate);
                    doc.ReplaceText("{{TERMINATION_DATE}}", terminationDate);

                    // Сохранение документа
                    doc.Save();
                }

                MessageBox.Show($"Договор успешно сформирован и сохранен по пути:\n{outputPath}",
                               "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

                // Попытка открыть документ
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = outputPath,
                        UseShellExecute = true
                    });
                }
                catch
                {
                    MessageBox.Show("Не удалось автоматически открыть документ. Документ сохранен, вы можете открыть его вручную.",
                                   "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при формировании договора: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}