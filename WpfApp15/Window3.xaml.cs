using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Windows.Media.Imaging;
 
using ClosedXML.Report.Utils;
using System.Net.Http.Headers;
using System.Net.NetworkInformation;
using System.Linq;
using System.Timers;
using System.IO;
using Npgsql;

namespace WpfApp15
{
    public partial class Window3 : Window
    {
        private readonly Dictionary<string, string> _knowledgeBase;
        private int _currentUserId;
        private const string ApiToken = "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"; // Замените на свой
        

        // Таймер для периодического обновления данных об использовании сети
        private System.Timers.Timer _networkStatsTimer;

        // Начальные значения для отслеживания использованных данных
        private Dictionary<string, long> _initialNetworkStats = new Dictionary<string, long>();
        private Dictionary<string, long> _currentNetworkStats = new Dictionary<string, long>();
        private double _totalUsageGB = 0;
        private DateTime _lastSaveTime = DateTime.Now;

        // Путь к файлу для сохранения статистики
        private readonly string _dataUsageFilePath;

        private const string ApiUrl = "https://router.huggingface.co/nebius/v1/chat/completions";
        private const string ApiKey = "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";

        private readonly HttpClient _httpClient;

        public Window3(int userId = 1)
        {
            try
            {
                InitializeComponent();
                _currentUserId = userId;
                _httpClient = new HttpClient();
                _knowledgeBase = InitializeKnowledgeBase();

                // Определяем путь к файлу сохранения для конкретного пользователя
                string dataDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WpfApp15", "UserData");
                _dataUsageFilePath = Path.Combine(dataDirectory, $"datausage_{userId}.json");

                // Создаем директорию для хранения данных, если она не существует
                if (!Directory.Exists(dataDirectory))
                {
                    Directory.CreateDirectory(dataDirectory);
                }

                CreateExitButton();
                TogetherAIClient();

                // Загружаем сохраненные данные об использовании сети
                LoadSavedNetworkUsage();

                // Инициализация и запуск таймера для периодического обновления статистики
                _networkStatsTimer = new System.Timers.Timer(10000); // Обновление каждые 10 секунд
                _networkStatsTimer.Elapsed += OnNetworkStatsTimerElapsed;
                _networkStatsTimer.AutoReset = true;
                _networkStatsTimer.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при инициализации: {ex.Message}",
                              "Ошибка",
                              MessageBoxButton.OK,
                              MessageBoxImage.Error);
                Close();
            }
        }

        private void OnNetworkStatsTimerElapsed(object sender, ElapsedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                UpdateNetworkStatistics();
                if ((DateTime.Now - _lastSaveTime).TotalMinutes >= 10) // Сохранение каждые 10 минут
                {
                    SaveNetworkUsage();
                    _lastSaveTime = DateTime.Now;
                }
            });
        }

        private void LoadSavedNetworkUsage()
        {
            try
            {
                // Сначала пытаемся загрузить из базы данных
                using (var conn = new NpgsqlConnection(ConnectionString.Path))
                {
                    conn.Open();
                    string query = @"
                        SELECT total_usage_gb
                        FROM NetworkUsage
                        WHERE client_id = @clientId
                        ORDER BY usage_date DESC
                        LIMIT 1";

                    using (var cmd = new NpgsqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@clientId", _currentUserId);
                        var result = cmd.ExecuteScalar();
                        if (result != null)
                        {
                            _totalUsageGB = Convert.ToDouble(result);
                            Dispatcher.Invoke(() =>
                            {
                                DataUsageTextBlock.Text = $"Использовано данных: {_totalUsageGB:F2} GB";
                            });
                            Console.WriteLine($"Загружены данные об использовании из БД: {_totalUsageGB:F2} GB");
                            return;
                        }
                    }
                }

                // Если данных в БД нет, загружаем из файла
                if (File.Exists(_dataUsageFilePath))
                {
                    string json = File.ReadAllText(_dataUsageFilePath);
                    var savedData = JsonSerializer.Deserialize<Dictionary<string, object>>(json);

                    if (savedData != null && savedData.ContainsKey("TotalUsageGB"))
                    {
                        _totalUsageGB = Convert.ToDouble(savedData["TotalUsageGB"].ToString());
                        Dispatcher.Invoke(() =>
                        {
                            DataUsageTextBlock.Text = $"Использовано данных: {_totalUsageGB:F2} GB";
                        });
                        Console.WriteLine($"Загружены данные об использовании из файла: {_totalUsageGB:F2} GB");
                    }
                }
                else
                {
                    Console.WriteLine("Файл с данными об использовании сети не найден. Создаем новый.");
                    _totalUsageGB = 0;
                    SaveNetworkUsage();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка загрузки сохраненных данных: {ex.Message}");
                _totalUsageGB = 0;
            }
        }

        private void SaveNetworkUsage()
        {
            try
            {
                var dataToSave = new Dictionary<string, object>
                {
                    { "TotalUsageGB", _totalUsageGB },
                    { "LastUpdated", DateTime.Now },
                    { "UserId", _currentUserId }
                };

                string json = JsonSerializer.Serialize(dataToSave, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                File.WriteAllText(_dataUsageFilePath, json);
                Console.WriteLine($"Данные об использовании сети сохранены в файл: {_totalUsageGB:F2} GB");

                // Сохраняем данные в базу данных
                SaveNetworkUsageToDatabase();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка сохранения данных об использовании сети: {ex.Message}");
                MessageBox.Show($"Не удалось сохранить данные об использовании сети: {ex.Message}",
                               "Предупреждение",
                               MessageBoxButton.OK,
                               MessageBoxImage.Warning);
            }
        }

        private void SaveNetworkUsageToDatabase()
        {
            try
            {
                using (var conn = new NpgsqlConnection(ConnectionString.Path))
                {
                    conn.Open();
                    string query = @"
                        INSERT INTO NetworkUsage (client_id, usage_date, total_usage_gb, user_id)
                        VALUES (@clientId, @usageDate, @totalUsageGb, @userId)";

                    using (var cmd = new NpgsqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@clientId", _currentUserId);
                        cmd.Parameters.AddWithValue("@usageDate", DateTime.Now);
                        cmd.Parameters.AddWithValue("@totalUsageGb", (decimal)_totalUsageGB);
                        cmd.Parameters.AddWithValue("@userId", _currentUserId);

                        cmd.ExecuteNonQuery();
                    }
                }
                Console.WriteLine($"Данные об использовании сети сохранены в БД: {_totalUsageGB:F2} GB");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка сохранения данных в БД: {ex.Message}");
                MessageBox.Show($"Не удалось сохранить данные об использовании сети в базу данных: {ex.Message}",
                               "Ошибка",
                               MessageBoxButton.OK,
                               MessageBoxImage.Error);
            }
        }

        private void UpdateNetworkStatistics()
        {
            try
            {
                NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();
                Dictionary<string, long> newStats = new Dictionary<string, long>();

                foreach (NetworkInterface adapter in adapters)
                {
                    if (adapter.OperationalStatus == OperationalStatus.Up &&
                        (adapter.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 ||
                         adapter.NetworkInterfaceType == NetworkInterfaceType.Ethernet ||
                         adapter.NetworkInterfaceType == NetworkInterfaceType.GigabitEthernet))
                    {
                        IPv4InterfaceStatistics stats = adapter.GetIPv4Statistics();
                        string adapterId = adapter.Id;
                        long totalBytes = stats.BytesReceived + stats.BytesSent;
                        newStats[adapterId] = totalBytes;
                    }
                }

                if (_initialNetworkStats.Count == 0)
                {
                    foreach (var kvp in newStats)
                    {
                        _initialNetworkStats[kvp.Key] = kvp.Value;
                    }
                }

                double sessionUsageGB = 0;
                foreach (var kvp in newStats)
                {
                    if (_initialNetworkStats.ContainsKey(kvp.Key))
                    {
                        long bytesUsed = kvp.Value - _initialNetworkStats[kvp.Key];
                        if (bytesUsed > 0)
                        {
                            sessionUsageGB += bytesUsed / 1073741824.0;
                        }
                    }
                }

                _totalUsageGB += sessionUsageGB;
                _initialNetworkStats = new Dictionary<string, long>(newStats);

                double windowsDataUsageGB = GetWindowsDataUsage();
                if (windowsDataUsageGB > 0)
                {
                    _totalUsageGB = windowsDataUsageGB + sessionUsageGB;
                }

                DataUsageTextBlock.Text = $"Использовано данных: {_totalUsageGB:F2} GB";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при обновлении статистики сети: {ex.Message}");
            }
        }

        private double GetWindowsDataUsage()
        {
            try
            {
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo.FileName = "netsh";
                process.StartInfo.Arguments = "wlan show interfaces";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.CreateNoWindow = true;

                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                double totalUsageMB = 0;
                try
                {
                    using (var upCounter = new System.Diagnostics.PerformanceCounter("Network Interface", "Bytes Sent/sec", GetActiveNetworkInterfaceName(), true))
                    using (var downCounter = new System.Diagnostics.PerformanceCounter("Network Interface", "Bytes Received/sec", GetActiveNetworkInterfaceName(), true))
                    {
                    }
                }
                catch
                {
                }

                return totalUsageMB / 1024.0;
            }
            catch
            {
                return 0;
            }
        }

        private string GetActiveNetworkInterfaceName()
        {
            try
            {
                NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();
                foreach (NetworkInterface adapter in adapters)
                {
                    if (adapter.OperationalStatus == OperationalStatus.Up &&
                        (adapter.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 ||
                         adapter.NetworkInterfaceType == NetworkInterfaceType.Ethernet))
                    {
                        return adapter.Description;
                    }
                }
                return "";
            }
            catch
            {
                return "";
            }
        }

        public void TogetherAIClient()
        {
            _httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", ApiKey);
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
        }

        private async Task<string> GetYandexGptResponse(string message)
        {
            try
            {
                var requestData = new
                {
                    messages = new[]
                    {
                        new
                        {
                            role = "system",
                            content = "Ты — AI-ассистент интернет-провайдера. " +
                                     "Отвечай вежливо, технически точно, но простыми словами. " +
                                     "Не придумывай несуществующие тарифы или услуги. " +
                                     "Если вопрос неясен — уточни детали. " +
                                     "Сообщение должно быть лаконичным. " +
                                     "Исключи лишние знаки, которыми ты делаешь обводку или что-то, это не работает здесь. Только текст и отступы. Не используй систему MarkDown " +
                                     "Ответ не должен быть громоздким, чтобы не заставлять пользователя ждать. Лучше дели большое сообщение на несколько небольших." +
                                     "В случае если не удаётся решить проблему, то ты выдаёшь номер телефона поддержки +79088252404"
                        },
                        new
                        {
                            role = "user",
                            content = message
                        }
                    },
                    max_tokens = 250,
                    model = "deepseek-ai/DeepSeek-V3-0324-fast",
                };

                var json = JsonSerializer.Serialize(requestData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync(ApiUrl, content);
                response.EnsureSuccessStatusCode();

                var responseJson = await response.Content.ReadAsStringAsync();
                var responseObject = JsonSerializer.Deserialize<JsonElement>(responseJson);

                var assistantMessage = responseObject
                    .GetProperty("choices")[0]
                    .GetProperty("message")
                    .GetProperty("content")
                    .GetString();

                return assistantMessage ?? "Пустой ответ";
            }
            catch (Exception ex)
            {
                return $"Ошибка подключения к HuggingFace: {ex.Message}";
            }
        }

        public class Service
        {
            public int ServiceId { get; set; }
            public string ServiceName { get; set; }
            public decimal ServicePrice { get; set; }
            public string ServiceDescription { get; set; }
        }

        private void CreateExitButton()
        {
            Button exitButton = new Button
            {
                Content = "Выйти",
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, 12, 30, 0), // Increased right margin to shift left
                Width = 120, // Consistent with other buttons
                FontSize = 16 // Match font size with other buttons
            };
            exitButton.Click += ExitButton_Click;
            Grid.SetColumn(exitButton, 1);
            Grid.SetRow(exitButton, 0);
            MainGrid.Children.Add(exitButton);
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            SaveNetworkUsage();
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                LoadDefaultUserData();
                LoadDefaultStatistics();
                UpdateNetworkStatistics();
                AddMessageToChat("Поддержка", "Здравствуйте! Чем могу помочь вам сегодня?");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}",
                                "Ошибка",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
            }
        }

        private void LoadDefaultUserData()
        {
            try
            {
                NameTextBlock.Text = "Иванов Иван";
                ContractNumberTextBlock.Text = GetContractNumber();
                LoadCurrentTariff();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}",
                                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                ContractNumberTextBlock.Text = "Не найден";
            }
        }

        private string GetContractNumber()
        {
            try
            {
                using (var conn = new NpgsqlConnection(ConnectionString.Path))
                {
                    conn.Open();
                    var query = "SELECT contract_id FROM ClientContracts WHERE client_id = @clientId";
                    using (var cmd = new NpgsqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@clientId", _currentUserId);
                        var result = cmd.ExecuteScalar();
                        return result?.ToString() ?? "Не найден";
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка загрузки номера договора: {ex.Message}");
                return "Не найден";
            }
        }

        private void LoadDefaultStatistics()
        {
            LastLoginTextBlock.Text = $"Последний вход: {DateTime.Now.AddDays(-0):dd.MM.yyyy HH:mm}";
            BillStatusTextBlock.Text = GetPaymentStatus();
        }

        private string GetPaymentStatus()
        {
            try
            {
                using (var conn = new NpgsqlConnection(ConnectionString.Path))
                {
                    conn.Open();
                    var query = "SELECT balance FROM Invoices WHERE client_id = @clientId ORDER BY invoice_id DESC LIMIT 1";
                    using (var cmd = new NpgsqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@clientId", _currentUserId);
                        var result = cmd.ExecuteScalar();
                        if (result != null)
                        {
                            decimal balance = Convert.ToDecimal(result);
                            return balance <= 0 ? "Оплачено" : $"Ожидает оплаты (баланс: {balance:F2} руб.)";
                        }
                        return "Нет данных";
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка загрузки статуса оплаты: {ex.Message}");
                return "Ошибка";
            }
        }

        private List<Service> GetAvailableTariffs()
        {
            var tariffs = new List<Service>();
            try
            {
                using (var conn = new NpgsqlConnection(ConnectionString.Path))
                {
                    conn.Open();
                    var query = "SELECT service_id, service_name, service_price, service_description FROM Services";
                    using (var cmd = new NpgsqlCommand(query, conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            tariffs.Add(new Service
                            {
                                ServiceId = reader.GetInt32(0),
                                ServiceName = reader.GetString(1),
                                ServicePrice = reader.GetDecimal(2),
                                ServiceDescription = reader.GetString(3)
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки тарифов: {ex.Message}",
                                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return tariffs;
        }

        private void UpdateClientTariff(int newServiceId)
        {
            try
            {
                using (var conn = new NpgsqlConnection(ConnectionString.Path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand(
                        "SELECT COUNT(*) FROM ClientServices WHERE client_id = @clientId", conn))
                    {
                        cmd.Parameters.AddWithValue("@clientId", _currentUserId);
                        int count = Convert.ToInt32(cmd.ExecuteScalar());

                        if (count > 0)
                        {
                            using (var updateCmd = new NpgsqlCommand(
                                "UPDATE ClientServices SET service_id = @serviceId WHERE client_id = @clientId", conn))
                            {
                                updateCmd.Parameters.AddWithValue("@clientId", _currentUserId);
                                updateCmd.Parameters.AddWithValue("@serviceId", newServiceId);
                                updateCmd.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            using (var insertCmd = new NpgsqlCommand(
                                "INSERT INTO ClientServices (client_id, service_id) VALUES (@clientId, @serviceId)", conn))
                            {
                                insertCmd.Parameters.AddWithValue("@clientId", _currentUserId);
                                insertCmd.Parameters.AddWithValue("@serviceId", newServiceId);
                                insertCmd.ExecuteNonQuery();
                            }
                        }
                    }
                }

                LoadCurrentTariff();
                MessageBox.Show("Тариф успешно изменен!", "Успех",
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка изменения тарифа: {ex.Message}",
                                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadCurrentTariff()
        {
            try
            {
                using (var conn = new NpgsqlConnection(ConnectionString.Path))
                {
                    conn.Open();
                    var query = @"
                    SELECT s.service_name 
                    FROM Services s
                    JOIN ClientServices cs ON s.service_id = cs.service_id
                    WHERE cs.client_id = @clientId";
                    using (var cmd = new NpgsqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@clientId", _currentUserId);
                        var result = cmd.ExecuteScalar();
                        TariffTextBlock.Text = result?.ToString() ?? "Тариф не выбран";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки тарифа: {ex.Message}",
                              "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UploadAvatar_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Изображения (*.jpg, *.jpeg, *.png)|*.jpg;*.jpeg;*.png"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    string fileName = openFileDialog.FileName;
                    if (System.IO.File.Exists(fileName))
                    {
                        AvatarImage.Source = new BitmapImage(new Uri(fileName));
                        AvatarImage.Visibility = Visibility.Visible;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка загрузки аватара: {ex.Message}",
                                    "Ошибка",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Error);
                }
            }
        }

        private void AddMessageToChat(string sender, string message)
        {
            ChatHistory.Children.Add(new TextBlock
            {
                Text = $"{sender}: {message}",
                Margin = new Thickness(5),
                TextWrapping = TextWrapping.Wrap
            });
        }

        private Dictionary<string, string> InitializeKnowledgeBase()
        {
            return new Dictionary<string, string>
            {
                { "тариф", "Ваш текущий тариф: '{0}'. Для смены тарифа вы можете:\n1. Выбрать новый тариф в личном кабинете\n2. Позвонить в поддержку\n3. Посетить наш офис" },
                { "оплата", "Способы оплаты услуг:\n1. Банковской картой через личный кабинет\n2. В офисах компании\n3. Через терминалы самообслуживания\n4. Банковским переводом" },
                { "интернет", "Диагностика проблем с интернетом:\n1. Перезагрузите роутер\n2. Проверьте кабельное подключение\n3. Запустите диагностику Windows\n4. Измерьте скорость на speedtest.net" },
                { "техподдержка", "Наша техподдержка доступна 24/7:\n1. По телефону: 8-800-XXX-XXXX\n2. Email: support@provider.ru\n3. Через личный кабинет\n4. В офисах обслуживания" }
            };
        }

        private async void SendMessage_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string userMessage = ChatInputBox.Text;
                if (string.IsNullOrWhiteSpace(userMessage))
                {
                    MessageBox.Show("Введите сообщение.",
                                  "Предупреждение",
                                  MessageBoxButton.OK,
                                  MessageBoxImage.Warning);
                    return;
                }

                AddMessageToChat("Вы", userMessage);
                ChatInputBox.Clear();

                var botResponse = await GetYandexGptResponse(userMessage);
                AddMessageToChat("Поддержка", botResponse);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка отправки сообщения: {ex.Message}",
                              "Ошибка",
                              MessageBoxButton.OK,
                              MessageBoxImage.Error);
            }
        }

        private void ConnectTariff_Click(object sender, RoutedEventArgs e)
        {
            var tariffs = GetAvailableTariffs();
            if (tariffs.Count == 0)
            {
                MessageBox.Show("Нет доступных тарифов",
                              "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dialog = new TariffSelectionWindow(tariffs);
            if (dialog.ShowDialog() == true)
            {
                UpdateClientTariff(dialog.SelectedService.ServiceId);
            }
        }

        private void ShowPaymentHistory_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var paymentHistoryWindow = new PaymentHistoryWindow(_currentUserId);
                paymentHistoryWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка открытия истории платежей: {ex.Message}",
                              "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            try
            {
                SaveNetworkUsage();
                if (_networkStatsTimer != null)
                {
                    _networkStatsTimer.Stop();
                    _networkStatsTimer.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при закрытии окна: {ex.Message}");
            }
            finally
            {
                base.OnClosed(e);
            }
        }
    }
}