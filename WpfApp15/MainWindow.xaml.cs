using System;
using System.Windows;
using System.Windows.Media;
using Npgsql;

namespace WpfApp15
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void LoginButton_Click(object sender, RoutedEventArgs e)
        {
       

            string username = UsernameTextBox.Text;
            string password = PasswordBox.Password;
            var (role, userId) = ValidateCredentials(username, password);

            if (role != null)
            {
                MessageBox.Show("Вход выполнен успешно!");
                switch (role.ToLowerInvariant())
                {
                    case "manager":
                        Window2 window2 = new Window2();
                        window2.Show();
                        break;
                    case "support":
                        Window1 window1 = new Window1();
                        window1.Show();
                        break;
                    // Открываем профиль клиента для распространенных названий роли
                    case "client":
                    case "user":
                    case "client_user":
                    case "клиент":
                    case "mas562":
                        Window3 window3 = new Window3(userId);  // Передаем userId
                        window3.Show();
                        break;
                    default:
                        MessageBox.Show("Неизвестная роль. Обратитесь к администратору.");
                        break;
                }
                this.Close();
            }
            else
            {
                MessageBox.Show("Неверный логин или пароль.");
            }
        }
    

        private (string role, int userId) ValidateCredentials(string username, string password)
        {
            string connString = ConnectionString.Path;
            using (var conn = new NpgsqlConnection(connString))
            {
                try
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand(
                        "SELECT role.name_role, \"user\".id FROM \"user\" " +
                        "JOIN role ON \"user\".id = role.user_id " +
                        "WHERE login = @login AND password = @password", conn))
                    {
                        cmd.Parameters.AddWithValue("login", username);
                        cmd.Parameters.AddWithValue("password", password);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                string role = reader.GetString(0);
                                int userId = reader.GetInt32(1);
                                return (role, userId);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка подключения к базе данных: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            return (null, 0);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Получаем текущий словарь ресурсов
            var currentDict = Application.Current.Resources.MergedDictionaries[0];

            // Определяем, какая тема сейчас активна
            bool isDarkTheme = currentDict.Source != null &&
                              currentDict.Source.ToString().Contains("DarkTheme.xaml");

            // Загружаем противоположную тему
            var newTheme = new ResourceDictionary();
            newTheme.Source = new Uri(isDarkTheme
                ? "Themes/LightTheme.xaml"
                : "Themes/DarkTheme.xaml", UriKind.Relative);

            // Заменяем текущую тему
            Application.Current.Resources.MergedDictionaries[0] = newTheme;

            // Обновляем иконку кнопки
            ThemeToggleButton.Content = isDarkTheme ? "🌙" : "☀";
        }
    }
}