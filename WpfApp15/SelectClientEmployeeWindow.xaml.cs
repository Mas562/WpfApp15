using System;
using System.Collections.Generic;
using System.Windows;
using Npgsql;

namespace WpfApp15
{
    public partial class SelectClientEmployeeWindow : Window
    {
        private string connectionString = ConnectionString.Path;
        private List<Client> clients;
        private List<Employee> employees;

        public class Client
        {
            public int ClientId { get; set; }
            public string DisplayName { get; set; }
        }

        public class Employee
        {
            public int EmployeeId { get; set; }
            public string DisplayName { get; set; }
        }

        public SelectClientEmployeeWindow()
        {
            InitializeComponent();
            LoadClientsWithoutContracts();
            LoadEmployees();
        }

        private void LoadClientsWithoutContracts()
        {
            clients = new List<Client>();
            try
            {
                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    string query = @"
                        SELECT 
                            c.client_id,
                            CONCAT(c.last_name, ' ', c.first_name, ' ', COALESCE(c.middle_name, '')) AS full_name
                        FROM Clients c
                        LEFT JOIN ClientContracts cc ON c.client_id = cc.client_id
                        WHERE cc.contract_id IS NULL
                        ORDER BY c.last_name, c.first_name";
                    using (var cmd = new NpgsqlCommand(query, conn))
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                clients.Add(new Client
                                {
                                    ClientId = reader.GetInt32(0),
                                    DisplayName = reader.GetString(1).Trim()
                                });
                            }
                        }
                    }
                }
                ClientComboBox.ItemsSource = clients;
                if (clients.Count > 0)
                    ClientComboBox.SelectedIndex = 0;
                else
                    MessageBox.Show("Нет клиентов без договоров. Добавьте нового клиента или удалите существующий договор.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке клиентов: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadEmployees()
        {
            employees = new List<Employee>();
            try
            {
                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    string query = @"
                        SELECT 
                            employee_id,
                            CONCAT(last_name, ' ', first_name, ' ', COALESCE(middle_name, '')) AS full_name
                        FROM Employees
                        ORDER BY last_name, first_name";
                    using (var cmd = new NpgsqlCommand(query, conn))
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                employees.Add(new Employee
                                {
                                    EmployeeId = reader.GetInt32(0),
                                    DisplayName = reader.GetString(1).Trim()
                                });
                            }
                        }
                    }
                }
                EmployeeComboBox.ItemsSource = employees;
                if (employees.Count > 0)
                    EmployeeComboBox.SelectedIndex = 0;
                else
                    MessageBox.Show("Список сотрудников пуст. Добавьте хотя бы одного сотрудника в таблицу Employees.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке сотрудников: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CreateContract_Click(object sender, RoutedEventArgs e)
        {
            if (ClientComboBox.SelectedItem == null || EmployeeComboBox.SelectedItem == null)
            {
                MessageBox.Show("Выберите клиента и сотрудника.", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var selectedClient = (Client)ClientComboBox.SelectedItem;
            var selectedEmployee = (Employee)EmployeeComboBox.SelectedItem;

            try
            {
                int contractId;
                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();

                    // Генерация уникального contract_id
                    contractId = GenerateUniqueContractId(conn);

                    // Вставка в ClientContracts
                    string insertClientContract = @"
                        INSERT INTO ClientContracts
                            (client_id, contract_date, contract_id)
                        VALUES
                            (@clientId, CURRENT_DATE, @contractId)
                        RETURNING contract_id";
                    using (var cmd = new NpgsqlCommand(insertClientContract, conn))
                    {
                        cmd.Parameters.AddWithValue("clientId", selectedClient.ClientId);
                        cmd.Parameters.AddWithValue("contractId", contractId);
                        contractId = Convert.ToInt32(cmd.ExecuteScalar());
                    }

                    // Вставка в EmployeeContracts
                    string insertEmployeeContract = @"
                        INSERT INTO EmployeeContracts
                            (employee_id, contract_id, contract_date)
                        VALUES
                            (@employeeId, @contractId, CURRENT_DATE)";
                    using (var cmd = new NpgsqlCommand(insertEmployeeContract, conn))
                    {
                        cmd.Parameters.AddWithValue("employeeId", selectedEmployee.EmployeeId);
                        cmd.Parameters.AddWithValue("contractId", contractId);
                        cmd.ExecuteNonQuery();
                    }
                }

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании договора: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private int GenerateUniqueContractId(NpgsqlConnection conn)
        {
            var rnd = new Random();
            int contractId;
            bool isUnique;
            do
            {
                contractId = rnd.Next(100000, 999999);
                string checkQuery = "SELECT COUNT(*) FROM ClientContracts WHERE contract_id = @contractId";
                using (var cmd = new NpgsqlCommand(checkQuery, conn))
                {
                    cmd.Parameters.AddWithValue("contractId", contractId);
                    isUnique = Convert.ToInt32(cmd.ExecuteScalar()) == 0;
                }
            } while (!isUnique);
            return contractId;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}