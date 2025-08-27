using System.Windows;
using Npgsql;
using System.Data;
using System.Collections.Generic;
using System;

namespace WpfApp15
{
    public partial class PaymentHistoryWindow : Window
    {
        private int _clientId;
        public PaymentHistoryWindow(int clientId)
        {
            InitializeComponent();
            _clientId = clientId;
            LoadPaymentHistory();
        }
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
        private void LoadPaymentHistory()
        {
            var payments = new List<Payment>();

            try
            {
                using (var conn = new NpgsqlConnection(ConnectionString.Path))
                {
                    conn.Open();
                    var query = @"
                        SELECT 
                            p.amount, 
                            c.date AS payment_date,
                            'Оплата' AS type
                        FROM Payments p
                        JOIN Charges c ON p.payment_id = c.charge_id
                        WHERE p.client_id = @clientId
                        UNION ALL
                        SELECT 
                            ch.amount, 
                            ch.date AS payment_date,
                            'Начисление' AS type
                        FROM Charges ch
                        WHERE ch.contract_id IN (
                            SELECT contract_id 
                            FROM ClientContracts 
                            WHERE client_id = @clientId
                        )
                        ORDER BY payment_date DESC";

                    using (var cmd = new NpgsqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@clientId", _clientId);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                payments.Add(new Payment
                                {
                                    Date = reader.GetDateTime(1).ToString("dd.MM.yyyy"),
                                    Amount = reader.GetDecimal(0).ToString("C"),
                                    Type = reader.GetString(2)
                                });
                            }
                        }
                    }
                }

                PaymentsDataGrid.ItemsSource = payments;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки истории платежей: {ex.Message}",
                              "Ошибка",
                              MessageBoxButton.OK,
                              MessageBoxImage.Error);
            }
        }
    }


    public class Payment
    {
        public string Date { get; set; }
        public string Amount { get; set; }
        public string Type { get; set; }
    }
}