using System;
using System.Data;
using System.Windows.Forms;
using System.Xml;
using System.Data.SqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Collections.Generic;

namespace TestSolid
{
    public partial class Form1 : Form
    {
        private string connectionString = "Server=127.0.0.1;Database=Testt;User Id=sa;Password=Admin12345;";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime date = dateTimePicker1.Value;

            // Получаем котировки валют на эту дату
            DataTable table = GetCurrencyRates(date);

            // Отображаем котировки в DataGridView
            dataGridView1.DataSource = table;

            // Сохраняем котировки в базу данных 
            SaveCurrencyRates(table, date);

            // Создаем Excel-файл из базы данных // Добавить вызов метода для создания Excel-файла
            CreateExcelFile(date);
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime date = dateTimePicker1.Value;

            DataTable table = GetCurrencyRates(date);

            dataGridView1.DataSource = table;
        }

        // Метод для получения котировок валют на указанную дату
        private DataTable GetCurrencyRates(DateTime date)
        {
            DataTable table = new DataTable();

            table.Columns.Add("ID", typeof(string));
            table.Columns.Add("NumCode", typeof(string));
            table.Columns.Add("CharCode", typeof(string));
            table.Columns.Add("Nominal", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Rate", typeof(decimal));

            string url = "http://www.cbr.ru/scripts/XML_daily.asp?date_req=" + date.ToString("dd/MM/yyyy");

            XmlDocument xml = new XmlDocument();

            xml.Load(url);

            XmlNodeList nodes = xml.SelectNodes("//Valute");

            foreach (XmlNode node in nodes)
            {
                string id = node.Attributes["ID"].Value;
                int numCode = int.Parse(node["NumCode"].InnerText);
                string charCode = node["CharCode"].InnerText;
                int nominal = int.Parse(node["Nominal"].InnerText);
                string name = node["Name"].InnerText;
                decimal rate = decimal.Parse(node["Value"].InnerText);

                table.Rows.Add(id, numCode, charCode, nominal, name, rate);
            }

            return table;
        }

        // Метод для сохранения котировок валют в базу данных 
        private void SaveCurrencyRates(DataTable table, DateTime date)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    foreach (DataRow row in table.Rows)
                    {

                        string id = row["ID"].ToString();
                        string numCode = row["NumCode"].ToString();
                        string charCode = row["CharCode"].ToString();
                        int nominal = (int)row["Nominal"];
                        string name = row["Name"].ToString();
                        decimal rate = (decimal)row["Rate"];

                        string checkCurrencyQuery = "SELECT COUNT(*) FROM Currency WHERE CharCode = @charCode";
                        SqlCommand checkCurrencyCommand = new SqlCommand(checkCurrencyQuery, connection);
                        checkCurrencyCommand.Parameters.AddWithValue("@charCode", charCode);
                        int count = (int)checkCurrencyCommand.ExecuteScalar();

                        if (count == 0)
                        {
                            string insertCurrencyQuery = "INSERT INTO Currency (ID, NumCode, CharCode) VALUES (@id, @numCode, @charCode)";
                            SqlCommand insertCurrencyCommand = new SqlCommand(insertCurrencyQuery, connection);
                            insertCurrencyCommand.Parameters.AddWithValue("@id", id);
                            insertCurrencyCommand.Parameters.AddWithValue("@numCode", numCode);
                            insertCurrencyCommand.Parameters.AddWithValue("@charCode", charCode);
                            insertCurrencyCommand.ExecuteNonQuery();
                        }

                        string checkRateQuery = "SELECT COUNT(*) FROM Rate R INNER JOIN Currency C ON R.CurrencyID = C.CurrencyID WHERE C.CharCode = @charCode AND R.Date = CONVERT(date, @date, 104)";
                        SqlCommand checkRateCommand = new SqlCommand(checkRateQuery, connection);
                        checkRateCommand.Parameters.AddWithValue("@charCode", charCode);
                        checkRateCommand.Parameters.AddWithValue("@date", date);
                        count = (int)checkRateCommand.ExecuteScalar();

                        SqlCommand getCurrencyIdCommand = new SqlCommand("SELECT CurrencyID FROM Currency WHERE CharCode = @charCode", connection);
                        getCurrencyIdCommand.Parameters.AddWithValue("@charCode", charCode);
                        int currencyId = (int)getCurrencyIdCommand.ExecuteScalar();

                        if (count > 0)
                        {
                            string updateRateQuery = "UPDATE Rate SET Nominal = @nominal, Value = @value WHERE CurrencyID = @currencyid AND Date = @date";
                            SqlCommand updateRateCommand = new SqlCommand(updateRateQuery, connection);
                            updateRateCommand.Parameters.AddWithValue("@nominal", nominal);
                            updateRateCommand.Parameters.AddWithValue("@value", rate);
                            updateRateCommand.Parameters.AddWithValue("@currencyid", currencyId);
                            updateRateCommand.Parameters.AddWithValue("@date", date);
                            updateRateCommand.ExecuteNonQuery();
                        }
                        else
                        {
                            string insertRateQuery = "INSERT INTO Rate (CurrencyID, Date, Nominal, Value) VALUES (@currencyid, @date, @nominal, @value)";
                            SqlCommand insertRateCommand = new SqlCommand(insertRateQuery, connection);
                            insertRateCommand.Parameters.AddWithValue("@currencyid", currencyId);
                            insertRateCommand.Parameters.AddWithValue("@date", date);
                            insertRateCommand.Parameters.AddWithValue("@nominal", nominal);
                            insertRateCommand.Parameters.AddWithValue("@value", rate);
                            insertRateCommand.ExecuteNonQuery();
                        }
                    }

                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка при работе с базой данных");
            }
        }

        // Метод для создания Excel-файла из базы данных
        private void CreateExcelFile(DateTime date)
        {
            DataTable dataTable = new DataTable();

            string query = @"SELECT c.CharCode, r.Nominal, r.Value FROM Currency c JOIN Rate r ON c.CurrencyID = r.CurrencyID WHERE r.Date = CONVERT(date, @Date, 104)";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Date", date);
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }
            }

            if (dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для указанной даты.");
                return;
            }

            string fileName = date.ToString("yyyyMMdd") + ".xlsx";
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(fileName)))
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    string currencyCode = row["CharCode"].ToString();
                    int nominal = Convert.ToInt32(row["Nominal"]);
                    decimal value = Convert.ToDecimal(row["Value"]);

                    ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add(currencyCode);

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        excelWorksheet.Cells[1, i + 2].Value = dataTable.Rows[i]["CharCode"].ToString();
                        excelWorksheet.Cells[1, i + 2].Style.Font.Bold = true;
                        excelWorksheet.Cells[1, i + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        excelWorksheet.Cells[1, i + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        excelWorksheet.Cells[i + 2, 1].Value = dataTable.Rows[i]["CharCode"].ToString();
                        excelWorksheet.Cells[i + 2, 1].Style.Font.Bold = true;
                        excelWorksheet.Cells[i + 2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        excelWorksheet.Cells[i + 2, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Rows.Count; j++)
                        {
                            int otherNominal = Convert.ToInt32(dataTable.Rows[j]["Nominal"]);
                            decimal otherValue = Convert.ToDecimal(dataTable.Rows[j]["Value"]);

                            decimal crossRate = Math.Round((nominal / value) / (otherNominal / otherValue), 4);

                            excelWorksheet.Cells[i + 2, j + 2].Value = crossRate;
                            excelWorksheet.Cells[i + 2, j + 2].Style.Numberformat.Format = "0.0000";
                            excelWorksheet.Cells[i + 2, j + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            excelWorksheet.Cells[i + 2, j + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        }
                    }

                    excelWorksheet.Cells.AutoFitColumns();
                }

                excelPackage.Save();
            }

            MessageBox.Show("Файл Excel успешно создан: " + fileName);
        }


    }
}
            