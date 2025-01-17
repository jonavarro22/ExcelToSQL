using System;
using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using CsvHelper;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelToSQL
{
    public partial class MainWindow : Window
    {
        private const string ConfigFilePath = "C:\\ProgramData\\FloatyRock\\ExcelToSQL\\settings.json";
        private string selectedSQLType = "MSSQL"; // Default to MSSQL
        private string CurrentFilePath; // To store the path of the loaded file


        public MainWindow()
        {
            InitializeComponent();
            LoadSettings(); // Load saved settings on startup
            SetLocalizedText(); // Set default localized text
        }

        /// <summary>
        /// Handles changes to the Operation ComboBox (Create Table or Update Table).
        /// </summary>
        private void OperationComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SaveSettings(); // Save the current selection whenever it changes
        }

        /// <summary>
        /// Handles changes to the Language Selector ComboBox.
        /// </summary>
        private void LanguageSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems[0] is ComboBoxItem selectedItem && selectedItem.Tag is not null)
            {
                string selectedLanguage = selectedItem.Tag.ToString();
                if (!string.IsNullOrEmpty(selectedLanguage))
                {
                    LocalizationManager.SetLanguage(selectedLanguage);
                    SetLocalizedText(); // Update UI with localized text
                    SaveSettings(); // Save the selected language
                }
            }
        }

        /// <summary>
        /// Handles changes to the Target SQL ComboBox (MSSQL, MySQL, PostgreSQL).
        /// </summary>
        private void TargetSQLComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0 && e.AddedItems[0] is ComboBoxItem selectedItem)
            {
                if (selectedItem.Tag != null)
                {
                    selectedSQLType = selectedItem.Tag.ToString();
                }
                else
                {
                    // Handle the case where Tag is null
                    selectedSQLType = string.Empty; // or some default value
                }
            }
        }

        /// <summary>
        /// Handles the drag-and-drop event for files.
        /// </summary>
        private void Border_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
        }
        /// <summary>
        /// Handles the drag-and-drop event for files.
        /// </summary>
        private void Border_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                string filePath = files[0]; // Handle the first dropped file
                LoadFile(filePath);
            }
        }

        /// <summary>
        /// Opens a file dialog for manual file upload.
        /// </summary>
        private void UploadFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "All Supported Files (*.xlsx, *.csv)|*.xlsx;*.csv|Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                LoadFile(filePath);
            }
        }

        /// <summary>
        /// Loads the selected file and displays its content in the DataGrid.
        /// </summary>
        private void LoadFile(string filePath)
        {
            try
            {
                CurrentFilePath = filePath; // Update the file path

                if (filePath.EndsWith(".csv"))
                {
                    string delimiter = DetectOrGetSelectedDelimiter(filePath);

                    if (string.IsNullOrEmpty(delimiter))
                    {
                        MessageBox.Show("Could not autodetect the delimiter. Please select one manually.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    using (var reader = new StreamReader(filePath))
                    using (var csv = new CsvReader(reader, new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture)
                    {
                        Delimiter = delimiter
                    }))
                    {
                        var records = csv.GetRecords<dynamic>().ToList();
                        var dataTable = ConvertToDataTable(records);
                        var inferredTable = InferColumnTypes(dataTable); // Infer column types
                        DataPreviewGrid.ItemsSource = inferredTable.DefaultView;
                    }
                }
                else if (filePath.EndsWith(".xlsx"))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        var dataTable = new DataTable();

                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            string header = worksheet.Cells[1, col].Text;
                            if (!string.IsNullOrWhiteSpace(header))
                            {
                                dataTable.Columns.Add(header);
                            }
                        }

                        for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                        {
                            var newRow = dataTable.NewRow();
                            bool hasData = false;

                            for (int col = 1; col <= dataTable.Columns.Count; col++)
                            {
                                var cellValue = worksheet.Cells[row, col].Text;
                                if (!string.IsNullOrWhiteSpace(cellValue))
                                {
                                    newRow[col - 1] = cellValue;
                                    hasData = true;
                                }
                            }

                            if (hasData)
                            {
                                dataTable.Rows.Add(newRow);
                            }
                        }

                        RemoveEmptyColumns(dataTable);
                        var inferredTable = InferColumnTypes(dataTable); // Infer column types
                        DataPreviewGrid.ItemsSource = inferredTable.DefaultView;
                    }
                }
                else
                {
                    MessageBox.Show("Unsupported file format. Please upload a CSV or Excel file.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Handles the infer column types button click.
        /// </summary>
        private DataTable InferColumnTypes(DataTable originalTable)
        {
            var newTable = new DataTable();

            // Infer column types and add columns to the new table
            foreach (DataColumn column in originalTable.Columns)
            {
                var nonNullValues = originalTable.AsEnumerable()
                    .Where(row => !row.IsNull(column))
                    .Select(row => row[column]);

                Type columnType;

                if (nonNullValues.All(value => int.TryParse(value.ToString(), out _)))
                {
                    columnType = typeof(int);
                }
                else if (nonNullValues.All(value => double.TryParse(value.ToString(), out _)))
                {
                    columnType = typeof(double);
                }
                else if (nonNullValues.All(value => DateTime.TryParse(value.ToString(), out _)))
                {
                    columnType = typeof(DateTime);
                }
                else
                {
                    columnType = typeof(string);
                }

                newTable.Columns.Add(column.ColumnName, columnType);
            }

            // Copy data to the new table
            foreach (DataRow row in originalTable.Rows)
            {
                var newRow = newTable.NewRow();
                foreach (DataColumn column in originalTable.Columns)
                {
                    newRow[column.ColumnName] = row[column];
                }
                newTable.Rows.Add(newRow);
            }

            return newTable;
        }

        /// <summary>
        /// Helper method to convert a list of dynamic objects to a DataTable.
        /// </summary>
        private DataTable ConvertToDataTable(IEnumerable<dynamic> records)
        {
            var dataTable = new DataTable();

            if (records.Any())
            {
                // Add columns
                foreach (var key in ((IDictionary<string, object>)records.First()).Keys)
                {
                    dataTable.Columns.Add(key);
                }

                // Add rows
                foreach (var record in records)
                {
                    var row = dataTable.NewRow();
                    foreach (var kvp in (IDictionary<string, object>)record)
                    {
                        if (dataTable.Columns.Contains(kvp.Key))
                        {
                            row[kvp.Key] = kvp.Value ?? DBNull.Value;
                        }
                    }
                    dataTable.Rows.Add(row);
                }
            }

            RemoveEmptyColumns(dataTable);
            InferColumnTypes(dataTable);

            return dataTable;
        }

        /// <summary>
        /// Removes empty columns from the DataTable.
        /// </summary>
        private void RemoveEmptyColumns(DataTable table)
        {
            for (int i = table.Columns.Count - 1; i >= 0; i--)
            {
                bool isEmpty = table.AsEnumerable().All(row => row.IsNull(i) || string.IsNullOrWhiteSpace(row[i].ToString()));
                if (isEmpty)
                {
                    table.Columns.RemoveAt(i);
                }
            }
        }

        /// <summary>
        /// Detects the delimiter used in a CSV file.
        /// </summary>
        private string DetectDelimiter(string filePath)
        {
            string[] possibleDelimiters = { ",", ";", "|", "\t" };
            string firstLine = File.ReadLines(filePath).FirstOrDefault();

            if (string.IsNullOrWhiteSpace(firstLine))
                return null;

            // Check for the most common delimiter
            return possibleDelimiters.OrderByDescending(delimiter => firstLine.Count(c => c.ToString() == delimiter))
                                      .FirstOrDefault(delimiter => firstLine.Contains(delimiter));
        }

        /// <summary>
        /// Handles the delimiter detection or selection logic.
        private string DetectOrGetSelectedDelimiter(string filePath)
        {
            var selectedItem = DelimiterComboBox.SelectedItem as ComboBoxItem;

            // Use selected delimiter if it's not "Auto"
            if (selectedItem != null && selectedItem.Tag != null && !string.IsNullOrEmpty(selectedItem.Tag.ToString()))
            {
                return selectedItem.Tag.ToString();
            }

            // Autodetect the delimiter
            return DetectDelimiter(filePath);
        }

        /// <summary>
        /// Handles the delimiter selection change event.
        /// </summary>
        private void DelimiterComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
            {
                var selectedItem = e.AddedItems[0] as ComboBoxItem;

                if (selectedItem != null && selectedItem.Tag != null)
                {
                    string selectedDelimiter = selectedItem.Tag.ToString();

                    // Notify the user if needed
                    if (selectedDelimiter == "")
                    {
                        MessageBox.Show("Delimiter set to Auto. The application will attempt to detect it.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
        }

        /// <summary>
        /// Generates SQL based on the current settings and loaded data.
        /// </summary>
        private string GenerateSQL(DataTable table, bool isCreateTable, string tableName)
        {
            var sqlBuilder = new System.Text.StringBuilder();

            if (isCreateTable)
            {
                // CREATE TABLE Query
                sqlBuilder.AppendLine($"CREATE TABLE {tableName} (");

                foreach (DataColumn column in table.Columns)
                {
                    string sqlType = GetSQLType(column.DataType);
                    sqlBuilder.AppendLine($"    {column.ColumnName} {sqlType},");
                }

                sqlBuilder.Length -= 3; // Remove the last comma
                sqlBuilder.AppendLine(");");

                // Add a line break for readability
                sqlBuilder.AppendLine();
            }

            // INSERT INTO Query (included for both Create and Update modes)
            sqlBuilder.AppendLine($"INSERT INTO {tableName} (");

            foreach (DataColumn column in table.Columns)
            {
                sqlBuilder.AppendLine($"    {column.ColumnName},");
            }

            sqlBuilder.Length -= 3; // Remove the last comma
            sqlBuilder.AppendLine(") VALUES");

            foreach (DataRow row in table.Rows)
            {
                sqlBuilder.Append("    (");

                foreach (DataColumn column in table.Columns)
                {
                    if (column.DataType == typeof(string) || column.DataType == typeof(DateTime))
                    {
                        sqlBuilder.Append($"'{row[column]?.ToString().Replace("'", "''")}',");
                    }
                    else if (column.DataType == typeof(Guid) && string.IsNullOrWhiteSpace(row[column]?.ToString()))
                    {
                        // Generate GUID based on the selected SQL type
                        if (selectedSQLType == "MSSQL")
                            sqlBuilder.Append($"NEWID(),");
                        else if (selectedSQLType == "MySQL")
                            sqlBuilder.Append($"UUID(),");
                        else if (selectedSQLType == "PostgreSQL")
                            sqlBuilder.Append($"gen_random_uuid(),");
                    }
                    else
                    {
                        sqlBuilder.Append($"{row[column]?.ToString() ?? "NULL"},");
                    }
                }

                sqlBuilder.Length -= 1; // Remove the last comma
                sqlBuilder.AppendLine("),");
            }

            sqlBuilder.Length -= 3; // Remove the last comma
            sqlBuilder.AppendLine(";");

            return sqlBuilder.ToString();
        }

        private string GetSQLType(Type type)
        {
            if (type == typeof(int))
                return "INT";
            if (type == typeof(double))
                return "FLOAT";
            if (type == typeof(DateTime))
                return selectedSQLType == "MySQL" ? "DATETIME" : "TIMESTAMP";
            if (type == typeof(Guid))
            {
                if (selectedSQLType == "MSSQL")
                    return "UNIQUEIDENTIFIER";
                if (selectedSQLType == "MySQL")
                    return "CHAR(36)";
                if (selectedSQLType == "PostgreSQL")
                    return "UUID";
            }
            return selectedSQLType == "MySQL" ? "TEXT" : "NVARCHAR(MAX)";
        }

        /// <summary>
        /// Generates SQL based on the current settings and loaded data.
        /// </summary>
        private void GenerateSQL_Click(object sender, RoutedEventArgs e)
        {
            if (DataPreviewGrid.ItemsSource is null)
            {
                MessageBox.Show("No data loaded. Please upload or drag a file.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var table = ((DataView)DataPreviewGrid.ItemsSource).ToTable();
            bool isCreateTable = OperationComboBox.SelectedIndex == 0; // 0 = Create Table, 1 = Update Table
            string fileName = Path.GetFileNameWithoutExtension(CurrentFilePath); // Assume CurrentFilePath stores the uploaded file's path

            try
            {
                var sql = GenerateSQL(table, isCreateTable, fileName);
                SaveSQLToFile(sql, fileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating SQL: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Saves the generated SQL to a file. Prompts the user for a file name and location.
        /// </summary>
        private void SaveSQLToFile(string sql, string defaultFileName)
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save SQL File",
                Filter = "SQL Files (*.sql)|*.sql",
                FileName = $"{defaultFileName}.sql", // Default file name
                DefaultExt = ".sql",                // Default file extension
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) // Default save location
            };

            // Show the dialog and save if confirmed
            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    File.WriteAllText(saveFileDialog.FileName, sql);
                    MessageBox.Show($"SQL file saved: {saveFileDialog.FileName}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error saving file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Save operation cancelled.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }


        /// <summary>
        /// Loads user settings from a configuration file.
        /// </summary>
        private void LoadSettings()
        {
            if (File.Exists(ConfigFilePath))
            {
                try
                {
                    var settings = JsonSerializer.Deserialize<UserSettings>(File.ReadAllText(ConfigFilePath));

                    // Restore ComboBox selections
                    OperationComboBox.SelectedIndex = settings.OperationIndex;
                    foreach (ComboBoxItem item in LanguageSelector.Items)
                    {
                        if (item.Tag?.ToString() == settings.LanguageCode)
                        {
                            item.IsSelected = true;
                            break;
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Could not load user settings. Defaults will be used.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        /// <summary>
        /// Saves user settings to a configuration file.
        /// </summary>
        private void SaveSettings()
        {
            var settings = new UserSettings
            {
                OperationIndex = OperationComboBox?.SelectedIndex ?? -1,
                LanguageCode = ((ComboBoxItem)LanguageSelector?.SelectedItem)?.Tag?.ToString() ?? string.Empty
            };

            var directoryPath = Path.GetDirectoryName(ConfigFilePath);
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            File.WriteAllText(ConfigFilePath, JsonSerializer.Serialize(settings));
        }

        /// <summary>
        /// Sets localized text for all UI elements.
        /// </summary>
        private void SetLocalizedText()
        {
            DragFileText.Text = LocalizationManager.GetString("DragFileHere");
            UploadButton.Content = LocalizationManager.GetString("UploadFile");
            GenerateSQLButton.Content = LocalizationManager.GetString("GenerateSQL");
        }
    }

    /// <summary>
    /// Represents the user settings to be saved and loaded.
    /// </summary>
    public class UserSettings
    {
        public int OperationIndex { get; set; }
        public required string LanguageCode { get; set; }
    }
}
