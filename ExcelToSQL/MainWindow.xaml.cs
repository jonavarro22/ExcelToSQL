using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using CsvHelper;
using OfficeOpenXml;
using System.Data;

namespace ExcelToSQL
{
    public partial class MainWindow : Window
    {
        private const string ConfigFilePath = "C:\\ProgramData\\ExcelToSQL\\settings.json";
        private string selectedSQLType = "MSSQL"; // Default to MSSQL
        private bool ComponentsInitialized = false;

        private string DelimiterWarningText { get; set; }
        private string NoSheetSelectedText { get; set; }
        private string UnsupportedFileText { get; set; }
        private string SelectSheetText { get; set; }
        private string SelectSheetTitle { get; set; }
        private string AllSupportedFilesText { get; set; }
        private string ErrorLoadingFileText { get; set; }
        private string DelimiterAutoText { get; set; }
        private string NoDataLoadedErrorText { get; set; }
        private string ErrorGeneratingSQLText { get; set; }
        private string SQLFileSavedText { get; set; }
        private string ErrorSavingFileText { get; set; }
        private string SaveOperationCancelledText { get; set; }
        private string CantLoadSettingsText { get; set; }
        private string SaveSQLFileTitle { get; set; }
        private string InputFilePath { get; set; }
        private string OutputFilePath { get; set; }


        public MainWindow()
        {
            InitializeComponent();
            ComponentsInitialized = true;
            LoadSettings(); // Load saved settings on startup
            SetLocalizedText(); // Set default localized text
        }

        /// <summary>
        /// Handles changes to the Operation ComboBox (Create Table or Update Table).
        /// </summary>
        private void OperationComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ComponentsInitialized)
            {
                SaveSettings(); // Save the current selection whenever it changes
            }
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
                    if (ComponentsInitialized)
                    {
                        SaveSettings(); // Save the selected language whenever it changes
                    }
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
                    if (ComponentsInitialized)
                    {
                        SaveSettings(); // Save the selected language whenever it changes
                    }
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
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = AllSupportedFilesText + " (*.xlsx, *.csv)|*.xlsx;*.csv|Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv",
                InitialDirectory = !string.IsNullOrEmpty(InputFilePath)
                    ? Path.GetDirectoryName(InputFilePath)
                    : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) // Default to My Documents
            };

            if (openFileDialog.ShowDialog() == true)
            {
                InputFilePath = openFileDialog.FileName; // Store the selected file's full path
                SaveSettings(); // Save the updated InputFilePath
                LoadFile(InputFilePath);
            }
        }


        /// <summary>
        /// Loads the selected file and displays its content in the DataGrid.
        /// </summary>
        private void LoadFile(string filePath)
        {
            try
            {

                if (filePath.EndsWith(".csv"))
                {
                    // CSV processing remains unchanged
                    string delimiter = DetectOrGetSelectedDelimiter(filePath);

                    if (string.IsNullOrEmpty(delimiter))
                    {
                        MessageBox.Show(DelimiterWarningText, "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                        var sheetNames = package.Workbook.Worksheets.Select(ws => ws.Name).ToList();

                        // Prompt the user to select a sheet if multiple exist
                        string selectedSheet = SelectSheet(sheetNames);
                        if (string.IsNullOrEmpty(selectedSheet))
                        {
                            MessageBox.Show(NoSheetSelectedText, "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                            return;
                        }

                        var worksheet = package.Workbook.Worksheets[selectedSheet];
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
                    MessageBox.Show(UnsupportedFileText, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ErrorLoadingFileText + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Prompts the user to select a sheet from the Excel file.
        /// </summary>
        /// <param name="sheetNames"></param>
        /// <returns></returns>
        private string SelectSheet(List<string> sheetNames)
        {
            var sheetSelectionWindow = new Window
            {
                Title = SelectSheetTitle,
                Width = 300,
                Height = 200,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var stackPanel = new StackPanel { Margin = new Thickness(10) };

            var comboBox = new ComboBox { ItemsSource = sheetNames, Margin = new Thickness(0, 0, 0, 10) };
            comboBox.SelectedIndex = 0;

            var okButton = new Button { Content = "OK", Width = 80, HorizontalAlignment = HorizontalAlignment.Center };
            okButton.Click += (s, e) => sheetSelectionWindow.DialogResult = true;

            stackPanel.Children.Add(new TextBlock { Text = SelectSheetText, Margin = new Thickness(0, 0, 0, 10) });
            stackPanel.Children.Add(comboBox);
            stackPanel.Children.Add(okButton);

            sheetSelectionWindow.Content = stackPanel;

            if (sheetSelectionWindow.ShowDialog() == true)
            {
                return comboBox.SelectedItem as string;
            }

            return null; // User cancelled
        }


        /// <summary>
        /// Handles the infer column types button click.
        /// </summary>
        private DataTable InferColumnTypes(DataTable originalTable)
        {
            var newTable = new DataTable();

            foreach (DataColumn column in originalTable.Columns)
            {
                var nonNullTrimmedValues = originalTable.AsEnumerable()
                    .Where(row =>
                        !row.IsNull(column) &&
                        !string.IsNullOrWhiteSpace(row[column].ToString()) &&
                        !string.Equals(row[column].ToString().Trim(), "NULL", StringComparison.OrdinalIgnoreCase)
                    )
                    .Select(row => row[column].ToString().Trim())
                    .ToList();

                Type columnType;

                // If there are no actual (non-whitespace) values, default to string
                if (!nonNullTrimmedValues.Any())
                {
                    columnType = typeof(int);
                }
                else if (nonNullTrimmedValues.All(v => Guid.TryParse(v, out _)))
                {
                    columnType = typeof(Guid);
                }
                else if (nonNullTrimmedValues.All(v => DateTime.TryParse(v, out _)))
                {
                    columnType = typeof(DateTime);
                }
                //else if (nonNullTrimmedValues.All(v => TimeSpan.TryParse(v, out _)))
                //{
                //    columnType = typeof(DateTime);
                //}
                else if (nonNullTrimmedValues.All(v => int.TryParse(v, out _)))
                {
                    columnType = typeof(int);
                }
                else if (nonNullTrimmedValues.All(v => decimal.TryParse(v, out _)))
                {
                    columnType = typeof(decimal);
                }
                else if (nonNullTrimmedValues.All(v => double.TryParse(v, out _)))
                {
                    columnType = typeof(double);
                }
                else if (nonNullTrimmedValues.All(v => bool.TryParse(v, out _)))
                {
                    columnType = typeof(bool);
                }
                else if (nonNullTrimmedValues.All(v => byte.TryParse(v, out _)))
                {
                    columnType = typeof(byte);
                }
                else
                {
                    columnType = typeof(string);
                }

                // Create the column in the new table
                DataColumn newColumn = newTable.Columns.Add(column.ColumnName, columnType);
                // Now set AllowDBNull for that column
                newColumn.AllowDBNull = true;  // or false, based on your requirements
            }

            // Copy data
            foreach (DataRow row in originalTable.Rows)
            {
                var newRow = newTable.NewRow();
                foreach (DataColumn column in originalTable.Columns)
                {
                    // Remove the extra assignment that overwrote your attempt at DBNull
                    if (row[column]?.ToString().Trim().Equals("NULL", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        newRow[column.ColumnName] = DBNull.Value;
                    }
                    else
                    {
                        newRow[column.ColumnName] = row[column];
                    }
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
                        MessageBox.Show(DelimiterAutoText, "Info", MessageBoxButton.OK, MessageBoxImage.Information);
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
            switch (type)
            {
                case Type t when t == typeof(int):
                    switch (selectedSQLType)
                    {
                        case "MSSQL": return "INT";
                        case "MySQL": return "INT";
                        case "PostgreSQL": return "INTEGER";
                        default: return "INT";
                    }

                case Type t when t == typeof(double):
                    switch (selectedSQLType)
                    {
                        case "MSSQL": return "FLOAT";
                        case "MySQL": return "DOUBLE";
                        case "PostgreSQL": return "DOUBLE PRECISION";
                        default: return "FLOAT";
                    }

                case Type t when t == typeof(decimal):
                    switch (selectedSQLType)
                    {
                        case "MSSQL": return "DECIMAL";
                        case "MySQL": return "DECIMAL";
                        case "PostgreSQL": return "DECIMAL";
                        default: return "DECIMAL";
                    }

                case Type t when t == typeof(bool):
                    switch (selectedSQLType)
                    {
                        case "MSSQL": return "BIT";
                        case "MySQL": return "TINYINT(1)";
                        case "PostgreSQL": return "BOOLEAN";
                        default: return "BIT";
                    }

                case Type t when t == typeof(byte):
                    switch (selectedSQLType)
                    {
                        case "MSSQL": return "TINYINT";
                        case "MySQL": return "TINYINT UNSIGNED";
                        case "PostgreSQL": return "SMALLINT";
                        default: return "TINYINT";
                    }

                case Type t when t == typeof(DateTime):
                    switch (selectedSQLType)
                    {
                        case "MSSQL": return "DATETIME2"; // DATETIME2 is generally recommended over DATETIME
                        case "MySQL": return "DATETIME";
                        case "PostgreSQL": return "TIMESTAMP WITHOUT TIME ZONE"; // You might also use TIMESTAMP WITH TIME ZONE
                        default: return "TIMESTAMP";
                    }

                case Type t when t == typeof(Guid):
                    switch (selectedSQLType)
                    {
                        case "MSSQL": return "UNIQUEIDENTIFIER";
                        case "MySQL": return "CHAR(36)";
                        case "PostgreSQL": return "UUID";
                        default: return "NVARCHAR(MAX)";
                    }

                default:
                    switch (selectedSQLType)
                    {
                        case "MSSQL": return "NVARCHAR(MAX)";
                        case "MySQL": return "TEXT";
                        case "PostgreSQL": return "TEXT";
                        default:
                            return "NVARCHAR(MAX)";
                    }
            }
        }

        /// <summary>
        /// Generates SQL based on the current settings and loaded data.
        /// </summary>
        private void GenerateSQL_Click(object sender, RoutedEventArgs e)
        {
            if (DataPreviewGrid.ItemsSource is null)
            {
                MessageBox.Show(NoDataLoadedErrorText, "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var table = ((DataView)DataPreviewGrid.ItemsSource).ToTable();
            bool isCreateTable = OperationComboBox.SelectedIndex == 0; // 0 = Create Table, 1 = Update Table
            string fileName = Path.GetFileNameWithoutExtension(InputFilePath) ?? "GeneratedSQL"; // Use InputFilePath or default name

            try
            {
                // Generate the SQL script
                var sql = GenerateSQL(table, isCreateTable, fileName);

                // Determine the save path
                if (string.IsNullOrEmpty(OutputFilePath))
                {
                    // Default to My Documents if OutputFilePath is not set
                    OutputFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), $"{fileName}.sql");
                }
                else if (!OutputFilePath.EndsWith(".sql", StringComparison.OrdinalIgnoreCase))
                {
                    // Ensure the file name ends with .sql
                    OutputFilePath = Path.Combine(OutputFilePath, $"{fileName}.sql");
                }

                // Save the SQL script
                SaveSQLToFile(sql, OutputFilePath);
                SaveSettings();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ErrorGeneratingSQLText + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }



        /// <summary>
        /// Saves the generated SQL to a file. Prompts the user for a file name and location.
        /// </summary>
        private void SaveSQLToFile(string sql, string defaultFileName)
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = SaveSQLFileTitle,
                Filter = "SQL (*.sql)|*.sql",
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
                    MessageBox.Show(SQLFileSavedText + saveFileDialog.FileName, "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ErrorSavingFileText + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show(SaveOperationCancelledText, "Info", MessageBoxButton.OK, MessageBoxImage.Information);
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
                    // Load settings from the configuration file
                    var settings = JsonSerializer.Deserialize<UserSettings>(File.ReadAllText(ConfigFilePath));

                    OperationComboBox.SelectedIndex = settings.OperationIndex;
                    foreach (ComboBoxItem item in LanguageSelector.Items)
                    {
                        if (item.Tag?.ToString() == settings.LanguageCode)
                        {
                            item.IsSelected = true;
                            break;
                        }
                    }
                    TargetSQLComboBox.SelectedIndex = settings.TargetSQLIndex;
                    InputFilePath = settings.InputPath;
                    OutputFilePath = settings.OutputPath;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading settings: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                try
                {
                    // Create and write default settings
                    var defaultSettings = new UserSettings
                    {
                        InputPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        OutputPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        OperationIndex = 0, // Default to "Create Table" or your first operation
                        LanguageCode = "en", // Default language
                        TargetSQLIndex = 0 // Default SQL type index (e.g., MSSQL)
                    };

                    // Ensure the directory exists
                    string directoryPath = Path.GetDirectoryName(ConfigFilePath);
                    if (!Directory.Exists(directoryPath))
                    {
                        Directory.CreateDirectory(directoryPath);
                    }

                    // Write the default settings to the file
                    File.WriteAllText(ConfigFilePath, JsonSerializer.Serialize(defaultSettings));
                    MessageBox.Show("Default settings file created.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);

                    // Apply the default settings
                    InputFilePath = defaultSettings.InputPath;
                    OutputFilePath = defaultSettings.OutputPath;
                    OperationComboBox.SelectedIndex = defaultSettings.OperationIndex;
                    TargetSQLComboBox.SelectedIndex = defaultSettings.TargetSQLIndex;
                    foreach (ComboBoxItem item in LanguageSelector.Items)
                    {
                        if (item.Tag?.ToString() == defaultSettings.LanguageCode)
                        {
                            item.IsSelected = true;
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error creating default settings: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
                LanguageCode = ((ComboBoxItem)LanguageSelector?.SelectedItem)?.Tag?.ToString() ?? string.Empty,
                TargetSQLIndex = TargetSQLComboBox?.SelectedIndex ?? -1,
                InputPath = InputFilePath ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                OutputPath = OutputFilePath ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            try
            {
                File.WriteAllText(ConfigFilePath, JsonSerializer.Serialize(settings));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving settings: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        /// <summary>
        /// Sets localized text for all UI elements.
        /// </summary>
        private void SetLocalizedText()
        {
            DragFileText.Text = LocalizationManager.GetString("DragFileHere");
            UploadButton.Content = LocalizationManager.GetString("UploadFile");
            GenerateSQLButton.Content = LocalizationManager.GetString("GenerateSQL");
            DelimiterText.Text = LocalizationManager.GetString("Delimiter");
            OperationText.Text = LocalizationManager.GetString("Operation");
            LanguageText.Text = LocalizationManager.GetString("Language");
            TargetSQLText.Text = LocalizationManager.GetString("TargetSQL");
            UpdateTableOption.Content = LocalizationManager.GetString("UpdateTable");
            CreateTableOption.Content = LocalizationManager.GetString("CreateTable");
            DelimiterWarningText = LocalizationManager.GetString("DelimiterWarning");
            NoSheetSelectedText = LocalizationManager.GetString("NoSheetSelected");
            UnsupportedFileText = LocalizationManager.GetString("UnsupportedFile");
            SelectSheetText = LocalizationManager.GetString("SelectSheet");
            SelectSheetTitle = LocalizationManager.GetString("SelectSheetTitle");
            AllSupportedFilesText = LocalizationManager.GetString("AllSupportedFiles");
            ErrorLoadingFileText = LocalizationManager.GetString("ErrorLoadingFile");
            DelimiterAutoText = LocalizationManager.GetString("DelimiterAuto");
            NoDataLoadedErrorText = LocalizationManager.GetString("NoDataLoadedError");
            ErrorGeneratingSQLText = LocalizationManager.GetString("ErrorGeneratingSQL");
            SQLFileSavedText = LocalizationManager.GetString("SQLFileSaved");
            ErrorSavingFileText = LocalizationManager.GetString("ErrorSavingFile");
            SaveOperationCancelledText = LocalizationManager.GetString("SaveOperationCancelled");
            CantLoadSettingsText = LocalizationManager.GetString("CantLoadSettings");
            SaveSQLFileTitle = LocalizationManager.GetString("SaveSQLFileTitle");

        }
    }

    /// <summary>
    /// Represents the user settings to be saved and loaded.
    /// </summary>
    public class UserSettings
    {
        public int OperationIndex { get; set; }
        public string LanguageCode { get; set; }
        public int TargetSQLIndex { get; set; }
        public string InputPath { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        public string OutputPath { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    }

}
