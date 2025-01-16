using System;
using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;

namespace ExcelToSQL
{
    public partial class MainWindow : Window
    {
        private const string ConfigFilePath = "C:\\ProgramData\\FloatyRock\\ExcelToSQL\\settings.json";

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
        /// Handles the drag-and-drop event for files.
        /// </summary>
        private void Border_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
        }

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
                Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv"
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
                // Placeholder: Replace with logic to parse and display the file's contents
                MessageBox.Show($"File loaded: {filePath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Generates SQL based on the current settings and loaded data.
        /// </summary>
        private void GenerateSQL_Click(object sender, RoutedEventArgs e)
        {
            if (DataPreviewGrid.ItemsSource == null)
            {
                MessageBox.Show("No data loaded. Please upload or drag a file.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                string operation = OperationComboBox.SelectedIndex == 0 ? "CREATE TABLE" : "INSERT INTO";
                MessageBox.Show($"SQL generation for {operation} completed!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating SQL: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
