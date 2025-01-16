using System.Globalization;
using System.Resources;

namespace ExcelToSQL
{
    public static class LocalizationManager
    {
        // ResourceManager connected to the Resources.resx file
        private static ResourceManager resourceManager = new ResourceManager("ExcelToSQL.Resources.Resources", typeof(LocalizationManager).Assembly);

        /// <summary>
        /// Retrieves a localized string for the given key.
        /// </summary>
        /// <param name="key">The key to look up in the resource file.</param>
        /// <returns>The localized string or the key itself if not found.</returns>
        public static string GetString(string key)
        {
            try
            {
                return resourceManager.GetString(key, CultureInfo.CurrentUICulture) ?? key;
            }
            catch
            {
                return key; // Return the key as a fallback
            }
        }

        /// <summary>
        /// Sets the current language for the application.
        /// </summary>
        /// <param name="cultureCode">The culture code, e.g., "en", "es", "fr".</param>
        public static void SetLanguage(string cultureCode)
        {
            CultureInfo.CurrentUICulture = new CultureInfo(cultureCode);
        }
    }
}
