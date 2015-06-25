namespace Nop.Ncc
{
    /// <summary>
    /// Heleper methods for import functionality
    /// </summary>
    public static class ImportHelper
    {
        public const string CatalogCategoryName = "Каталог";

        public const string ExistingInStore = "Товары в магазине";


        /// <summary>
        /// Constructs Category name from file name
        /// </summary>
        /// <param name="fileName">File name</param>
        /// <returns>Categor name</returns>
        public static string ConstractCategoryName(string fileName)
        {
            return fileName.Replace(".xlsx", "");
        }
    }
}