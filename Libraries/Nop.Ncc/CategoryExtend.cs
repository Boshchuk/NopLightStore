using System.Linq;
using Nop.Core.Domain.Catalog;
using Nop.Services.Catalog;

namespace Nop.Ncc
{
    public static class CategoryExtend
    {
        /// <summary>
        /// Returns category by name
        /// Takes First or Default
        /// </summary>
        /// <param name="categoryService">Category Service implimentation</param>
        /// <param name="catalogCategoryName">Category name</param>
        /// <returns></returns>
        public static Category GetCategoryByName(this ICategoryService categoryService, string catalogCategoryName)
        {
            return categoryService.GetAllCategories().FirstOrDefault(c => c.Name == catalogCategoryName);
        }
    }
}