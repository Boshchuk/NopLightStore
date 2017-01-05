using Nop.Core.Domain.Catalog;

namespace Nop.Ncc
{
    public class ProductProcessResult
    {
        public Product Product { get; set; }

        public NewProductType ProductType { get; set; }
    }
}