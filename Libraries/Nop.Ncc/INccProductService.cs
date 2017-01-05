using System.Collections.Generic;
using Nop.Core.Domain.Catalog;
using Nop.Services.Catalog;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace Nop.Ncc
{
    public interface INccProductService : IProductService
    {
        void DeleteProducts(Product[] products);
        void InsertProducts(Product[] products);
        void UpdateProducts(Product[] products);
        
    }
}