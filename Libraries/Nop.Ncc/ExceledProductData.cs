using Nop.Core.Domain.Catalog;
using Nop.Core.Domain.Media;

namespace Nop.Ncc
{
    public class ExceledProductData
    {
        public Product Product { get; set; }

        public bool IsNew { get; set; }

        public Picture Picture { get; set; }

        public int CategoryId { get; set; }
    }
}