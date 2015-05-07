using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Nop.Services.ExportImport;

namespace Nop.Ncc
{
    public class NccImportManager : IImportManager
    {
        /// <summary>
        /// Import products from XLSX file
        /// </summary>
        /// <param name="stream">Stream</param>
        public void ImportProductsFromXlsx(Stream stream)
        {
            throw new NotImplementedException();
        }

        public int ImportNewsletterSubscribersFromTxt(Stream stream)
        {
            throw new NotImplementedException();
        }

        public int ImportStatesFromTxt(Stream stream)
        {
            throw new NotImplementedException();
        }
    }
}
