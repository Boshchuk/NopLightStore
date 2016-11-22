using System;
using System.IO;
//using Nop.Ncc;
using Nop.Services.Catalog;
using Nop.Services.Media;
using Nop.Services.Seo;
using Nop.Services.Tasks;


namespace Nop.Services.ExportImport
{
    //used by type name
    public partial class ImportProductFromFolderTask : ITask
    {
        private readonly IProductService _productService;
        private readonly ICategoryService _categoryService;
        private readonly IManufacturerService _manufacturerService;

        private readonly IPictureService _pictureService;
        private readonly IUrlRecordService _urlRecordService;

        public ImportProductFromFolderTask(IProductService productService, ICategoryService categoryService, IManufacturerService manufacturerService, IManufacturerService manufacturerService1, IPictureService pictureService, IUrlRecordService urlRecordService)
        {
            _productService = productService;
            _categoryService = categoryService;
            _manufacturerService = manufacturerService1;
            _pictureService = pictureService;
            _urlRecordService = urlRecordService;
        }


        public void Execute()
        {
            //var importManager = new NccImportManager(_productService,
            //this._categoryService,
            //this._manufacturerService,
            //this._pictureService,
            //this._urlRecordService,
            //null,
            //null,
            //null,
            //null);

            //var path = System.Configuration.ConfigurationSettings.AppSettings["CatalogLocation"];

            //string[] files = System.IO.Directory.GetFiles(path, "*.xlsx", SearchOption.AllDirectories);

            //foreach (string filePath in files)
            //{

            //    using (var stream = new FileStream(filePath, FileMode.Open))
            //    {
            //        var pos = filePath.IndexOf("\\");
            //        var fileName = filePath.Substring(pos + 2);
            //        importManager.InportInCatalog(stream, fileName);
            //    }

            //    //  var fileName = file.FileName;


            //}
        }
    }
}