using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Nop.Services.ExportImport;

using Nop.Core;
using Nop.Core.Domain.Catalog;
using Nop.Core.Domain.Directory;
using Nop.Core.Domain.Media;
using Nop.Core.Domain.Messages;
using Nop.Services.Catalog;
using Nop.Services.Directory;
using Nop.Services.Media;
using Nop.Services.Messages;
using Nop.Services.Seo;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using Product = Nop.Core.Domain.Catalog.Product;


namespace Nop.Ncc
{
    public class NccImportManager : ImportManager
    {
        /// <summary>
        /// Import products from XLSX file
        /// </summary>
        /// <param name="stream">Stream</param>
        public override void ImportProductsFromXlsx(Stream stream)
        {
            var productDatas = GetProductsProductDatas(stream);

            foreach (var exceledProductData in productDatas)
            {
                ProceesProduct(exceledProductData);
            }
        }

        public List<ExceledProductData> GetProductsProductDatas(Stream stream)
        {
            var result = new List<ExceledProductData>();

            using (var xlPackage = new ExcelPackage(stream))
            {
                // get the first worksheet in the workbook
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    throw new NopException("No worksheet found");
                }

                const int columsDataLength = 6;
                int startRow = 2;
                while (true)
                {
                    bool allColumnsAreEmpty = true;
                    for (var i = 1; i <= columsDataLength; i++)
                    {
                        if (worksheet.Cells[startRow, i].Value != null && 
                            !String.IsNullOrEmpty(worksheet.Cells[startRow, i].Value.ToString()))
                        {
                            allColumnsAreEmpty = false;
                            break;
                        }
                    }
                    if (allColumnsAreEmpty)
                    {
                        break;
                    }


                    result.Add(ConstructProduct(worksheet, startRow, 1));
                    result.Add(ConstructProduct(worksheet, startRow, 3));
                    result.Add(ConstructProduct(worksheet, startRow, 5));

                    //next 3 product
                    startRow++;
                }
            }
            
            return result;
        }

        #region .ctor
        public NccImportManager(IProductService productService,
            ICategoryService categoryService,
            IManufacturerService manufacturerService,
            IPictureService pictureService,
            IUrlRecordService urlRecordService,
            IStoreContext storeContext,
            INewsLetterSubscriptionService newsLetterSubscriptionService,
            ICountryService countryService,
            IStateProvinceService stateProvinceService) : base(productService, categoryService, manufacturerService, pictureService, urlRecordService, storeContext, newsLetterSubscriptionService, countryService, stateProvinceService)
        {
            
        }

        public NccImportManager()
            : base(null, null, null, null, null, null, null, null, null)
        {

        }
        #endregion


        public Picture GetPictureStrict(ExcelWorksheet worksheet, int row, int column , bool isNew)
        {

            var pictureRow = row - 1;
            var pictureColumn = column - 1;


            var pictureOnCoordinates = worksheet.Drawings.Where(p => ((p.To.Column == pictureColumn || p.To.Column == pictureColumn+1) && (p.To.Row == pictureRow)));

            ExcelDrawing excelDrawing = null;

            if (pictureOnCoordinates.Count() == 1)
            {
                excelDrawing = pictureOnCoordinates.FirstOrDefault();
            }
            else
            {
                //TODO: find upper image
                excelDrawing = pictureOnCoordinates.LastOrDefault();
            }

            //ExcelDrawing excelDrawing = worksheet.Drawings.FirstOrDefault(p => ((p.From.Column == column -1) && (p.From.Row == row-1)));

            

            var picture = excelDrawing as ExcelPicture;


            var stream = new MemoryStream();
            picture.Image.Save(stream, picture.ImageFormat);
            BinaryReader streamreader = new BinaryReader(stream);

            stream.Position = 0;

            var data = streamreader.ReadBytes((int) stream.Length);




            var pict = new Picture() { IsNew = isNew, PictureBinary = data, MimeType = picture.ImageFormat.ToString() };

            return pict;
        }

        public ExceledProductData ConstructProduct(ExcelWorksheet worksheet, int iRow, int column)
        {
         // pict.Image.Save(path);
            var priceColumn = column + 1;

            string name = Convert.ToString(worksheet.Cells[iRow, column].Value);
            string shortDescription = Convert.ToString(worksheet.Cells[iRow, column].Value);
            string fullDescription = Convert.ToString(worksheet.Cells[iRow, column].Value);

            string sku = Convert.ToString(worksheet.Cells[iRow, column].Value);

            decimal price = Convert.ToDecimal(worksheet.Cells[iRow, priceColumn].Value);

          //  string picture1 = Convert.ToString(worksheet.Cells[row, column].Value); //TODO: get picture
            
            

            Product product = null;

            if (_productService != null)
            {
                product = _productService.GetProductBySku(sku);
            }
            
            var isNew = false;
            if (product == null)
            {
                product = new Product();
                isNew = true;
            }

            product.Name = name;
            product.ShortDescription = shortDescription;
            product.FullDescription = fullDescription;

            product.Price = price;

            product.UpdatedOnUtc = DateTime.UtcNow;
            product.CreatedOnUtc = DateTime.UtcNow;


            //product.ProductPictures.Add(new ProductPicture(){});

            var picture = GetPictureStrict(worksheet, iRow, column, isNew);


            return new ExceledProductData
            {
                Product = product, InNew = isNew, Picture = picture
            };
        }

        public void ProceesProduct( ExceledProductData productData)
        {
            var newProduct = productData.InNew;
            var product = productData.Product;
            
            if (newProduct)
            {
                _productService.InsertProduct(product);
            }
            else
            {
                _productService.UpdateProduct(product);
            }
            #region
            //search engine name
           // _urlRecordService.SaveSlug(product, product.ValidateSeName(seName, product.Name, true), 0); TODO: victor, invistigate

            //category mappings
            //if (!String.IsNullOrEmpty(categoryIds))
            //{
            //    foreach (var id in categoryIds.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => Convert.ToInt32(x.Trim())))
            //    {
            //        if (product.ProductCategories.FirstOrDefault(x => x.CategoryId == id) == null)
            //        {
            //            //ensure that category exists
            //            var category = _categoryService.GetCategoryById(id);
            //            if (category != null)
            //            {
            //                var productCategory = new ProductCategory
            //                {
            //                    ProductId = product.Id,
            //                    CategoryId = category.Id,
            //                    IsFeaturedProduct = false,
            //                    DisplayOrder = 1
            //                };
            //                _categoryService.InsertProductCategory(productCategory);
            //            }
            //        }
            //    }
            //}

            //manufacturer mappings
            //if (!String.IsNullOrEmpty(manufacturerIds))
            //{
            //    foreach (var id in manufacturerIds.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => Convert.ToInt32(x.Trim())))
            //    {
            //        if (product.ProductManufacturers.FirstOrDefault(x => x.ManufacturerId == id) == null)
            //        {
            //            //ensure that manufacturer exists
            //            var manufacturer = _manufacturerService.GetManufacturerById(id);
            //            if (manufacturer != null)
            //            {
            //                var productManufacturer = new ProductManufacturer
            //                {
            //                    ProductId = product.Id,
            //                    ManufacturerId = manufacturer.Id,
            //                    IsFeaturedProduct = false,
            //                    DisplayOrder = 1
            //                };
            //                _manufacturerService.InsertProductManufacturer(productManufacturer);
            //            }
            //        }
            //    }
            //}

            #endregion

            //pictures
            //foreach (var picturePath in new[] { picture1, picture2, picture3 })

            if (productData.Picture == null)
            {
               
            }
            else
            {
                var mimeType =productData.Picture.MimeType;
                var newPictureBinary = productData.Picture.PictureBinary;
                var pictureAlreadyExists = false;
                if (!newProduct)
                {
                    //compare with existing product pictures
                    var existingPictures = _pictureService.GetPicturesByProductId(product.Id);
                    foreach (var existingPicture in existingPictures)
                    {
                        var existingBinary = _pictureService.LoadPictureBinary(existingPicture);
                        //picture binary after validation (like in database)
                        var validatedPictureBinary = _pictureService.ValidatePicture(newPictureBinary, mimeType);
                        if (existingBinary.SequenceEqual(validatedPictureBinary))
                        {
                            //the same picture content
                            pictureAlreadyExists = true;
                            break;
                        }
                    }
                }

                if (!pictureAlreadyExists)
                {
                    product.ProductPictures.Add(new ProductPicture
                    {
                        Picture = _pictureService.InsertPicture(newPictureBinary, mimeType, _pictureService.GetPictureSeName(product.Name), true),
                        DisplayOrder = 1,
                    });
                    _productService.UpdateProduct(product);
                }

            }

            //update "HasTierPrices" and "HasDiscountsApplied" properties
            _productService.UpdateHasTierPricesProperty(product);
            _productService.UpdateHasDiscountsApplied(product);
        }
    }
}
