#region Usings
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Nop.Core;
using Nop.Core.Domain.Catalog;
using Nop.Core.Domain.Media;
using Nop.Services.Catalog;
using Nop.Services.Directory;
using Nop.Services.ExportImport;
using Nop.Services.Media;
using Nop.Services.Messages;
using Nop.Services.Seo;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

#endregion

namespace Nop.Ncc
{
    public class NccImportManager : ImportManager
    {
        #region .ctor
        public NccImportManager(IProductService productService,
            ICategoryService categoryService,
            IManufacturerService manufacturerService,
            IPictureService pictureService,
            IUrlRecordService urlRecordService,
            IStoreContext storeContext,
            INewsLetterSubscriptionService newsLetterSubscriptionService,
            ICountryService countryService,
            IStateProvinceService stateProvinceService)
            : base(productService, categoryService, manufacturerService, pictureService, urlRecordService, storeContext, newsLetterSubscriptionService, countryService, stateProvinceService)
        {

        }

        public NccImportManager()
            : base(null, null, null, null, null, null, null, null, null)
        {

        }
        #endregion

        #region temp unused

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
                var worksheet = xlPackage.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    throw new NopException("No worksheet found");
                }

                const int columsDataLength = 6;

                const int firstItemPos = 1;
                const int secondItemPos = 3;
                const int therdItemPos = 5;


                // TODO impliment search mechanithm to find start position
                var startRow = 2;
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

                    //TODO get file name
                    var fileName = "test"; // remove xslt

                    var cat = _categoryService.GetAllCategories(fileName); //TODO chek by == name

                    var catId = 0;


                    if (cat.Count == 0)
                    {
                        var categoryName = fileName;

                        var category = new Category
                        {
                            Name = categoryName
                        };

                        _categoryService.InsertCategory(category);

                        catId = _categoryService.GetAllCategories(categoryName).FirstOrDefault().Id;
                    }
                    else
                    {
                        catId = cat.FirstOrDefault().Id;
                    }




                    result.Add(ConstructProduct(worksheet, startRow, firstItemPos, catId));
                    result.Add(ConstructProduct(worksheet, startRow, secondItemPos, catId));
                    result.Add(ConstructProduct(worksheet, startRow, therdItemPos, catId));

                    //next 3 product
                    startRow++;
                }
            }

            return result;
        }

        #endregion

        public void InportInCategory(Stream stream, string fileName)
        {
            var productDatas = GetProductsProductDatas(stream,fileName);

            foreach (var exceledProductData in productDatas)
            {
                ProceesProduct(exceledProductData);
            }
        }

        public List<ExceledProductData> GetProductsProductDatas(Stream stream, string fileName)
        {
            var result = new List<ExceledProductData>();

            using (var xlPackage = new ExcelPackage(stream))
            {

                // get the first worksheet in the workbook
                var worksheet = xlPackage.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    throw new NopException("No worksheet found");
                }

                const int columsDataLength = 6;

                const int firstItemPos = 1;
                const int secondItemPos = 3;
                const int therdItemPos = 5;


                // TODO impliment search mechanithm to find start position
                var startRow = 2;
                while (true)
                {
                    var allColumnsAreEmpty = true;
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

                    //TODO get file name
                    var name = fileName.Replace(".xlsx",""); // remove xslt

                    var cat = _categoryService.GetAllCategories().FirstOrDefault(c => c.Name == name); //TODO chek by == name

                    var catId = 0;


                    if (cat == null)
                    {
                        var categoryName = name;

                        var category = new Category
                        {
                            Name = categoryName
                        };


                        category.CreatedOnUtc = DateTime.UtcNow;
                        category.UpdatedOnUtc = DateTime.UtcNow;
                        category.Published = true;
                        category.PageSize = 16;

                        _categoryService.InsertCategory(category);

                        var firstOrDefault = _categoryService.GetAllCategories().FirstOrDefault(c => c.Name == categoryName);
                        if (firstOrDefault != null)
                        {
                            catId = firstOrDefault.Id;
                        }
                    }
                    else
                    {
                        catId = cat.Id;
                    }
                    
                    result.Add(ConstructProduct(worksheet, startRow, firstItemPos, catId));
                    result.Add(ConstructProduct(worksheet, startRow, secondItemPos, catId));
                    result.Add(ConstructProduct(worksheet, startRow, therdItemPos, catId));

                    //next 3 product
                    startRow++;
                }
            }

            return result;
        }



        
        public Picture GetPictureStrict(ExcelWorksheet worksheet, int row, int column , bool isNew)
        {
            var pictureRow = row - 1;
            var pictureColumn = column - 1;


            var pictureOnCoordinates = worksheet.Drawings.Where(p => ((p.To.Column == pictureColumn || p.To.Column == pictureColumn+1) && (p.To.Row == pictureRow)));

            var excelDrawing = pictureOnCoordinates.Count() == 1 ? pictureOnCoordinates.FirstOrDefault() : pictureOnCoordinates.LastOrDefault();
            
            var picture = excelDrawing as ExcelPicture;


            var stream = new MemoryStream();
            picture.Image.Save(stream, picture.ImageFormat);
            var streamreader = new BinaryReader(stream);

            stream.Position = 0;

            var data = streamreader.ReadBytes((int) stream.Length);
            
            var pict = new Picture
            {
                IsNew = isNew, PictureBinary = data, MimeType = picture.ImageFormat.ToString()
            };

            return pict;
        }

        public ExceledProductData ConstructProduct(ExcelWorksheet worksheet, int iRow, int column, int categoryId)
        {
            var priceColumn = column + 1;

            var name = Convert.ToString(worksheet.Cells[iRow, column].Value);
            var shortDescription = Convert.ToString(worksheet.Cells[iRow, column].Value);
            var fullDescription = Convert.ToString(worksheet.Cells[iRow, column].Value);

            var sku = Convert.ToString(worksheet.Cells[iRow, column].Value);

            var price = Convert.ToDecimal(worksheet.Cells[iRow, priceColumn].Value);

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
            product.Published = true;

            
            product.ProductType = ProductType.SimpleProduct;
            product.VisibleIndividually = true;
         

            var picture = GetPictureStrict(worksheet, iRow, column, isNew);

          


            

            return new ExceledProductData
            {
                Product = product,
                InNew = isNew,
                Picture = picture,
                CategoryId = categoryId
            };
        }

        public void ProceesProduct(ExceledProductData productData)
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

            
            Category test = _categoryService.GetAllCategories().FirstOrDefault(c => c.Id == productData.CategoryId);
            
            if (test != null)
            {
                var productCategory = new ProductCategory
                {
                    ProductId = product.Id,
                    CategoryId = productData.CategoryId,
                  //DisplayOrder = model.DisplayOrder
                };
                 product.ProductCategories.Add(productCategory);
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
