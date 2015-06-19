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
    public static class CategoryExtend
    {
        public static Category GetCategoryByName(this ICategoryService categoryService, string catalogCategoryName)
        {
            return categoryService.GetAllCategories().FirstOrDefault(c => c.Name == catalogCategoryName);
        }
    }

    public class NccImportManager : ImportManager
    {
        private const string ErrorMessage = "Неправильный формат файла. Листы в нем не обнаружены";

        private const string CatalogCategoryName = "Каталог";

        private const string ExistingInStore = "Товары в магазине";

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

        /// <summary>
        /// Import products from XLSX file
        /// </summary>
        /// <param name="stream">Stream</param>
        public override void ImportProductsFromXlsx(Stream stream)
        {
            var productDatas = GetProductsProductDatas(stream, ExistingInStore, false);
            ProcessData(productDatas);

            UpdateCategoryImage(ExistingInStore);
        }
        
        public void InportInCategory(Stream stream, string fileName)
        {
            var productDatas = GetProductsProductDatas(stream, fileName, true);
            ProcessData(productDatas);

            var categoryName = ConstractCategoryName(fileName);

            // for category and ...
            UpdateCategoryImage(categoryName);

            // for root category
            var category = _categoryService.GetCategoryByName(categoryName);
            if (category.ParentCategoryId != 0)
            {
                UpdateCategoryImage(category.Id, category.ParentCategoryId);
            }
        }

        private void UpdateCategoryImage(int categoryId)
        {
            var category = _categoryService.GetCategoryById(categoryId);
            UpdateCategoryImage(category.Id, category.Id);
        }

        private void UpdateCategoryImage(string categoryName)
        {
            var categoryId = _categoryService.GetCategoryByName(categoryName).Id;
            UpdateCategoryImage(categoryId, categoryId);
        }

        private void UpdateCategoryImage(int categoryToGetProductsId, int categoryIdToChangeImage)
        {
            var productCategory = _categoryService.GetProductCategoriesByCategoryId(categoryToGetProductsId, 0, 1000000).FirstOrDefault();

            var product = _productService.GetProductById(productCategory.ProductId);

            if (product != null)
            {
                var productPicture = product.ProductPictures.FirstOrDefault();
                if (productPicture != null)
                {
                    var categortToUpdate = _categoryService.GetCategoryById(categoryIdToChangeImage);
                    categortToUpdate.PictureId = productPicture.PictureId;
                    _categoryService.UpdateCategory(categortToUpdate);
                }
            }
        }

        private string ConstractCategoryName(string fileName)
        {
            return  fileName.Replace(".xlsx", "");
        }

        private void AddInitCategory(string categoryName)
        {
            // TODO: invistigate is this possible to have one method for category creation??? 
            // adding part
            var category = new Category
            {
                Name = categoryName,
                CreatedOnUtc = DateTime.UtcNow,
                UpdatedOnUtc = DateTime.UtcNow,
                Published = true,
                PageSize = 16,
                ShowOnHomePage = true,
                IncludeInTopMenu = true,
                DisplayOrder = 1,

            };
            _categoryService.InsertCategory(category);


            var seName = category.ValidateSeName(category.Name, category.Name, true);
            _urlRecordService.SaveSlug(category, seName, 0);
        }

        private List<ExceledProductData> GetProductsProductDatas(Stream stream, string fileName, bool addInCatalog = false)
        {
            var result = new List<ExceledProductData>();

            using (var xlPackage = new ExcelPackage(stream))
            {
                // get the first worksheet in the workbook
                var worksheet = xlPackage.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    throw new NopException(ErrorMessage);
                }

                const int columsDataLength = 6;

                const int firstItemPos = 1;
                const int secondItemPos = 3;
                const int therdItemPos = 5;

                var catalogName = ConstractCategoryName(fileName);

                var rootCatId = 0;
                if (addInCatalog)
                {
                    var rootCatalogCategory = _categoryService.GetAllCategories().FirstOrDefault(c => c.Name == CatalogCategoryName);

                    if (rootCatalogCategory != null)
                    {
                        rootCatId = rootCatalogCategory.Id;
                    }
                    else
                    {
                        AddInitCategory(CatalogCategoryName);
                        var category = _categoryService.GetAllCategories().FirstOrDefault(c => c.Name == CatalogCategoryName);
                        if (category != null)
                        {
                            rootCatId = category.Id;
                        }
                    }
                }

                var cat = _categoryService.GetAllCategories().FirstOrDefault(c => c.Name == catalogName); //TODO chek by == name
                var catId = 0;

                if (cat != null)
                {
                    catId = cat.Id;

                    var list = _categoryService.GetProductCategoriesByCategoryId(catId, 0, 1000000);

                    var ids = list.Select(c => c.ProductId);

                    foreach (var id in ids)
                    {
                        var productToDelete = _productService.GetProductById(id);
                        _productService.DeleteProduct(productToDelete);
                    }
                }
                else
                {
                    var categoryName = catalogName;

                    var category = new Category
                    {
                        Name = categoryName,
                        CreatedOnUtc = DateTime.UtcNow,
                        UpdatedOnUtc = DateTime.UtcNow,
                        Published = true,
                        PageSize = 16,
                        ShowOnHomePage = true,// TODO: invistigate why its not dispalyed,
                        IncludeInTopMenu = true,
                        DisplayOrder = 1,
                    };

                    if (addInCatalog)
                    {
                        category.ParentCategoryId = rootCatId;
                    }
                    
                    _categoryService.InsertCategory(category);

                    var seName = category.ValidateSeName(category.Name, category.Name, true);
                    _urlRecordService.SaveSlug(category, seName, 0);


                    var firstOrDefault = _categoryService.GetAllCategories().FirstOrDefault(c => c.Name == categoryName);
                    if (firstOrDefault != null)
                    {
                        catId = firstOrDefault.Id;
                    }
                }

                // TODO impliment search mechanithm to find start position
                var startRow = 2;

                var callForPrice = addInCatalog;

                var skipList = new List<int>();

                while (true)
                {
                    var allColumnsAreEmpty = true;
                    for (var i = 1; i <= columsDataLength; i++)
                    {
                        if (worksheet.Cells[startRow, i].Value != null &&
                            !String.IsNullOrEmpty(worksheet.Cells[startRow, i].Value.ToString()))
                        {
                            allColumnsAreEmpty = false;
                            //break;
                        }
                        else
                        {
                            skipList.Add(i);
                        }

                    }
                    if (allColumnsAreEmpty)
                    {
                        break;
                    }
                    if (!skipList.Contains(firstItemPos))
                    {
                        result.Add(ConstructProduct(worksheet, startRow, firstItemPos, catId, callForPrice));
                    }
                  
                    result.Add(ConstructProduct(worksheet, startRow, secondItemPos, catId, callForPrice));
                    result.Add(ConstructProduct(worksheet, startRow, therdItemPos, catId, callForPrice));

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

        public ExceledProductData ConstructProduct(ExcelWorksheet worksheet, int iRow, int column, int categoryId, bool callForPrice = false)
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
            product.CallForPrice = callForPrice;

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

            var seName = product.ValidateSeName(product.Name, product.Name, true);
            _urlRecordService.SaveSlug(product, seName, 0);

            var category = _categoryService.GetAllCategories().FirstOrDefault(c => c.Id == productData.CategoryId);
            
            if (category != null)
            {
                var productCategory = new ProductCategory
                {
                    ProductId = product.Id,
                    CategoryId = productData.CategoryId,
                  //DisplayOrder = model.DisplayOrder
                };
                product.ProductCategories.Add(productCategory);
            }

            // TODO: delete unnided
            #region previso version
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

        private void ProcessData(IEnumerable<ExceledProductData> productDatas)
        {
            foreach (var exceledProductData in productDatas)
            {
                ProceesProduct(exceledProductData);
            }
        }
    }
}
