﻿#region Usings

using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        private const string ErrorMessage = "Неправильный формат файла. Листы в нем не обнаружены";

        protected readonly INccProductService _productService;


        // TODO: почитать что такое DI IOC
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

        public NccImportManager(INccProductService productService,
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
            _productService = productService;
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
            var productDatas = GetProductsProductDatas(stream, ImportHelper.ExistingInStore, false);
            ProcessData(productDatas);

            UpdateCategoryImage(ImportHelper.ExistingInStore);
        }

        /// <summary>
        /// Import Product in catalog
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="fileName"></param>
        public void InportInCatalog(Stream stream, string fileName)
        {
            var productDatas = GetProductsProductDatas(stream, fileName, true);
            ProcessData(productDatas);

            var categoryName = ImportHelper.ConstractCategoryName(fileName);

            // for category and ...
            UpdateCategoryImage(categoryName);

            // for root category
            var category = _categoryService.GetCategoryByName(categoryName);
            if (category.ParentCategoryId != 0)
            {
                UpdateCategoryImage(category.Id, category.ParentCategoryId);
            }
        }

        /// <summary>
        /// Add some of base needed category for product
        /// </summary>
        /// <param name="categoryName">Category Name </param>
        /// <param name="showOnHomePage">Show on home page</param>
        private void AddInitCategory(string categoryName, bool showOnHomePage = true)
        {
            // TODO: invistigate is this possible to have one method for category creation??? 
            var category = new Category
            {
                Name = categoryName,
                CreatedOnUtc = DateTime.UtcNow,
                UpdatedOnUtc = DateTime.UtcNow,
                Published = true,
                PageSize = 16,
                ShowOnHomePage = showOnHomePage,
                IncludeInTopMenu = true,
                DisplayOrder = 1,

            };
            _categoryService.InsertCategory(category);
            
            // uri part
            var seName = category.ValidateSeName(category.Name, category.Name, true);
            _urlRecordService.SaveSlug(category, seName, 0);
        }

        /// <summary>
        /// Exctract data from exel
        /// </summary>
        /// <param name="stream">Steam from excel file</param>
        /// <param name="fileName">File name</param>
        /// <param name="addInCatalog">Add to catalog or use like exist in store</param>
        /// <returns></returns>
        private IEnumerable<ExceledProductData> GetProductsProductDatas(Stream stream, string fileName, bool addInCatalog = false)
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

                var catalogName = ImportHelper.ConstractCategoryName(fileName);

                var rootCatId = 0;

                // add aditional CATALOG category as root category
                if (addInCatalog)
                {
                    var rootCatalogCategory = _categoryService.GetCategoryByName(ImportHelper.CatalogCategoryName);

                    if (rootCatalogCategory != null)
                    {
                        rootCatId = rootCatalogCategory.Id;
                    }
                    else
                    {
                        // we dont want display it on home page so second parameter is fales
                        AddInitCategory(ImportHelper.CatalogCategoryName, false);
                        var category = _categoryService.GetCategoryByName(ImportHelper.CatalogCategoryName);
                        if (category != null)
                        {
                            rootCatId = category.Id;
                        }
                    }
                }

                var cat = _categoryService.GetCategoryByName(catalogName);
                var catId = 0;

                if (cat != null)
                {
                    catId = cat.Id;

                    var list = _categoryService.GetProductCategoriesByCategoryId(catId, 0, 1000000);
                    if (list.Count > 0)
                    {
                        var ids = list.Select(c => c.ProductId);
                        var productsToDelete = new List<Product>();

                        productsToDelete = _productService.GetProductsByIds(ids.ToArray()).ToList();
                        _productService.DeleteProducts(productsToDelete.ToArray());
                    }
                    
                    //foreach (var id in ids)
                    //{
                    //    var productToDelete = _productService.GetProductById(id);

                    //    productToDelete.Deleted = true;
                    //    productsToDelete.Add(productToDelete);

                    //}
                    //foreach (var product in productsToDelete)
                    //{
                    //    _productService.DeleteProducts();
                    //}

                    //_productService.UpdateProducts(productsToDelete.ToArray());
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
                        ShowOnHomePage = true,
                        IncludeInTopMenu = true,
                        DisplayOrder = 1,
                    };

                    // Add as child category if we have root category
                    if (addInCatalog)
                    {
                        category.ParentCategoryId = rootCatId;
                    }
                    
                    _categoryService.InsertCategory(category);

                    var seName = category.ValidateSeName(category.Name, category.Name, true);
                    _urlRecordService.SaveSlug(category, seName, 0);


                    var firstOrDefault = _categoryService.GetCategoryByName(categoryName);
                    if (firstOrDefault != null)
                    {
                        catId = firstOrDefault.Id;
                    }
                }

                // TODO impliment search mechanithm to find start position
                var startRow = 2;

                var callForPrice = addInCatalog;

                var skipList = new List<int>();

                var dispalyOrder = 1;
                var attempt = 0;


                while (true)
                {
                    var allColumnsAreEmpty = true;
                    for (var i = 1; i <= columsDataLength; i++)
                    {
                        if (worksheet.Cells[startRow, i].Value != null &&
                            !string.IsNullOrEmpty(worksheet.Cells[startRow, i].Value.ToString()))
                        {
                            allColumnsAreEmpty = false;
                            attempt = 0;
                            //break;
                        }
                        else
                        {
                            skipList.Add(i);
                        }

                    }
                    if (allColumnsAreEmpty)
                    {
                        attempt++;
                        if (attempt > 1)
                        {
                            break;
                        }
                    }

                    // if we have additional header
                    if (skipList.Count == 5)
                    {
                        // skipping all try to add
                        skipList.Clear();
                        startRow++;
                        continue;
                    }

                    if (!skipList.Contains(firstItemPos))
                    {
                        var constructProduct = ConstructProduct(worksheet, startRow, firstItemPos, catId, dispalyOrder,callForPrice);
                        if (constructProduct != null)
                        {
                            result.Add(constructProduct);
                        }
                        
                    }
                    dispalyOrder++;

                    Trace.WriteLine(dispalyOrder);

                    if (!skipList.Contains(secondItemPos))
                    {
                        var constructProduct = ConstructProduct(worksheet, startRow, secondItemPos, catId, dispalyOrder,
                            callForPrice);


                        if (constructProduct != null)
                        {
                            result.Add(constructProduct);
                        }
                    }
                    dispalyOrder++;

                    Trace.WriteLine(dispalyOrder);


                    if (!skipList.Contains(therdItemPos))
                    {
                        var constructProduct = ConstructProduct(worksheet, startRow, therdItemPos, catId, dispalyOrder, callForPrice);

                        if (constructProduct != null)
                        {
                            result.Add(constructProduct);
                        }
                    }

                    dispalyOrder++;

                    Trace.WriteLine(dispalyOrder);
                    skipList.Clear();

                    //next 3 product
                    startRow++;
                }
            }

            return result;
        }

        private Picture GetPictureStrict(ExcelWorksheet worksheet, int row, int column, bool isNew)
        {
            var pictureRow = row - 1;
            var pictureColumn = column - 1;


            var pictureOnCoordinates = worksheet.Drawings.Where(p => ((p.To.Column == pictureColumn || p.To.Column == pictureColumn+1) && (p.To.Row == pictureRow)));

            var excelDrawing = pictureOnCoordinates.Count() == 1 ? pictureOnCoordinates.FirstOrDefault() : pictureOnCoordinates.LastOrDefault();
            
            var picture = excelDrawing as ExcelPicture;

            if (picture != null)
            {
                var stream = new MemoryStream();
                picture.Image.Save(stream, picture.ImageFormat);
                byte[] data;
                using (var streamreader = new BinaryReader(stream))
                {
                    stream.Position = 0;

                    data = streamreader.ReadBytes((int) stream.Length);
                }

                var pict = new Picture
                {
                    IsNew = isNew,
                    PictureBinary = data,
                    MimeType = picture.ImageFormat.ToString()
                };
                return pict;
            }

            return null;
        }

        private ExceledProductData ConstructProduct(ExcelWorksheet worksheet, int iRow, int column, int categoryId, int displayOrder,bool callForPrice = false)
        {
            var priceColumn = column + 1;

            var name = worksheet.Cells[iRow, column].Value.ToString();
            var shortDescription = name; // Convert.ToString(worksheet.Cells[iRow, column].Value);
            var fullDescription = name;  //Convert.ToString(worksheet.Cells[iRow, column].Value);

            // var sku = Convert.ToString(worksheet.Cells[iRow, column].Value);

            decimal price = 0;
            try
            {
                price  = Convert.ToDecimal(worksheet.Cells[iRow, priceColumn].Value);
            }
            catch (FormatException)
            {
                var str = worksheet.Cells[iRow, priceColumn].Value.ToString();
                var newStr = (from c in str let isDigit = char.IsDigit(c) where isDigit select c).Aggregate(string.Empty, (current, c) => current + c);

                try
                {
                    price = Convert.ToDecimal(newStr);
                }
                catch (FormatException)
                {
                    return null;
                }
            }


            var product = new Product();
            
            const bool isNew = true;
            

            product.Name = name;
            product.ShortDescription = shortDescription;
            product.FullDescription = fullDescription;

            product.Price = price;


            var utcNow = DateTime.UtcNow;
            product.UpdatedOnUtc = utcNow;
            product.CreatedOnUtc = utcNow;
            product.Published = true;

            
            product.ProductType = ProductType.SimpleProduct;
            product.VisibleIndividually = true;
            product.CallForPrice = callForPrice;

            product.DisplayOrder = displayOrder;
            product.Sku = name;

            var picture = GetPictureStrict(worksheet, iRow, column, isNew);
            
            return new ExceledProductData
            {
                Product = product,
                IsNew = isNew,
                Picture = picture,
                CategoryId = categoryId
            };
        }

        private ProductProcessResult ProceesProduct(ExceledProductData productData, bool isCategoryNotNool)
        {
            var newProduct = productData.IsNew;
            var product = productData.Product;
            product.HasTierPrices = false;
            product.HasDiscountsApplied = false;
           

            if (isCategoryNotNool)
            {
                var productCategory = new ProductCategory
                {
                    ProductId = product.Id,
                    CategoryId = productData.CategoryId,
                };
                product.ProductCategories.Add(productCategory);
            }

            if (productData.Picture == null)
            {
               // TODO Set default pictures.
            }
            else
            {
                var mimeType =productData.Picture.MimeType;
                var newPictureBinary = productData.Picture.PictureBinary;
                var pictureAlreadyExists = false;
                if (!newProduct)
                {
                    ////compare with existing product pictures
                    //var existingPictures = _pictureService.GetPicturesByProductId(product.Id);
                    //foreach (var existingPicture in existingPictures)
                    //{
                    //    var existingBinary = _pictureService.LoadPictureBinary(existingPicture);
                    //    //picture binary after validation (like in database)
                    //    var validatedPictureBinary = _pictureService.ValidatePicture(newPictureBinary, mimeType);
                    //    if (existingBinary.SequenceEqual(validatedPictureBinary))
                    //    {
                    //        //the same picture content
                    //        pictureAlreadyExists = true;
                    //        break;
                    //    }
                    //}
                }

                if (!pictureAlreadyExists)
                {
                    product.ProductPictures.Add(new ProductPicture
                    {
                        Picture = _pictureService.InsertPicture(newPictureBinary,
                        mimeType,
                        _pictureService.GetPictureSeName(product.Name),
                        false),
                        DisplayOrder = 1,
                    });
                }
            }

            var seName = product.ValidateSeName(product.Name, product.Name, true);
            _urlRecordService.SaveSlug(product, seName, 0);

            if (newProduct)
            {
                //_productService.InsertProduct(product);
                return new ProductProcessResult()
                {
                    Product = product,
                    ProductType = NewProductType.ToInsert
                };
            }
            else
            {
               // _productService.UpdateProduct(product);
                return new ProductProcessResult()
                {
                    Product = product,
                    ProductType = NewProductType.ToUpdate
                };
            }
        }

        private void ProcessData(IEnumerable<ExceledProductData> productDatas)
        {
            var firstOrDefault = productDatas.FirstOrDefault();
            if (firstOrDefault != null)
            {
                var category = _categoryService.GetCategoryById(firstOrDefault.CategoryId);

                var siCategoryNotNull = category != null;


                var pruductsToInsert = new List<Product>();
                var pruductsToUpdate = new List<Product>();

                foreach (var exceledProductData in productDatas)
                {
                     var proceesProduct =   ProceesProduct(exceledProductData, siCategoryNotNull);

                    if (proceesProduct.ProductType == NewProductType.ToInsert)
                    {
                        pruductsToInsert.Add(proceesProduct.Product);
                    }
                    if (proceesProduct.ProductType == NewProductType.ToUpdate)
                    {
                        pruductsToUpdate.Add(proceesProduct.Product);
                    }
                }

                if (pruductsToInsert.Count > 0)
                {
                    _productService.InsertProducts(pruductsToInsert.ToArray());
                }


                if (pruductsToUpdate.Count > 0)
                {
                    _productService.UpdateProducts(pruductsToUpdate.ToArray());
                }
            }
        }


        /// <summary>
        /// Updates category image by taking first products image in this category
        /// </summary>
        /// <param name="categoryName">Category name</param>
        private void UpdateCategoryImage(string categoryName)
        {
            var categoryId = _categoryService.GetCategoryByName(categoryName).Id;
            UpdateCategoryImage(categoryId, categoryId);
        }

        /// <summary>
        /// Updates category image by getting image from some of product of specidied category
        /// </summary>
        /// <param name="categoryToToGetProductsFrom">Category to get products from </param>
        /// <param name="categoryIdToChangeImage">Category where image will be chaged</param>
        private void UpdateCategoryImage(int categoryToToGetProductsFrom, int categoryIdToChangeImage)
        {
            var productCategory = _categoryService.GetProductCategoriesByCategoryId(categoryToToGetProductsFrom, 0, 1000000).FirstOrDefault();

            if (productCategory != null)
            {
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
        }

     
    }
}
