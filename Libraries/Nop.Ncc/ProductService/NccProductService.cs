using System;
using System.Collections.Generic;
using Nop.Core;
using Nop.Core.Caching;
using Nop.Core.Data;
using Nop.Core.Domain.Catalog;
using Nop.Core.Domain.Common;
using Nop.Core.Domain.Localization;
using Nop.Core.Domain.Security;
using Nop.Core.Domain.Stores;
using Nop.Data;
using Nop.Ncc.Repository;
using Nop.Services.Catalog;
using Nop.Services.Events;
using Nop.Services.Localization;
using Nop.Services.Messages;
using Nop.Services.Security;
using Nop.Services.Stores;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace Nop.Ncc
{
    public class NccProductService : ProductService, INccProductService
    {
        private const string PRODUCTS_PATTERN_KEY = "Nop.product.";

        protected readonly IRepository<Product> _productRepository;
        protected readonly ICacheManager _cacheManager;

        public NccProductService(ICacheManager cacheManager,
            IRepository<Product> productRepository,
            IRepository<RelatedProduct> relatedProductRepository,
            IRepository<CrossSellProduct> crossSellProductRepository,
            IRepository<TierPrice> tierPriceRepository,
            IRepository<ProductPicture> productPictureRepository,
            IRepository<LocalizedProperty> localizedPropertyRepository,
            IRepository<AclRecord> aclRepository,
            IRepository<StoreMapping> storeMappingRepository,
            IRepository<ProductSpecificationAttribute> productSpecificationAttributeRepository,
            IRepository<ProductReview> productReviewRepository,
            IRepository<ProductWarehouseInventory> productWarehouseInventoryRepository,
            IProductAttributeService productAttributeService,
            IProductAttributeParser productAttributeParser,
            ILanguageService languageService,
            IWorkflowMessageService workflowMessageService,
            IDataProvider dataProvider,
            Nop.Data.IDbContext dbContext,
            IWorkContext workContext,
            IStoreContext storeContext,
            LocalizationSettings localizationSettings,
            CommonSettings commonSettings,
            CatalogSettings catalogSettings,
            IEventPublisher eventPublisher,
            IAclService aclService,
            IStoreMappingService storeMappingService)
            : base(cacheManager,
                  productRepository,
                  relatedProductRepository,
                  crossSellProductRepository,
                  tierPriceRepository,
                  productPictureRepository,
                  localizedPropertyRepository,
                  aclRepository,
                  storeMappingRepository,
                  productSpecificationAttributeRepository,
                  productReviewRepository,
                  productWarehouseInventoryRepository,
                  productAttributeService,
                  productAttributeParser,
                  languageService,
                  workflowMessageService,
                  dataProvider,
                  dbContext,
                  workContext,
                  storeContext,
                  localizationSettings,
                  commonSettings,
                  catalogSettings,
                  eventPublisher,
                  aclService,
                  storeMappingService)
        {
            _productRepository = productRepository;
            _cacheManager = cacheManager;
        }


        public void DeleteProducts(Product[] products)
        {
            if (products == null)
                throw new ArgumentNullException("products");

            foreach (var product in products)
            {
                product.Deleted = true;
            }

            //delete product
            UpdateProducts(products);
        }

        public void InsertProducts(Product[] products)
        {
            //throw new NotImplementedException();

            if (products == null)
                throw new ArgumentNullException("products");

            //insert
            _productRepository.Insert(products);

            //clear cache
            _cacheManager.RemoveByPattern(PRODUCTS_PATTERN_KEY);

            //event notification
            //  _eventPublisher.EntityInserted(product);
            // TODO: Ncc need invistigate
        }

        public void UpdateProducts(Product[] products)
        {
            //throw new NotImplementedException();

            if (products == null)
                throw new ArgumentNullException("products");

            //update
            //_productRepository.Update(products);

            ((EfRepository<Product>) _productRepository).Update(products);



            //cache
            _cacheManager.RemoveByPattern(PRODUCTS_PATTERN_KEY);

            //event notification
            //_eventPublisher.EntityUpdated(product);
        }

        public void Update(IEnumerable<T> entities)
        {
            throw new NotImplementedException();
        }
    }
}