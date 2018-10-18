﻿using Com.DanLiris.Service.Purchasing.Lib;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentInternalPurchaseOrderFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentPurchaseRequestFacades;
using Com.DanLiris.Service.Purchasing.Lib.Models.GarmentInternalPurchaseOrderModel;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.GarmentInternalPurchaseOrderViewModel;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.GarmentInternalPurchaseOrderDataUtils;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.GarmentPurchaseRequestDataUtils;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using Moq;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using Xunit;

namespace Com.DanLiris.Service.Purchasing.Test.Facades.GarmentInternalPurchaseOrderTests
{
    public class BasicTest
    {
        private const string ENTITY = "GarmentInternalPurchaseOrder";

        private const string USERNAME = "Unit Test";
        private IServiceProvider ServiceProvider { get; set; }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public string GetCurrentMethod()
        {
            StackTrace st = new StackTrace();
            StackFrame sf = st.GetFrame(1);

            return string.Concat(sf.GetMethod().Name, "_", ENTITY);
        }

        private PurchasingDbContext _dbContext(string testName)
        {
            DbContextOptionsBuilder<PurchasingDbContext> optionsBuilder = new DbContextOptionsBuilder<PurchasingDbContext>();
            optionsBuilder
                .UseInMemoryDatabase(testName)
                .ConfigureWarnings(w => w.Ignore(InMemoryEventId.TransactionIgnoredWarning));

            PurchasingDbContext dbContext = new PurchasingDbContext(optionsBuilder.Options);

            return dbContext;
        }

        private GarmentInternalPurchaseOrderDataUtil dataUtil(GarmentInternalPurchaseOrderFacade facade, string testName)
        {
            var garmentPurchaseRequestFacade = new GarmentPurchaseRequestFacade(_dbContext(testName));
            var garmentPurchaseRequestDataUtil = new GarmentPurchaseRequestDataUtil(garmentPurchaseRequestFacade);

            return new GarmentInternalPurchaseOrderDataUtil(facade, garmentPurchaseRequestDataUtil);
        }

        [Fact]
        public async void Should_Success_Create_Multiple_Data()
        {
            var facade = new GarmentInternalPurchaseOrderFacade(_dbContext(GetCurrentMethod()));
            var listData = dataUtil(facade, GetCurrentMethod()).GetNewData();
            var Response = await facade.CreateMultiple(listData, USERNAME);
            Assert.NotEqual(Response, 0);
        }

        [Fact]
        public async void Should_Error_Create_Multiple_Data()
        {
            var facade = new GarmentInternalPurchaseOrderFacade(_dbContext(GetCurrentMethod()));
            var listData = dataUtil(facade, GetCurrentMethod()).GetNewData();
            foreach (var data in listData)
            {
                data.Items = null;
            }
            Exception e = await Assert.ThrowsAsync<Exception>(async () => await facade.CreateMultiple(listData, USERNAME));
            Assert.NotNull(e.Message);
        }

        [Fact]
        public async void Should_Success_Get_All_Data()
        {
            var facade = new GarmentInternalPurchaseOrderFacade(_dbContext(GetCurrentMethod()));
            var listData = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var Response = facade.Read();
            Assert.NotEqual(Response.Item1.Count, 0);
        }

        [Fact]
        public async void Should_Success_Get_All_Data_With_Items_Order()
        {
            var facade = new GarmentInternalPurchaseOrderFacade(_dbContext(GetCurrentMethod()));
            var listData = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var Response = facade.Read(Order: "{\"Items.ProductName\" : \"desc\"}");
            Assert.NotEqual(Response.Item1.Count, 0);
        }

        [Fact]
        public async void Should_Success_Get_Data_By_Id()
        {
            var facade = new GarmentInternalPurchaseOrderFacade(_dbContext(GetCurrentMethod()));
            var listData = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var Response = facade.ReadById((int) listData.First().Id);
            Assert.NotNull(Response);
        }

        [Fact]
        public async void Should_Success_Check_Cuplicate_Data()
        {
            var facade = new GarmentInternalPurchaseOrderFacade(_dbContext(GetCurrentMethod()));
            var listData = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var Response = facade.CheckDuplicate(listData.First());
            Assert.Equal(Response, false);
        }

        [Fact]
        public async void Should_Success_Split_Data()
        {
            var dbContext = _dbContext(GetCurrentMethod());
            var facade = new GarmentInternalPurchaseOrderFacade(dbContext);
            var listData = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var data = dbContext.GarmentInternalPurchaseOrders.AsNoTracking().Include(m => m.Items).Single(m => m.Id == listData.First().Id);

            var Response = await facade.Split((int)data.Id, data, USERNAME);
            Assert.NotEqual(Response, 0);
        }

        [Fact]
        public async void Should_Error_Split_Data()
        {
            var facade = new GarmentInternalPurchaseOrderFacade(_dbContext(GetCurrentMethod()));
            var listData = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var data = listData.First();
            data.Items = null;
            Exception e = await Assert.ThrowsAsync<Exception>(async () => await facade.Split((int)data.Id, data, USERNAME));
            Assert.NotNull(e.Message);
        }

        [Fact]
        public async void Should_Success_Delete_Data()
        {
            var facade = new GarmentInternalPurchaseOrderFacade(_dbContext(GetCurrentMethod()));
            var listData = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var data = listData.First();
            var Response = await facade.Delete((int)data.Id, USERNAME);
            Assert.NotEqual(Response, 0);
        }

        [Fact]
        public async void Should_Error_Delete_Data()
        {
            var facade = new GarmentInternalPurchaseOrderFacade(_dbContext(GetCurrentMethod()));

            Exception e = await Assert.ThrowsAsync<Exception>(async () => await facade.Delete(0, USERNAME));
            Assert.NotNull(e.Message);
        }

        [Fact]
        public async void Should_Success_Validate_Data()
        {
            var viewModelNullItems = new GarmentInternalPurchaseOrderViewModel
            {
                Items = null
            };
            Assert.True(viewModelNullItems.Validate(null).Count() > 0);

            var viewModelZeroQuantity = new GarmentInternalPurchaseOrderViewModel
            {
                Items = new List<GarmentInternalPurchaseOrderItemViewModel>
                {
                    new GarmentInternalPurchaseOrderItemViewModel
                    {
                        Quantity = 0
                    }
                }
            };
            Assert.True(viewModelZeroQuantity.Validate(null).Count() > 0);

            var facade = new GarmentInternalPurchaseOrderFacade(_dbContext(GetCurrentMethod()));
            var listData = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var data = listData.First();
            var item = data.Items.First();

            Mock<IServiceProvider> serviceProvider = new Mock<IServiceProvider>();
            serviceProvider.
                Setup(x => x.GetService(typeof(PurchasingDbContext)))
                .Returns(_dbContext(GetCurrentMethod()));

            var viewModelDuplicate = new GarmentInternalPurchaseOrderViewModel
            {
                Id = data.Id,
                Items = new List<GarmentInternalPurchaseOrderItemViewModel>
                {
                    new GarmentInternalPurchaseOrderItemViewModel
                    {
                        Id = item.Id,
                        Quantity = item.Quantity
                    }
                }
            };
            ValidationContext validationDuplicateContext = new ValidationContext(viewModelDuplicate, serviceProvider.Object, null);
            Assert.True(viewModelDuplicate.Validate(validationDuplicateContext).Count() > 0);

            var viewModelNotFoundDuplicate = new GarmentInternalPurchaseOrderViewModel
            {
                Id = 1,
                Items = new List<GarmentInternalPurchaseOrderItemViewModel>
                {
                    new GarmentInternalPurchaseOrderItemViewModel
                    {
                        Id = 0,
                        Quantity = 1
                    }
                }
            };
            ValidationContext validationNotFoundDuplicateContext = new ValidationContext(viewModelNotFoundDuplicate, serviceProvider.Object, null);
            Assert.True(viewModelNotFoundDuplicate.Validate(validationNotFoundDuplicateContext).Count() > 0);
        }
    }
}
