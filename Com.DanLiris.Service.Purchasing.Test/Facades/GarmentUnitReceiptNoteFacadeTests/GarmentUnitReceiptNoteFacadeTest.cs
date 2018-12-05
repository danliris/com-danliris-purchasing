﻿using AutoMapper;
using Com.DanLiris.Service.Purchasing.Lib;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentDeliveryOrderFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentExternalPurchaseOrderFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentInternalPurchaseOrderFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentPurchaseRequestFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentUnitReceiptNoteFacades;
using Com.DanLiris.Service.Purchasing.Lib.Interfaces;
using Com.DanLiris.Service.Purchasing.Lib.Models.GarmentUnitReceiptNoteModel;
using Com.DanLiris.Service.Purchasing.Lib.Models.GarmentUnitReceiptNoteViewModels;
using Com.DanLiris.Service.Purchasing.Lib.Services;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.NewIntegrationViewModel;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.GarmentDeliveryOrderDataUtils;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.GarmentExternalPurchaseOrderDataUtils;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.GarmentInternalPurchaseOrderDataUtils;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.GarmentPurchaseRequestDataUtils;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.GarmentUnitReceiptNoteDataUtils;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using Moq;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Text;
using Xunit;

namespace Com.DanLiris.Service.Purchasing.Test.Facades.GarmentUnitReceiptNoteFacadeTests
{
    public class GarmentUnitReceiptNoteFacadeTest
    {
        private const string ENTITY = "GarmentUnitReceiptNote";

        private const string USERNAME = "Unit Test";

        private IServiceProvider GetServiceProvider() {
            HttpResponseMessage httpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK);
            httpResponseMessage.Content = new StringContent("{\"apiVersion\":\"1.0\",\"statusCode\":200,\"message\":\"Ok\",\"data\":[{\"Id\":7,\"code\":\"USD\",\"rate\":13700.0,\"date\":\"2018/10/20\"}],\"info\":{\"count\":1,\"page\":1,\"size\":1,\"total\":2,\"order\":{\"date\":\"desc\"},\"select\":[\"Id\",\"code\",\"rate\",\"date\"]}}");

            var httpClientService = new Mock<IHttpClientService>();
            httpClientService
                .Setup(x => x.GetAsync(It.IsAny<string>()))
                .ReturnsAsync(httpResponseMessage);

            var mapper = new Mock<IMapper>();
            mapper
                .Setup(x => x.Map<GarmentUnitReceiptNoteViewModel>(It.IsAny<GarmentUnitReceiptNote>()))
                .Returns(new GarmentUnitReceiptNoteViewModel {
                    Id = 1,
                    Items = new List<GarmentUnitReceiptNoteItemViewModel>
                    {
                        new GarmentUnitReceiptNoteItemViewModel()
                    }
                });

            var serviceProviderMock = new Mock<IServiceProvider>();
            serviceProviderMock
                .Setup(x => x.GetService(typeof(IdentityService)))
                .Returns(new IdentityService { Username = "Username" });
            serviceProviderMock
                .Setup(x => x.GetService(typeof(IHttpClientService)))
                .Returns(httpClientService.Object);
            serviceProviderMock
                .Setup(x => x.GetService(typeof(IMapper)))
                .Returns(mapper.Object);

            return serviceProviderMock.Object;
        }

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

        private GarmentUnitReceiptNoteDataUtil dataUtil(GarmentUnitReceiptNoteFacade facade, string testName)
        {
            var garmentPurchaseRequestFacade = new GarmentPurchaseRequestFacade(_dbContext(testName));
            var garmentPurchaseRequestDataUtil = new GarmentPurchaseRequestDataUtil(garmentPurchaseRequestFacade);

            var garmentInternalPurchaseOrderFacade = new GarmentInternalPurchaseOrderFacade(_dbContext(testName));
            var garmentInternalPurchaseOrderDataUtil = new GarmentInternalPurchaseOrderDataUtil(garmentInternalPurchaseOrderFacade, garmentPurchaseRequestDataUtil);

            var garmentExternalPurchaseOrderFacade = new GarmentExternalPurchaseOrderFacade(GetServiceProvider(), _dbContext(testName));
            var garmentExternalPurchaseOrderDataUtil = new GarmentExternalPurchaseOrderDataUtil(garmentExternalPurchaseOrderFacade, garmentInternalPurchaseOrderDataUtil);

            var garmentDeliveryOrderFacade = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(testName));
            var garmentDeliveryOrderDataUtil = new GarmentDeliveryOrderDataUtil(garmentDeliveryOrderFacade, garmentExternalPurchaseOrderDataUtil);

            return new GarmentUnitReceiptNoteDataUtil(facade, garmentDeliveryOrderDataUtil);
        }

        [Fact]
        public async void Should_Success_Get_All_Data()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();
            var Response = facade.Read();
            Assert.NotEqual(Response.Data.Count, 0);
        }

        [Fact]
        public async void Should_Success_Get_Data_By_Id()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();
            var Response = facade.ReadById((int)data.Id);
            Assert.NotEqual(Response.Id, 0);
        }

        [Fact]
        public void Should_Success_Generate_Pdf()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var Response = facade.GeneratePdf(new GarmentUnitReceiptNoteViewModel());
            Assert.IsType<MemoryStream>(Response);
        }

        [Fact]
        public async void Should_Success_Create_Data()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = dataUtil(facade, GetCurrentMethod()).GetNewDataWithStorage();
            var Response = await facade.Create(data);
            Assert.NotEqual(Response, 0);
        }

        [Fact]
        public async void Should_Error_Create_Data_Null_Items()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = dataUtil(facade, GetCurrentMethod()).GetNewDataWithStorage();
            data.Items = null;
            Exception e = await Assert.ThrowsAsync<Exception>(async () => await facade.Create(data));
            Assert.NotNull(e.Message);
        }

        private GarmentUnitReceiptNote CopyModel (GarmentUnitReceiptNote data)
        {
            var copiedData = new GarmentUnitReceiptNote
            {
                Active = data.Active,
                CreatedAgent = data.CreatedAgent,
                CreatedBy = data.CreatedBy,
                CreatedUtc = data.CreatedUtc,
                DOId = data.DOId,
                DONo = data.DONo,
                DeletedAgent = data.DeletedAgent,
                DeletedBy = data.DeletedBy,
                DeletedUtc = data.DeletedUtc,
                Id = data.Id,
                IsCorrection = data.IsCorrection,
                IsDeleted = data.IsDeleted,
                IsStorage = data.IsStorage,
                IsUnitDO = data.IsUnitDO,
                LastModifiedAgent = data.LastModifiedAgent,
                LastModifiedBy = data.LastModifiedBy,
                LastModifiedUtc = data.LastModifiedUtc,
                ReceiptDate = data.ReceiptDate,
                Remark = data.Remark,
                StorageCode = data.StorageCode,
                StorageId = data.StorageId,
                StorageName = data.StorageName,
                SupplierCode = data.SupplierCode,
                SupplierId = data.SupplierId,
                SupplierName = data.SupplierName,
                UId = data.UId,
                URNNo = data.URNNo,
                UnitCode = data.UnitCode,
                UnitId = data.UnitId,
                UnitName = data.UnitName,
                Items = data.Items.Select(i => CopyModelItem(i)).ToList()
            };

            return copiedData;
        }

        private GarmentUnitReceiptNoteItem CopyModelItem(GarmentUnitReceiptNoteItem i)
        {
            var copiedItem = new GarmentUnitReceiptNoteItem
            {
                Active = i.Active,
                Conversion = i.Conversion,
                CreatedAgent = i.CreatedAgent,
                CreatedBy = i.CreatedBy,
                CreatedUtc = i.CreatedUtc,
                DODetailId = i.DODetailId,
                DeletedAgent = i.DeletedAgent,
                DeletedBy = i.DeletedBy,
                DeletedUtc = i.DeletedUtc,
                DesignColor = i.DesignColor,
                EPOItemId = i.EPOItemId,
                GarmentUnitReceiptNote = i.GarmentUnitReceiptNote,
                Id = i.Id,
                IsCorrection = i.IsCorrection,
                IsDeleted = i.IsDeleted,
                LastModifiedAgent = i.LastModifiedAgent,
                LastModifiedBy = i.LastModifiedBy,
                LastModifiedUtc = i.LastModifiedUtc,
                POId = i.POId,
                POItemId = i.POItemId,
                POSerialNumber = i.POSerialNumber,
                PRId = i.PRId,
                PRItemId = i.PRItemId,
                PRNo = i.PRNo,
                PricePerDealUnit = i.PricePerDealUnit,
                ProductCode = i.ProductCode,
                ProductId = i.ProductId,
                ProductName = i.ProductName,
                ProductRemark = i.ProductRemark,
                RONo = i.RONo,
                ReceiptQuantity = i.ReceiptQuantity,
                SmallQuantity = i.SmallQuantity,
                SmallUomId = i.SmallUomId,
                SmallUomUnit = i.SmallUomUnit,
                URNId = i.URNId,
                UomId = i.UomId,
                UomUnit = i.UomUnit
            };

            return copiedItem;
        }

        [Fact]
        public async void Should_Success_Update_Data()
        {
            var dbContext = _dbContext(GetCurrentMethod());
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), dbContext);

            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();

            var ResponseUpdate = await facade.Update((int)data.Id, data);
            Assert.NotEqual(ResponseUpdate, 0);

            ////var itemAdded = CopyModelItem(data.Items.First());
            ////itemAdded.Id = 0;
            //data.Items.Clear();
            ////data.Items.Add(itemAdded);
            //var ResponseUpdateAddItem = await facade.Update((int)data.Id, data);
            //Assert.NotEqual(ResponseUpdateAddItem, 0);

            //data.Items.Remove(itemAdded);
            //var ResponseUpdateRemoveItem = await facade.Update((int)data.Id, data);
            //Assert.NotEqual(ResponseUpdateRemoveItem, 0);
        }

        [Fact]
        public async void Should_Error_Update_Data_Null_Items()
        {
            var dbContext = _dbContext(GetCurrentMethod());
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), dbContext);

            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();
            dbContext.Entry(data).State = EntityState.Detached;
            data.Items = null;

            Exception e = await Assert.ThrowsAsync<Exception>(async () => await facade.Update((int)data.Id, data));
            Assert.NotNull(e.Message);
        }

        [Fact]
        public async void Should_Success_Delete_Data()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();

            var Response = await facade.Delete((int)data.Id);
            Assert.NotEqual(Response, 0);
        }

        [Fact]
        public async void Should_Error_Delete_Data_Invalid_Id()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();

            Exception e = await Assert.ThrowsAsync<Exception>(async () => await facade.Delete(0));
            Assert.NotNull(e.Message);
        }

        [Fact]
        public void Should_Success_Validate_Data()
        {
            GarmentUnitReceiptNoteViewModel viewModel = new GarmentUnitReceiptNoteViewModel { IsStorage = true };
            Assert.True(viewModel.Validate(null).Count() > 0);

            GarmentUnitReceiptNoteViewModel viewModelCheckDeliveryOrder = new GarmentUnitReceiptNoteViewModel {
                Supplier = new SupplierViewModel { Id = 1 },
                Unit = new UnitViewModel { Id = "1" },
            };
            Assert.True(viewModelCheckDeliveryOrder.Validate(null).Count() > 0);

            GarmentUnitReceiptNoteViewModel viewModelCheckItemsCount = new GarmentUnitReceiptNoteViewModel { DOId = 1 };
            Assert.True(viewModelCheckItemsCount.Validate(null).Count() > 0);

            GarmentUnitReceiptNoteViewModel viewModelCheckItems = new GarmentUnitReceiptNoteViewModel {
                DOId = 1,
                Items = new List<GarmentUnitReceiptNoteItemViewModel>
                {
                    new GarmentUnitReceiptNoteItemViewModel()
                }
            };
            Assert.True(viewModelCheckItems.Validate(null).Count() > 0);
        }
    }
}
