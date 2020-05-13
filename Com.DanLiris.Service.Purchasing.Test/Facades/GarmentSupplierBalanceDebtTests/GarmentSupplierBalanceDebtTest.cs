﻿using Com.DanLiris.Service.Purchasing.Lib;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentDeliveryOrderFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentExternalPurchaseOrderFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentInternalPurchaseOrderFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentPurchaseRequestFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentSupplierBalanceDebtFacades;
using Com.DanLiris.Service.Purchasing.Lib.Interfaces;
using Com.DanLiris.Service.Purchasing.Lib.Models.GarmentSupplierBalanceDebtModel;
using Com.DanLiris.Service.Purchasing.Lib.Services;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.GarmentSupplierBalanceDebtViewModel;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.GarmentDeliveryOrderDataUtils;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.GarmentExternalPurchaseOrderDataUtils;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.GarmentInternalPurchaseOrderDataUtils;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.GarmentPurchaseRequestDataUtils;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using Moq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Com.DanLiris.Service.Purchasing.Test.Facades.GarmentSupplierBalanceDebtTests
{
    public class GarmentSupplierBalanceDebtTest
    {
        private const string ENTITY = "GarmentSupplierBalanceDebt";

        private const string USERNAME = "Unit Test";
        private IServiceProvider ServiceProvider { get; set; }

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

        private Mock<IServiceProvider> GetServiceProvider()
        {
            HttpResponseMessage message = new HttpResponseMessage(System.Net.HttpStatusCode.OK);
            message.Content = new StringContent("{\"apiVersion\":\"1.0\",\"statusCode\":200,\"message\":\"Ok\",\"data\":[{\"Id\":7,\"code\":\"USD\",\"rate\":13700.0,\"date\":\"2018/10/20\"}],\"info\":{\"count\":1,\"page\":1,\"size\":1,\"total\":2,\"order\":{\"date\":\"desc\"},\"select\":[\"Id\",\"code\",\"rate\",\"date\"]}}");
            var HttpClientService = new Mock<IHttpClientService>();
            HttpClientService
                .Setup(x => x.GetAsync(It.IsAny<string>()))
                .ReturnsAsync(message);

            var serviceProvider = new Mock<IServiceProvider>();
            serviceProvider
                .Setup(x => x.GetService(typeof(IdentityService)))
                .Returns(new IdentityService() { Token = "Token", Username = "Test" });

            serviceProvider
                .Setup(x => x.GetService(typeof(IHttpClientService)))
                .Returns(HttpClientService.Object);

            return serviceProvider;
        }

        private GarmentDeliveryOrderDataUtil dataUtil(GarmentDeliveryOrderFacade facade, string testName)
        {
            var garmentPurchaseRequestFacade = new GarmentPurchaseRequestFacade(ServiceProvider, _dbContext(testName));
            var garmentPurchaseRequestDataUtil = new GarmentPurchaseRequestDataUtil(garmentPurchaseRequestFacade);

            var garmentInternalPurchaseOrderFacade = new GarmentInternalPurchaseOrderFacade(_dbContext(testName));
            var garmentInternalPurchaseOrderDataUtil = new GarmentInternalPurchaseOrderDataUtil(garmentInternalPurchaseOrderFacade, garmentPurchaseRequestDataUtil);

            var garmentExternalPurchaseOrderFacade = new GarmentExternalPurchaseOrderFacade(ServiceProvider, _dbContext(testName));
            var garmentExternalPurchaseOrderDataUtil = new GarmentExternalPurchaseOrderDataUtil(garmentExternalPurchaseOrderFacade, garmentInternalPurchaseOrderDataUtil);

            return new GarmentDeliveryOrderDataUtil(facade, garmentExternalPurchaseOrderDataUtil);
        }

        [Fact]
        public async Task Should_Success_Create_Data()
        {
            var facade = new GarmentSupplierBalanceDebtFacade(ServiceProvider, _dbContext(GetCurrentMethod()));
            GarmentSupplierBalanceDebt data = new GarmentSupplierBalanceDebt
            {
                SupplierCode = $"BuyerCode{DateTimeOffset.Now.Ticks}a",
                DOCurrencyCode = "IDR",
                DOCurrencyRate = 1,
                CodeRequirment = $"{DateTimeOffset.Now.Ticks}a",
                CreatedBy = "UnitTest",
                Import = false,
                SupplierName = "SupplierTest123",
                TotalValas = DateTimeOffset.Now.Ticks,
                TotalAmountIDR = DateTimeOffset.Now.Ticks,
                DOCurrencyId = 1,
                SupplierId = 1,
                Year = 2020,
                Items = new List<GarmentSupplierBalanceDebtItem> {
                    new GarmentSupplierBalanceDebtItem{
                        BillNo = "BP181122142947000001",
                        ArrivalDate = DateTimeOffset.Now,
                        DONo = $"{DateTimeOffset.Now.Ticks}a",
                        DOId = 1,
                        InternNo = "InternNO1234",
                        IDR = DateTimeOffset.Now.Ticks,
                        Valas =DateTimeOffset.Now.Ticks

                    }
                }
            };
            var Responses = await facade.Create(data, USERNAME);
            Assert.NotEqual(0, Responses);
        }
        [Fact]
        public async Task Should_Error_Create_Data()
        {
            var facade = new GarmentSupplierBalanceDebtFacade(ServiceProvider, _dbContext(GetCurrentMethod()));
            GarmentSupplierBalanceDebt data = new GarmentSupplierBalanceDebt
            {
                SupplierCode = $"BuyerCode{DateTimeOffset.Now.Ticks}a",
                DOCurrencyCode = "IDR",
                DOCurrencyRate = 1,
                CodeRequirment = $"{DateTimeOffset.Now.Ticks}a",
                CreatedBy = "UnitTest",
                Import = false,
                SupplierName = "SupplierTest123",
                TotalValas = DateTimeOffset.Now.Ticks,
                TotalAmountIDR = DateTimeOffset.Now.Ticks,
                DOCurrencyId = 1,
                SupplierId = 1,
                Year = 2020,
                Items = new List<GarmentSupplierBalanceDebtItem> {
                    new GarmentSupplierBalanceDebtItem{
                        BillNo = "BP181122142947000001",
                        ArrivalDate = DateTimeOffset.Now,
                        DONo = $"{DateTimeOffset.Now.Ticks}a",
                        DOId = 1,
                        InternNo = "InternNO1234",
                        IDR = DateTimeOffset.Now.Ticks,
                        Valas =DateTimeOffset.Now.Ticks

                    }
                }
            };
            Exception e = await Assert.ThrowsAsync<Exception>(async () => await facade.Create(null, USERNAME));
            Assert.NotNull(e.Message);
        }
        [Fact]
        public async Task Should_Success_Get_Data()
        {
            var facade = new GarmentSupplierBalanceDebtFacade(ServiceProvider, _dbContext(GetCurrentMethod()));
            GarmentSupplierBalanceDebt data = new GarmentSupplierBalanceDebt {
                SupplierCode = $"BuyerCode{DateTimeOffset.Now.Ticks}a",
                DOCurrencyCode = "IDR",
                DOCurrencyRate = 1,
                CodeRequirment = $"{DateTimeOffset.Now.Ticks}a",
                CreatedBy = "UnitTest",
                Import = false,
                SupplierName = "SupplierTest123",
                TotalValas = DateTimeOffset.Now.Ticks,
                TotalAmountIDR = DateTimeOffset.Now.Ticks,
                DOCurrencyId = 1,
                SupplierId = 1,
                Year = 2020,
                Items = new List<GarmentSupplierBalanceDebtItem> {
                    new GarmentSupplierBalanceDebtItem{
                        BillNo = "BP181122142947000001",
                        ArrivalDate = DateTimeOffset.Now,
                        DONo = $"{DateTimeOffset.Now.Ticks}a",
                        DOId = 1,
                        InternNo = "InternNO1234",
                        IDR = DateTimeOffset.Now.Ticks,
                        Valas =DateTimeOffset.Now.Ticks

                    }
                }
            };
            var Responses = await facade.Create(data, USERNAME);
            var Response = facade.Read(1, 25, "{}", null, "{}");
            Assert.NotNull(Response.Item1);
        }
        [Fact]
        public async Task Should_Success_Get_Data_By_Id()
        {
            var facade = new GarmentSupplierBalanceDebtFacade(ServiceProvider, _dbContext(GetCurrentMethod()));
            GarmentSupplierBalanceDebt data = new GarmentSupplierBalanceDebt
            {
                SupplierCode = $"BuyerCode{DateTimeOffset.Now.Ticks}a",
                DOCurrencyCode = "IDR",
                DOCurrencyRate = 1,
                CodeRequirment = $"{DateTimeOffset.Now.Ticks}a",
                CreatedBy = "UnitTest",
                Import = false,
                SupplierName = "SupplierTest123",
                TotalValas = DateTimeOffset.Now.Ticks,
                TotalAmountIDR = DateTimeOffset.Now.Ticks,
                DOCurrencyId = 1,
                SupplierId = 1,
                Year = 2020,
                Items = new List<GarmentSupplierBalanceDebtItem> {
                    new GarmentSupplierBalanceDebtItem{
                        BillNo = "BP181122142947000001",
                        ArrivalDate = DateTimeOffset.Now,
                        DONo = $"{DateTimeOffset.Now.Ticks}a",
                        DOId = 1,
                        InternNo = "InternNO1234",
                        IDR = DateTimeOffset.Now.Ticks,
                        Valas =DateTimeOffset.Now.Ticks

                    }
                }
            };
            var Responses = await facade.Create(data, USERNAME);
            var Response = facade.Read((int)data.Id);
            Assert.NotNull(Response.Item1);
        }
        [Fact]
        public async Task Should_Success_Get_Loader()
        {
            GarmentDeliveryOrderFacade facade = new GarmentDeliveryOrderFacade(GetServiceProvider().Object, _dbContext(GetCurrentMethod()));
            var model = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var facadedebt = new GarmentSupplierBalanceDebtFacade(ServiceProvider, _dbContext(GetCurrentMethod()));
            var Response = facadedebt.ReadLoader(1,25,"{}",DateTime.Now.Year);
            Assert.NotEmpty(Response.Data);
        }

        [Fact]
        public async Task Should_Success_Get_Loader_With_Params()
        {
            GarmentDeliveryOrderFacade facade = new GarmentDeliveryOrderFacade(GetServiceProvider().Object, _dbContext(GetCurrentMethod()));
            var model = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var facadedebt = new GarmentSupplierBalanceDebtFacade(ServiceProvider, _dbContext(GetCurrentMethod()));
            var Response = facadedebt.ReadLoader(Search: "[\"BillNo\"]", Select: "{ \"billNo\" : \"BillNo\", \"dONo\" : \"DONo\", \"ArrivalDate\" : 1 }", year: DateTime.Now.Year);
            Assert.NotEmpty(Response.Data);
        }

        //[Fact]
        //public void Should_Fail_Upload_Validate_Data()
        //{
        //    var facade = new GarmentSupplierBalanceDebtFacade(ServiceProvider, _dbContext(GetCurrentMethod()));
        //    GarmentSupplierBalanceDebtViewModel data = new GarmentSupplierBalanceDebtViewModel
        //    {
        //        SupplierCode = "",
        //        CurrencyCode = "",
        //        CodeRequirment = "",
        //        CreatedBy = "",
        //        Import = null,
        //        SupplierName = "",
        //        TotalAmount = null,
        //        TotalAmountIDR = null,
        //        Month = null,
        //        Year = null,

        //    };
        //    //viewModel.Code = "11.1.1";
        //    List<GarmentSupplierBalanceDebtViewModel> coa = new List<GarmentSupplierBalanceDebtViewModel>() { data };
        //    var Response = facade.UploadValidate(ref coa, null);
        //    Assert.False(Response.Item1);
        //}
        //[Fact]
        //public void Should_Fail_Upload_Validate_Month_and_Import()
        //{
        //    var facade = new GarmentSupplierBalanceDebtFacade(ServiceProvider, _dbContext(GetCurrentMethod()));
        //    GarmentSupplierBalanceDebtViewModel data = new GarmentSupplierBalanceDebtViewModel
        //    {
        //        SupplierCode = "",
        //        CurrencyCode = "",
        //        CodeRequirment = "",
        //        CreatedBy = "",
        //        Import = "YA",
        //        SupplierName = "",
        //        TotalAmount = null,
        //        TotalAmountIDR = null,
        //        Month = 13,
        //        Year = null,

        //    };
        //    //viewModel.Code = "11.1.1";
        //    List<GarmentSupplierBalanceDebtViewModel> coa = new List<GarmentSupplierBalanceDebtViewModel>() { data };
        //    var Response = facade.UploadValidate(ref coa, null);
        //    Assert.False(Response.Item1);
        //}

    }
}
