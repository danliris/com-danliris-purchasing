﻿using AutoMapper;
using Com.DanLiris.Service.Purchasing.Lib;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentDeliveryOrderFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentExternalPurchaseOrderFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentInternalPurchaseOrderFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentPurchaseRequestFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentUnitReceiptNoteFacades;
using Com.DanLiris.Service.Purchasing.Lib.Interfaces;
using Com.DanLiris.Service.Purchasing.Lib.Models.GarmentUnitReceiptNoteModel;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.GarmentUnitReceiptNoteViewModels;
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
using Com.DanLiris.Service.Purchasing.Lib.Facades.MonitoringUnitReceiptFacades;
using Com.DanLiris.Service.Purchasing.Lib.Models.GarmentDeliveryOrderModel;
using System.Threading.Tasks;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentBeacukaiFacade;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.GarmentBeacukaiDataUtils;
using Com.DanLiris.Service.Purchasing.Lib.Facades.MonitoringCentralBillReceptionFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.MonitoringCentralBillExpenditureFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentCorrectionNoteFacades;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.GarmentCorrectionNoteDataUtils;
using Com.DanLiris.Service.Purchasing.Lib.Facades.MonitoringCorrectionNoteReceptionFacades;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.NewIntegrationDataUtils;
using Com.DanLiris.Service.Purchasing.Lib.Facades.MonitoringCorrectionNoteExpenditureFacades;
using Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentDailyPurchasingReportFacade;

namespace Com.DanLiris.Service.Purchasing.Test.Facades.GarmentUnitReceiptNoteFacadeTests
{
    public class GarmentUnitReceiptNoteFacadeTest
    {
        private const string ENTITY = "GarmentUnitReceiptNote";

        private const string USERNAME = "Unit Test";
		private IServiceProvider ServiceProvider { get; set; }

		private IServiceProvider GetServiceProvider() {
            HttpResponseMessage httpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK);
            httpResponseMessage.Content = new StringContent("{\"apiVersion\":\"1.0\",\"statusCode\":200,\"message\":\"Ok\",\"data\":[{\"Id\":7,\"code\":\"USD\",\"rate\":13700.0,\"date\":\"2018/10/20\"}],\"info\":{\"count\":1,\"page\":1,\"size\":1,\"total\":2,\"order\":{\"date\":\"desc\"},\"select\":[\"Id\",\"code\",\"rate\",\"date\"]}}");

            var httpClientService = new Mock<IHttpClientService>();
            httpClientService
                .Setup(x => x.GetAsync(It.IsAny<string>()))
                .ReturnsAsync(httpResponseMessage);

            httpClientService
               .Setup(x => x.GetAsync(It.Is<string>(s => s.Contains("master/garment-suppliers"))))
               .ReturnsAsync(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(new SupplierDataUtil().GetResultFormatterOkString()) });

            httpClientService
               .Setup(x => x.GetAsync(It.Is<string>(s => s.Contains("delivery-returns"))))
               .ReturnsAsync(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(new GarmentDeliveryReturnDataUtil().GetResultFormatterOkString()) });

            httpClientService
               .Setup(x => x.PutAsync(It.Is<string>(s => s.Contains("delivery-returns")), It.IsAny<HttpContent>()))
               .ReturnsAsync(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(new GarmentDeliveryReturnDataUtil().GetResultFormatterOkString()) });


            var mapper = new Mock<IMapper>();
            mapper
                .Setup(x => x.Map<GarmentUnitReceiptNoteViewModel>(It.IsAny<GarmentUnitReceiptNote>()))
                .Returns(new GarmentUnitReceiptNoteViewModel {
                    Id = 1,
                    DOId = 1,
                    DOCurrency = new CurrencyViewModel(),
                    Supplier = new SupplierViewModel(),
                    Unit = new UnitViewModel(),
                    Items = new List<GarmentUnitReceiptNoteItemViewModel>
                    {
                        new GarmentUnitReceiptNoteItemViewModel {
                            Product = new GarmentProductViewModel(),
                            Uom = new UomViewModel()
                        }
                    }
                });

            var mockGarmentDeliveryOrderFacade = new Mock<IGarmentDeliveryOrderFacade>();
            mockGarmentDeliveryOrderFacade
                .Setup(x => x.ReadById(It.IsAny<int>()))
                .Returns(new GarmentDeliveryOrder());

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
            serviceProviderMock
                .Setup(x => x.GetService(typeof(IGarmentDeliveryOrderFacade)))
                .Returns(mockGarmentDeliveryOrderFacade.Object);

            serviceProviderMock
                .Setup(x => x.GetService(typeof(IdentityService)))
                .Returns(new IdentityService() { Token = "Token", Username = "Test" });

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
            var garmentPurchaseRequestFacade = new GarmentPurchaseRequestFacade(GetServiceProvider(), _dbContext(testName));
            var garmentPurchaseRequestDataUtil = new GarmentPurchaseRequestDataUtil(garmentPurchaseRequestFacade);

            var garmentInternalPurchaseOrderFacade = new GarmentInternalPurchaseOrderFacade(_dbContext(testName));
            var garmentInternalPurchaseOrderDataUtil = new GarmentInternalPurchaseOrderDataUtil(garmentInternalPurchaseOrderFacade, garmentPurchaseRequestDataUtil);

            var garmentExternalPurchaseOrderFacade = new GarmentExternalPurchaseOrderFacade(GetServiceProvider(), _dbContext(testName));
            var garmentExternalPurchaseOrderDataUtil = new GarmentExternalPurchaseOrderDataUtil(garmentExternalPurchaseOrderFacade, garmentInternalPurchaseOrderDataUtil);

            var garmentDeliveryOrderFacade = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(testName));
            var garmentDeliveryOrderDataUtil = new GarmentDeliveryOrderDataUtil(garmentDeliveryOrderFacade, garmentExternalPurchaseOrderDataUtil);

            return new GarmentUnitReceiptNoteDataUtil(facade, garmentDeliveryOrderDataUtil);
        }

		private GarmentDeliveryOrderDataUtil dataUtilDO(GarmentDeliveryOrderFacade facade, string testName)
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
        public async Task Should_Success_Get_All_Data()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();
            var Response = facade.Read();
            Assert.NotEqual(Response.Data.Count, 0);
        }

        [Fact]
        public async Task Should_Success_Get_Data_By_Id()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();
            var Response = facade.ReadById((int)data.Id);
            Assert.NotEqual(Response.Id, 0);
        }

        [Fact]
        public async Task Should_Success_Generate_Pdf()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();
            var dataViewModel = facade.ReadById((int)data.Id);
            var temp = dataViewModel.DOCurrency.Rate + 1;
            var Response = facade.GeneratePdf(dataViewModel);
            Assert.IsType<MemoryStream>(Response);
        }

        [Fact]
        public async Task Should_Success_Create_Data()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetNewDataWithStorage();
            var Response = await facade.Create(data);
            Assert.NotEqual(Response, 0);

            //var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data1 = await dataUtil(facade, GetCurrentMethod()).GetNewDataWithStorage();
            data1.StorageId = data.StorageId;
            data1.Items.First().UomId = data.Items.First().UomId;
            data1.UnitId = data.UnitId;
            var Response1 = await facade.Create(data1);
            Assert.NotEqual(Response1, 0);
        }

        [Fact]
        public async Task Should_Success_Create_Data_PROSES()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetNewDataWithStorage();
            data.URNType = "PROSES";
            var Response = await facade.Create(data);
            Assert.NotEqual(Response, 0);

        }

        [Fact]
        public async Task Should_Error_Create_Data_Null_Items()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetNewDataWithStorage();
            data.Items = null;
            Exception e = await Assert.ThrowsAsync<Exception>(async () => await facade.Create(data));
            Assert.NotNull(e.Message);
        }

        [Fact]
        public async Task Should_Success_Update_Data()
        {
            var dbContext = _dbContext(GetCurrentMethod());
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), dbContext);

            var dataUtil = this.dataUtil(facade, GetCurrentMethod());
            var data = await dataUtil.GetTestDataWithStorage();
            dbContext.Entry(data).State = EntityState.Detached;
            foreach (var item in data.Items)
            {
                dbContext.Entry(item).State = EntityState.Detached;
            }

            dataUtil.SetDataWithStorage(data);
            var ResponseUpdateStorage = await facade.Update((int)data.Id, data);
            Assert.NotEqual(ResponseUpdateStorage, 0);

            //// Create Storage based on UnitId that contain longTick on create DataUtil
            //dataUtil.SetDataWithStorage(data, data.UnitId);
            //var ResponseRestoreStorage = await facade.Update((int)data.Id, data);
            //Assert.NotEqual(ResponseRestoreStorage, 0);

            //data.IsStorage = false;
            //var ResponseDeleteStorage = await facade.Update((int)data.Id, data);
            //Assert.NotEqual(ResponseDeleteStorage, 0);

            //dataUtil.SetDataWithStorage(data);
            //var ResponseAddStorage = await facade.Update((int)data.Id, data);
            //Assert.NotEqual(ResponseAddStorage, 0);
        }

        [Fact]
        public async Task Should_Error_Update_Data()
        {
            var dbContext = _dbContext(GetCurrentMethod());
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), dbContext);

            Exception e = await Assert.ThrowsAsync<Exception>(async () => await facade.Update(0, new GarmentUnitReceiptNote()));
            Assert.NotNull(e.Message);
        }

        [Fact]
        public async Task Should_Success_Delete_Data()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();

            var Response = await facade.Delete((int)data.Id, (string)data.DeletedReason);
            Assert.NotEqual(Response, 0);
        }

        [Fact]
        public async Task Should_Success_Delete_Data_Proses()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetNewDataWithStorage();
            data.URNType = "PROSES";
            await facade.Create(data);

            var Response = await facade.Delete((int)data.Id, (string)data.DeletedReason);
            Assert.NotEqual(Response, 0);
        }

        [Fact]
        public async Task Should_Success_Delete_Data2()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage2();

            var Response = await facade.Delete((int)data.Id, (string)data.DeletedReason);
            Assert.NotEqual(Response, 0);
        }

        [Fact]
        public async Task Should_Error_Delete_Data_Invalid_Id()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();

            Exception e = await Assert.ThrowsAsync<Exception>(async () => await facade.Delete(0,""));
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
                URNType="PEMBELIAN"
            };
            Assert.True(viewModelCheckDeliveryOrder.Validate(null).Count() > 0);

            GarmentUnitReceiptNoteViewModel viewModelCheckItemsCount = new GarmentUnitReceiptNoteViewModel { DOId = 1 };
            Assert.True(viewModelCheckItemsCount.Validate(null).Count() > 0);

            GarmentUnitReceiptNoteViewModel viewModelCheckItems = new GarmentUnitReceiptNoteViewModel {
                DOId = 1,
                URNType = "PEMBELIAN",
                Items = new List<GarmentUnitReceiptNoteItemViewModel>
                {
                    new GarmentUnitReceiptNoteItemViewModel()
                }
            };
            Assert.True(viewModelCheckItems.Validate(null).Count() > 0);

            GarmentUnitReceiptNoteViewModel viewModelCheckItemsConvertion = new GarmentUnitReceiptNoteViewModel
            {
                DOId = 1,
                URNType = "PEMBELIAN",
                Items = new List<GarmentUnitReceiptNoteItemViewModel>
                {
                    new GarmentUnitReceiptNoteItemViewModel
                    {
                        Uom = new UomViewModel
                        {
                            Id = "1"
                        },
                        SmallUom = new UomViewModel
                        {
                            Id = "1"
                        },
                        Conversion = 10,
                        CorrectionConversion= 10
                    }
                }
            };
            Assert.True(viewModelCheckItemsConvertion.Validate(null).Count() > 0);
        }
		//monitoring
		[Fact]
		public async Task Should_Success_Get_Report_Data()
		{
			GarmentDeliveryOrderFacade facade = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
			var data = await dataUtilDO(facade, GetCurrentMethod()).GetNewData();
			await facade.Create(data, USERNAME);
			var uFacade= new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
			var datas = await dataUtil(uFacade, GetCurrentMethod()).GetNewDataWithStorage();
			var Responses = await uFacade.Create(datas);
			Assert.NotEqual(Responses, 0);
			var Facade = new MonitoringUnitReceiptAllFacade(ServiceProvider, _dbContext(GetCurrentMethod()));
			var Response = Facade.GetReport(null, null, null, null, null, null, null, null, 1, 25, "{}", 7);
			Assert.NotEqual(Response.Item2, 0);

			var Response1 = Facade.GetReport(null, null, null, null, null, null, null, null, 1, 25, "{}", 7);
			Assert.NotNull(Response1.Item1);


		}

		[Fact]
		public async Task Should_Success_Get_Report_Excel()
		{
			GarmentDeliveryOrderFacade facade = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
			var data = await dataUtilDO(facade, GetCurrentMethod()).GetNewData();
			await facade.Create(data, USERNAME);
			var uFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
			var datas = await dataUtil(uFacade, GetCurrentMethod()).GetNewDataWithStorage();
			var Responses = await uFacade.Create(datas);
			Assert.NotEqual(Responses, 0);
			var Facade = new MonitoringUnitReceiptAllFacade(ServiceProvider, _dbContext(GetCurrentMethod()));
			var Response = Facade.GetReport(null, null, null, null, null, null, null, null, 1, 25, "{}", 7);
			Assert.NotEqual(Response.Item2, 0);

			var Response1 = Facade.GenerateExcel(null, null, null, null, null, null, null, null, 1, 25, "{}", 7);
			Assert.IsType(typeof(System.IO.MemoryStream), Response1);
		}

        [Fact]
        public async Task Should_Success_ReadForUnitDO()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();
            var Response = facade.ReadForUnitDO();
            Assert.NotEqual(Response.Count, 0);
        }

        [Fact]
        public async Task Should_Success_ReadForUnitDO_With_Filter()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();
            var filter = new
            {
                data.UnitId,
                data.StorageId,
            };
            var Response = facade.ReadForUnitDO("", JsonConvert.SerializeObject(filter));
            Assert.NotEqual(Response.Count, 0);
        }

        [Fact]
        public async Task Should_Success_ReadForUnitDOHeader()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();
            var Response = facade.ReadForUnitDOHeader();
            Assert.NotEqual(Response.Count, 0);
        }

        [Fact]
        public async Task Should_Success_ReadForUnitDOHeader_With_Filter()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();
            var filter = new
            {
                data.UnitId,
                data.StorageId,
                RONo = "xxx"
            };
            var Response = facade.ReadForUnitDOHeader("", JsonConvert.SerializeObject(filter));
            Assert.NotEqual(Response.Count, 0);
        }

        [Fact]
        public async Task Should_Success_ReadItem_With_Filter()
        {
            var facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var data = await dataUtil(facade, GetCurrentMethod()).GetTestDataWithStorage();
            var filter = new
            {
                data.UnitId,
                data.StorageId,
                data.DONo
            };
            var Response = facade.ReadURNItem();
            Assert.NotEqual(Response.Data.Count, 0);
        }

        //Monitoring Terima BP
        [Fact]
        public async Task Should_Success_Get_Terima_BP_Report_Data()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillReceptionFacade TerimaBP = new MonitoringCentralBillReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);

            var Response = TerimaBP.GetMonitoringTerimaBonPusatReport(null, null, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_BP_Report_Data_Null_Parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());
            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillReceptionFacade TerimaBP = new MonitoringCentralBillReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);

            DateTime d1 = dataBC.BeacukaiDate.DateTime.AddDays(30);            
            DateTime d2 = dataBC.BeacukaiDate.DateTime.AddDays(30); 

            var Response = TerimaBP.GetMonitoringTerimaBonPusatReport(d1, d2, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_BP_Report_Data_By_User()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());
            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillReceptionFacade TerimaBP = new MonitoringCentralBillReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);

            var Response = TerimaBP.GetMonitoringTerimaBonPusatByUserReport(null, null, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_BP_Report_Data_By_User_Null_Pameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());
            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillReceptionFacade TerimaBP = new MonitoringCentralBillReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);
            DateTime d1 = dataBC.BeacukaiDate.DateTime.AddDays(30);
            DateTime d2 = dataBC.BeacukaiDate.DateTime.AddDays(30);

            var Response = TerimaBP.GetMonitoringTerimaBonPusatByUserReport(d1, d2, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_BP_Report_Excel()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());
            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillReceptionFacade TerimaBP = new MonitoringCentralBillReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);

            var Response = TerimaBP.GenerateExcelMonitoringTerimaBonPusat(null, null, 1, 25, "{}", 7);

            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_BP_Report_Excel_Null_Parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());
            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillReceptionFacade TerimaBP = new MonitoringCentralBillReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);
            DateTime d1 = dataBC.BeacukaiDate.DateTime.AddDays(30);
            DateTime d2 = dataBC.BeacukaiDate.DateTime.AddDays(30);

            var Response = TerimaBP.GenerateExcelMonitoringTerimaBonPusat(d1, d2, 1, 25, "{}", 7);

            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_BP_Report_Excel_By_User()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());
            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillReceptionFacade TerimaBP = new MonitoringCentralBillReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);

            var Response = TerimaBP.GenerateExcelMonitoringTerimaBonPusatByUser(null, null, 1, 25, "{}", 7);

            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_BP_Report_Excel_By_User_Null_Parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());
            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillReceptionFacade TerimaBP = new MonitoringCentralBillReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);
            DateTime d1 = dataBC.BeacukaiDate.DateTime.AddDays(30);
            DateTime d2 = dataBC.BeacukaiDate.DateTime.AddDays(30);

            var Response = TerimaBP.GenerateExcelMonitoringTerimaBonPusatByUser(d1, d2, 1, 25, "{}", 7);

            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }
        //Monitoring Keluar BP
        [Fact]
        public async Task Should_Success_Get_Keluar_BP_Report_Data()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());
            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillExpenditureFacade KeluarBP = new MonitoringCentralBillExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);

            var Response = KeluarBP.GetMonitoringKeluarBonPusatReport(dataBon.ReceiptDate.DateTime, dataBon.ReceiptDate.DateTime, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_BP_Report_Data_Null_Parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());
            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillExpenditureFacade KeluarBP = new MonitoringCentralBillExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);

            DateTime d1 = dataBon.ReceiptDate.DateTime.AddDays(30);
            DateTime d2 = dataBon.ReceiptDate.DateTime.AddDays(30);

            var Response = KeluarBP.GetMonitoringKeluarBonPusatReport(d1, d2, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_BP_Report_Data_By_User()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());
            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillExpenditureFacade KeluarBP = new MonitoringCentralBillExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);

            var Response = KeluarBP.GetMonitoringKeluarBonPusatByUserReport(dataBon.ReceiptDate.DateTime, dataBon.ReceiptDate.DateTime, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_BP_Report_Data_By_User_Null_Pameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());
            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillExpenditureFacade KeluarBP = new MonitoringCentralBillExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);
            DateTime d1 = dataBon.ReceiptDate.DateTime.AddDays(30);
            DateTime d2 = dataBon.ReceiptDate.DateTime.AddDays(30);

            var Response = KeluarBP.GetMonitoringKeluarBonPusatByUserReport(d1, d2, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_BP_Report_Excel()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, datautilBon, garmentBeaCukaiFacade);

            MonitoringCentralBillExpenditureFacade KeluarBP = new MonitoringCentralBillExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataBC = await datautilBC.GetTestDataWithURN(USERNAME);

            var Response = KeluarBP.GenerateExcelMonitoringKeluarBonPusat(null, null, 1, 25, "{}", 7);

            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_BP_Report_Excel_Null_Parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());
            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillExpenditureFacade KeluarBP = new MonitoringCentralBillExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);
            DateTime d1 = dataBon.ReceiptDate.DateTime.AddDays(30);
            DateTime d2 = dataBon.ReceiptDate.DateTime.AddDays(30);

            var Response = KeluarBP.GenerateExcelMonitoringKeluarBonPusat(d1, d2, 1, 25, "{}", 7);

            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_BP_Report_Excel_By_User()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, datautilBon, garmentBeaCukaiFacade);

            MonitoringCentralBillExpenditureFacade KeluarBP = new MonitoringCentralBillExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataBC = await datautilBC.GetTestDataWithURN(USERNAME);

            var Response = KeluarBP.GenerateExcelMonitoringKeluarBonPusatByUser(null, null, 1, 25, "{}", 7);

            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_BP_Report_Excel_By_User_Null_Parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());
            var garmentDeliveryOrder = await Task.Run(() => datautilDO.GetNewData("User"));

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            MonitoringCentralBillExpenditureFacade KeluarBP = new MonitoringCentralBillExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataDO = await datautilDO.GetTestData();
            var dataBon = await datautilBon.GetTestData();
            var dataBC = await datautilBC.GetTestData(USERNAME, garmentDeliveryOrder);
            DateTime d1 = dataBon.ReceiptDate.DateTime.AddDays(30);
            DateTime d2 = dataBon.ReceiptDate.DateTime.AddDays(30);

            var Response = KeluarBP.GenerateExcelMonitoringKeluarBonPusatByUser(d1, d2, 1, 25, "{}", 7);

            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }
        //Monitoring Terima Nota Koreksi
        [Fact]
        public async Task Should_Success_Get_Terima_NK_Report_Data()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, datautilBon, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteReceptionFacade TerimaNK = new MonitoringCorrectionNoteReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            var Response = TerimaNK.GetMonitoringTerimaNKReport(dataNK.CorrectionDate.DateTime, dataNK.CorrectionDate.DateTime, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_NK_Report_Data_Null_Parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, datautilBon, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteReceptionFacade TerimaNK = new MonitoringCorrectionNoteReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            DateTime d1 = dataNK.CorrectionDate.DateTime.AddDays(30);
            DateTime d2 = dataNK.CorrectionDate.DateTime.AddDays(30);

            var Response = TerimaNK.GetMonitoringTerimaNKReport(d1, d2, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_NK_Report_Data_By_User()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, datautilBon, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteReceptionFacade TerimaNK = new MonitoringCorrectionNoteReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            var Response = TerimaNK.GetMonitoringTerimaNKByUserReport(dataNK.CorrectionDate.DateTime, dataNK.CorrectionDate.DateTime, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_NK_Report_Data_By_User_Null_Pameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, datautilBon, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteReceptionFacade TerimaNK = new MonitoringCorrectionNoteReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            DateTime d1 = dataNK.CorrectionDate.DateTime.AddDays(30);
            DateTime d2 = dataNK.CorrectionDate.DateTime.AddDays(30);

            var Response = TerimaNK.GetMonitoringTerimaNKByUserReport(d1, d2, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_NK_Report_Excel()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, datautilBon, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteReceptionFacade TerimaNK = new MonitoringCorrectionNoteReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            var Response = TerimaNK.GenerateExcelMonitoringTerimaNK(null, null, 1, 25, "{}", 7);
            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_NK_Report_Excel_Null_Parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteReceptionFacade TerimaNK = new MonitoringCorrectionNoteReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            MonitoringCentralBillExpenditureFacade KeluarBP = new MonitoringCentralBillExpenditureFacade(_dbContext(GetCurrentMethod()));

            DateTime d1 = dataNK.CorrectionDate.DateTime.AddDays(30);
            DateTime d2 = dataNK.CorrectionDate.DateTime.AddDays(30);

            var Response = TerimaNK.GenerateExcelMonitoringTerimaNK(d1, d2, 1, 25, "{}", 7);

            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_NK_Report_Excel_By_User()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteReceptionFacade TerimaNK = new MonitoringCorrectionNoteReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            var Response = TerimaNK.GenerateExcelMonitoringTerimaNKByUser(null, null, 1, 25, "{}", 7);
            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async Task Should_Success_Get_Terima_NK_Report_Excel_By_User_Null_Parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteReceptionFacade TerimaNK = new MonitoringCorrectionNoteReceptionFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            DateTime d1 = dataNK.CorrectionDate.DateTime.AddDays(30);
            DateTime d2 = dataNK.CorrectionDate.DateTime.AddDays(30);

            var Response = TerimaNK.GenerateExcelMonitoringTerimaNKByUser(d1, d2, 1, 25, "{}", 7);

            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        //Monitoring Keluar Nota Koreksi
        [Fact]
        public async Task Should_Success_Get_Keluar_NK_Report_Data()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, datautilBon, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteExpenditureFacade KeluarNK = new MonitoringCorrectionNoteExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            var Response = KeluarNK.GetMonitoringKeluarNKReport(dataNK.CorrectionDate.DateTime, dataNK.CorrectionDate.DateTime, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_NK_Report_Data_Null_Parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, datautilBon, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteExpenditureFacade KeluarNK = new MonitoringCorrectionNoteExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            DateTime d1 = dataNK.CorrectionDate.DateTime.AddDays(30);
            DateTime d2 = dataNK.CorrectionDate.DateTime.AddDays(30);

            var Response = KeluarNK.GetMonitoringKeluarNKReport(d1, d2, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_NK_Report_Data_By_User()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, datautilBon, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteExpenditureFacade KeluarNK = new MonitoringCorrectionNoteExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            var Response = KeluarNK.GetMonitoringKeluarNKByUserReport(dataNK.CorrectionDate.DateTime, dataNK.CorrectionDate.DateTime, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_NK_Report_Data_By_User_Null_Pameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, datautilBon, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteExpenditureFacade KeluarNK = new MonitoringCorrectionNoteExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            DateTime d1 = dataNK.CorrectionDate.DateTime.AddDays(30);
            DateTime d2 = dataNK.CorrectionDate.DateTime.AddDays(30);

            var Response = KeluarNK.GetMonitoringKeluarNKByUserReport(d1, d2, 1, 25, "{}", 7);
            Assert.NotNull(Response.Item1);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_NK_Report_Excel()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, datautilBon, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteExpenditureFacade KeluarNK = new MonitoringCorrectionNoteExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            var Response = KeluarNK.GenerateExcelMonitoringKeluarNK(null, null, 1, 25, "{}", 7);
            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_NK_Report_Excel_Null_Parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteExpenditureFacade KeluarNK = new MonitoringCorrectionNoteExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            DateTime d1 = dataNK.CorrectionDate.DateTime.AddDays(30);
            DateTime d2 = dataNK.CorrectionDate.DateTime.AddDays(30);

            var Response = KeluarNK.GenerateExcelMonitoringKeluarNK(d1, d2, 1, 25, "{}", 7);

            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_NK_Report_Excel_By_User()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteExpenditureFacade KeluarNK = new MonitoringCorrectionNoteExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            var Response = KeluarNK.GenerateExcelMonitoringKeluarNKByUser(null, null, 1, 25, "{}", 7);
            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async Task Should_Success_Get_Keluar_NK_Report_Excel_By_User_Null_Parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));

            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentunitreceiptnoteFacade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilBon = new GarmentUnitReceiptNoteDataUtil(garmentunitreceiptnoteFacade, datautilDO);

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            MonitoringCorrectionNoteExpenditureFacade KeluarNK = new MonitoringCorrectionNoteExpenditureFacade(_dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();

            DateTime d1 = dataNK.CorrectionDate.DateTime.AddDays(30);
            DateTime d2 = dataNK.CorrectionDate.DateTime.AddDays(30);

            var Response = KeluarNK.GenerateExcelMonitoringKeluarNKByUser(d1, d2, 1, 25, "{}", 7);

            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        //Buku Harian Pembelian
        [Fact]
        public async Task Should_Success_Get_Buku_Sub_Beli_Data()
        {
            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            GarmentDailyPurchasingReportFacade DataNK = new GarmentDailyPurchasingReportFacade(ServiceProvider, _dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();
            DateTime d1 = dataNK.CorrectionDate.DateTime;
            DateTime d2 = dataNK.CorrectionDate.DateTime;

            var Response = DataNK.GetGDailyPurchasingReport(null, true, null, null, null, 7);
            Assert.NotNull(Response.Item1);
            Assert.NotEqual(0, Response.Item2);
        }

        [Fact]
        public async Task Should_Success_Get_Buku_Sub_Beli_Null_Parameter()
        {
            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            GarmentDailyPurchasingReportFacade DataNK = new GarmentDailyPurchasingReportFacade(ServiceProvider, _dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();
            DateTime d1 = dataNK.CorrectionDate.DateTime;
            DateTime d2 = dataNK.CorrectionDate.DateTime;

            var Response = DataNK.GetGDailyPurchasingReport(null, true, null, null, null, 7);
            Assert.NotNull(Response.Item1);
            Assert.NotEqual(0, Response.Item2);
        }

        [Fact]
        public async Task Should_Success_Get_Buku_Sub_Beli_Excel()
        {
            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            GarmentDailyPurchasingReportFacade DataNK = new GarmentDailyPurchasingReportFacade(ServiceProvider, _dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();
            DateTime d1 = dataNK.CorrectionDate.DateTime;
            DateTime d2 = dataNK.CorrectionDate.DateTime;

            var Response = DataNK.GenerateExcelGDailyPurchasingReport(null, true, null, null, null, 7);
            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async Task Should_Success_Get_Buku_Sub_Beli_Excel_Null_Parameter()
        {
            GarmentDeliveryOrderFacade facadeDO = new GarmentDeliveryOrderFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilDO = dataUtilDO(facadeDO, GetCurrentMethod());

            var garmentBeaCukaiFacade = new GarmentBeacukaiFacade(_dbContext(GetCurrentMethod()), GetServiceProvider());
            var datautilBC = new GarmentBeacukaiDataUtil(datautilDO, garmentBeaCukaiFacade);

            var garmentCorrectionNoteFacade = new GarmentCorrectionNotePriceFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var datautilCN = new GarmentCorrectionNoteDataUtil(garmentCorrectionNoteFacade, datautilBC, datautilDO);

            GarmentDailyPurchasingReportFacade DataNK = new GarmentDailyPurchasingReportFacade(ServiceProvider, _dbContext(GetCurrentMethod()));

            var dataNK = await datautilCN.GetTestDataNotaKoreksi();
            DateTime d1 = dataNK.CorrectionDate.DateTime.AddDays(30);
            DateTime d2 = dataNK.CorrectionDate.DateTime.AddDays(30);

            var Response = DataNK.GenerateExcelGDailyPurchasingReport(null, true, null, null, null, 7);
            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        #region flow detail penerimaan 

        [Fact]
        public async void Should_Success_Get_FlowReport_Data()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var model = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var Response = facade.GetReportFlow(DateTime.MinValue, DateTime.MaxValue, model.UnitCode, "", 1, 25, "{}", 7);
            Assert.NotEqual(Response.Item1.Count, 0);
        }

        [Fact]
        public async void Should_Success_Get_FlowReport_Data_Null_Parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var model = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var Response = facade.GetReportFlow(null, null, "", "Bahan Baku", 1, 25, "{}", 7);
            Assert.Equal(Response.Item1.Count, 0);
        }

        [Fact]
        public async void Should_Success_Get_FlowReport_Data_Excel()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var model = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var Response = facade.GenerateExcelLow(DateTime.MinValue, DateTime.MaxValue, model.UnitCode, "", 7);
            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async void Should_Success_Get_Report_Data_Excel_Null_parameter()
        {
            GarmentUnitReceiptNoteFacade facade = new GarmentUnitReceiptNoteFacade(GetServiceProvider(), _dbContext(GetCurrentMethod()));
            var model = await dataUtil(facade, GetCurrentMethod()).GetTestData();
            var Response = facade.GenerateExcelLow(null, null, "0", "", 7);
            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        #endregion
    }


}
