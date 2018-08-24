﻿using Com.DanLiris.Service.Purchasing.Lib.Facades.ExternalPurchaseOrderFacade;
using Com.DanLiris.Service.Purchasing.Lib.Models.ExternalPurchaseOrderModel;
using Com.DanLiris.Service.Purchasing.Lib.Services;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.ExternalPurchaseOrderDataUtils;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.InternalPurchaseOrderDataUtils;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.DeliveryOrderDataUtils;
using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace Com.DanLiris.Service.Purchasing.Test.Facades.ExternalPurchaseOrderTests
{
    [Collection("ServiceProviderFixture Collection")]
    public class MonitoringTest
    {
        private IServiceProvider ServiceProvider { get; set; }

        public MonitoringTest(ServiceProviderFixture fixture)
        {
            ServiceProvider = fixture.ServiceProvider;

            IdentityService identityService = (IdentityService)ServiceProvider.GetService(typeof(IdentityService));
            identityService.Username = "Unit Test";
        }

        private InternalPurchaseOrderDataUtil DataUtil
        {
            get { return (InternalPurchaseOrderDataUtil)ServiceProvider.GetService(typeof(InternalPurchaseOrderDataUtil)); }
        }

        private ExternalPurchaseOrderDataUtil EPODataUtil
        {
            get { return (ExternalPurchaseOrderDataUtil)ServiceProvider.GetService(typeof(ExternalPurchaseOrderDataUtil)); }
        }
        private DeliveryOrderDataUtil DODataUtil
        {
            get { return (DeliveryOrderDataUtil)ServiceProvider.GetService(typeof(DeliveryOrderDataUtil)); }
        }

        private ExternalPurchaseOrderFacade Facade
        {
            get { return (ExternalPurchaseOrderFacade)ServiceProvider.GetService(typeof(ExternalPurchaseOrderFacade)); }
        }

        //Duration PO In-PO Ex
        [Fact]
        public async void Should_Success_Get_Report_POExDODuration_Data()
        {
            var model = await DODataUtil.GetTestData2("Unit test");
            var model2 = await EPODataUtil.GetTestData3("Unit test");
            var Response = Facade.GetEPODODurationReport(model2.UnitId, "31-60 hari", null, null, 1, 25, "{}", 7);
            Assert.NotEqual(Response.Item2, 0);
        }

        [Fact]
        public async void Should_Success_Get_Report_POExDODuration_Null_Parameter()
        {
            var model = await EPODataUtil.GetTestData3("Unit test");
            var model3 = await EPODataUtil.GetTestData3("Unit test");
            var Response = Facade.GetEPODODurationReport("", "61-90 hari", null, null, 1, 25, "{}", 7);
            Assert.NotEqual(Response.Item2, 0);
        }

        [Fact]
        public async void Should_Success_Get_Report_POEDODuration_Excel()
        {
            var model = await DODataUtil.GetTestData2("Unit test");
            var model2 = await EPODataUtil.GetTestData2("Unit test");
            var Response = Facade.GenerateExcelEPODODuration(model2.UnitId, "31-60 hari", null, null, 7);
            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }

        [Fact]
        public async void Should_Success_Get_Report_POEDODuration_Excel_Null_Parameter()
        {
            var model = await DODataUtil.GetTestData3("Unit test");
            var model3 = await EPODataUtil.GetTestData3("Unit test");
            var Response = Facade.GenerateExcelEPODODuration("", "61-90 hari", null, null, 7);
            Assert.IsType(typeof(System.IO.MemoryStream), Response);
        }
    }
}
