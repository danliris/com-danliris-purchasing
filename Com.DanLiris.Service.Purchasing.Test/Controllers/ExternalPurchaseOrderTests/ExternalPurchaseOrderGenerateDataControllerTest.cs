﻿using Com.DanLiris.Service.Purchasing.Lib.Models.ExternalPurchaseOrderModel;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.ExternalPurchaseOrderViewModel;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.ExternalPurchaseOrderDataUtils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Xunit;
namespace Com.DanLiris.Service.Purchasing.Test.Controllers.ExternalPurchaseOrderTests
{
    [Collection("TestServerFixture Collection")]
    public class ExternalPurchaseOrderGenerateDataControllerTest
    {
        private const string MediaType = "application/json";
        private readonly string URI = "v1/generating-data/purchase-order-external";

        private TestServerFixture TestFixture { get; set; }

        private HttpClient Client
        {
            get { return this.TestFixture.Client; }
        }

        protected ExternalPurchaseOrderDataUtil DataUtil
        {
            get { return (ExternalPurchaseOrderDataUtil)this.TestFixture.Service.GetService(typeof(ExternalPurchaseOrderDataUtil)); }
        }

        public ExternalPurchaseOrderGenerateDataControllerTest(TestServerFixture fixture)
        {
            TestFixture = fixture;
        }

        [Fact]
        public async Task Should_Success_Get_Report_Excel()
        {
            var response = await this.Client.GetAsync(URI + "/download");
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        }

        [Fact]
        public async Task Should_Error_Get_Report_Excel()
        {
            var response = await this.Client.GetAsync(URI + "/display");
            Assert.Equal(HttpStatusCode.InternalServerError, response.StatusCode);
        }
    }
}
