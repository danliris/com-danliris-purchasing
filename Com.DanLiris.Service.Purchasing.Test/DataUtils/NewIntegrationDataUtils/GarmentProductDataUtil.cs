using Com.DanLiris.Service.Purchasing.Lib.ViewModels.NewIntegrationViewModel;
using Com.DanLiris.Service.Purchasing.WebApi.Helpers;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace Com.DanLiris.Service.Purchasing.Test.DataUtils.NewIntegrationDataUtils
{
    public class GarmentProductDataUtil
    {
        public GarmentProductViewModel GetNewData()
        {
            long nowTicks = DateTimeOffset.Now.Ticks;

            var data = new GarmentProductViewModel
            {
                Code = "CodeTest123",
                Name = "Name123",
            };
            return data;
        }

        public GarmentProductViewModel GetNewDataBP()
        {
            long nowTicks = DateTimeOffset.Now.Ticks;

            var data = new GarmentProductViewModel
            {
                Code = "CodeTestBP123",
                Name = "Name123BP",
            };
            return data;
        }

        public Dictionary<string, object> GetResultFormatterOk()
        {
            var data = GetNewData();

            Dictionary<string, object> result =
                new ResultFormatter("1.0", General.OK_STATUS_CODE, General.OK_MESSAGE)
                .Ok(data);

            return result;
        }

        public Dictionary<string, object> GetResultFormatterNullOk()
        {
            var data = GetNewData();

            Dictionary<string, object> result =
                new ResultFormatter("1.0", General.OK_STATUS_CODE, General.OK_MESSAGE)
                .Ok(null);

            return result;
        }

        public Dictionary<string, object> GetMultipleResultFormatterOk()
        {
            var data = new List<GarmentProductViewModel> { GetNewData(), GetNewDataBP() };

            Dictionary<string, object> result =
                new ResultFormatter("1.0", General.OK_STATUS_CODE, General.OK_MESSAGE)
                .Ok(data);

            return result;
        }

        public string GetResultFormatterOkString()
        {
            var result = GetResultFormatterOk();

            return JsonConvert.SerializeObject(result);
        }

        public string GetMultipleResultFormatterOkString()
        {
            var result = GetMultipleResultFormatterOk();

            return JsonConvert.SerializeObject(result);
        }

        public string GetMultipleResultFormatterNullOkString()
        {
            var result = GetResultFormatterNullOk();

            return JsonConvert.SerializeObject(result);
        }
    }
}
