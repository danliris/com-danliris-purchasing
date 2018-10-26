﻿using Com.DanLiris.Service.Purchasing.Lib.Utilities;
using System.Collections.Generic;

namespace Com.DanLiris.Service.Purchasing.Lib.ViewModels.GarmentDeliveryOrderViewModel
{
    public class GarmentDeliveryOrderItemViewModel : BaseViewModel
    {
        public PurchaseOrderExternal purchaseOrderExternal { get; set; }
        public List<GarmentDeliveryOrderFulfillmentViewModel> fulfillments { get; set; }
        public string paymentType { get; set; }
        public string paymentMethod { get; set; }
        public int paymentDueDays { get; set; }
        public string pONo { get; set; }
        public int pOId { get; set; }
    }

    public class PurchaseOrderExternal
    {
        public long Id { get; set; }
        public string no { get; set; }
    }
}
