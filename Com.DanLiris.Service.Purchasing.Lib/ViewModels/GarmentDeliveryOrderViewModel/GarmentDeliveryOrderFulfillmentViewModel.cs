﻿using Com.DanLiris.Service.Purchasing.Lib.Utilities;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.NewIntegrationViewModel;

namespace Com.DanLiris.Service.Purchasing.Lib.ViewModels.GarmentDeliveryOrderViewModel
{
    public class GarmentDeliveryOrderFulfillmentViewModel : BaseViewModel
    {
        public long ePOItemId { get; set; }
        public int pOId { get; set; }
        public int pOItemId { get; set; }
        public long pRId { get; set; }
        public string pRNo { get; set; }
        public long pRItemId { get; set; }
        public string poSerialNumber { get; set; }

        public UnitViewModel unit { get; set; }

        public GarmentProductViewModel product { get; set; }
        public double doQuantity { get; set; }
        public double dealQuantity { get; set; }
        public double conversion { get; set; }
        public UomViewModel purchaseOrderUom { get; set; } // UOM

        public double smallQuantity { get; set; }
        public UomViewModel smallUom { get; set; }

        public double PricePerDealUnit { get; set; }
        public double PriceTotal { get; set; }


        public string rONo { get; set; }
        public double receiptQuantity { get; set; }
    }
}
