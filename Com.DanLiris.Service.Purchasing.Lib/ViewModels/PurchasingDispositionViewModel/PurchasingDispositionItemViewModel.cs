﻿using Com.DanLiris.Service.Purchasing.Lib.ViewModels.NewIntegrationViewModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Com.DanLiris.Service.Purchasing.Lib.ViewModels.PurchasingDispositionViewModel
{
    public class PurchasingDispositionItemViewModel
    {
        public string EPONo { get; set; }
        public long EPOId { get; set; }
        public UnitViewModel Unit { get; set; }
        public bool UseVat { get; set; }
        public bool UseIncomeTax { get; set; }
        public IncomeTaxViewModel IncomeTax { get; set; }

        public virtual List<PurchasingDispositionDetailViewModel> Details { get; set; }

    }
}
