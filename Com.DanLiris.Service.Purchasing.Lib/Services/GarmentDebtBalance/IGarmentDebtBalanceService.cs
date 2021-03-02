﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Com.DanLiris.Service.Purchasing.Lib.Services.GarmentDebtBalance
{
    public interface IGarmentDebtBalanceService
    {
        Task<int> CreateFromCustoms(CustomsFormDto form);
        Task<int> UpdateFromInvoice(int deliveryOrderId, InvoiceFormDto form);
        Task<int> UpdateFromInternalNote(int deliveryOrderId, InternalNoteFormDto form);
        Task<int> UpdateFromBankExpenditureNote(int deliveryOrderId, BankExpenditureNoteFormDto form);
    }
}
