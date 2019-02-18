﻿using Com.DanLiris.Service.Purchasing.Lib.Models.ExternalPurchaseOrderModel;
using Com.DanLiris.Service.Purchasing.Lib.Models.UnitPaymentOrderModel;
using Com.DanLiris.Service.Purchasing.Lib.Models.UnitReceiptNoteModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace Com.DanLiris.Service.Purchasing.Lib.Interfaces
{
    public interface IUnitPaymentOrderFacade
    {
        Tuple<List<UnitPaymentOrder>, int, Dictionary<string, string>> Read(int Page = 1, int Size = 25, string Order = "{}", string Keyword = null, string Filter = "{}");
        UnitPaymentOrder ReadById(int id);
        Task<int> Create(UnitPaymentOrder model, string username, bool isImport, int clientTimeZoneOffset = 7);
        Task<int> Update(int id, UnitPaymentOrder model, string user);
        Task<int> Delete(int id, string username);
        Tuple<List<UnitPaymentOrder>, int, Dictionary<string, string>> ReadSpb(int Page = 1, int Size = 25, string Order = "{}", string Keyword = null, string Filter = "{}");
        Tuple<List<UnitPaymentOrder>, int, Dictionary<string, string>> ReadPositionFiltered(int Page = 1, int Size = 25, string Order = "{}", string Keyword = null, string Filter = "{}");
        UnitReceiptNote GetUnitReceiptNote(long URNId);
        ExternalPurchaseOrder GetExternalPurchaseOrder(string EPONo);
        Tuple<List<ViewModels.UnitPaymentOrderViewModel.UnitPaymentOrderReportViewModel>, int> GetReportAll(string unitId, string supplierId, DateTime? dateFrom, DateTime? dateTo, int page, int size, string Order, int offset);
        MemoryStream GenerateExcel(string unitId, string supplierId, DateTime? dateFrom, DateTime? dateTo, int offset);
        MemoryStream GenerateDataExcel(DateTime? dateFrom, DateTime? dateTo, int offset);
    }
}
