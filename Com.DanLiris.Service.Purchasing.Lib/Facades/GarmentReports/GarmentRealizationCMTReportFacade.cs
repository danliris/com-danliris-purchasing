using Com.DanLiris.Service.Purchasing.Lib.Helpers;
using Com.DanLiris.Service.Purchasing.Lib.Interfaces;
using Com.DanLiris.Service.Purchasing.Lib.Services;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.GarmentReports;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.NewIntegrationViewModel;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.NewIntegrationViewModel.GarmentExpenditureGood;
using Com.Moonlay.NetCore.Lib;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentReports
{
    public class GarmentRealizationCMTReportFacade : IGarmentRealizationCMTReportFacade
    {
        private readonly IServiceProvider serviceProvider;
        private readonly IdentityService identityService;
        private readonly PurchasingDbContext dbContext;

        public GarmentRealizationCMTReportFacade(IServiceProvider serviceProvider, PurchasingDbContext dbContext)
        {
            this.serviceProvider = serviceProvider;
            identityService = (IdentityService)serviceProvider.GetService(typeof(IdentityService));

            this.dbContext = dbContext;
        }

        public IQueryable<GarmentRealizationCMTReportViewModel> GetQuery(DateTime? dateFrom, DateTime? dateTo, long unit, int offset)
        {
            DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : (DateTime)dateFrom;
            DateTime DateTo = dateTo == null ? DateTime.Now : (DateTime)dateTo;
            List<GarmentRealizationCMTReportViewModel> realizationCMT = new List<GarmentRealizationCMTReportViewModel>();
            var Query = (from a in dbContext.GarmentUnitExpenditureNotes
                         join b in dbContext.GarmentUnitExpenditureNoteItems on a.Id equals b.UENId
                         join c in dbContext.GarmentUnitDeliveryOrders on a.UnitDONo equals c.UnitDONo
                         join i in dbContext.GarmentUnitDeliveryOrderItems on c.Id equals i.UnitDOId
                         join e in dbContext.GarmentUnitReceiptNoteItems on b.URNItemId equals e.Id
                         join d in dbContext.GarmentUnitReceiptNotes on e.URNId equals d.Id
                         join f in dbContext.GarmentDeliveryOrderDetails on i.DODetailId equals f.Id
                         join g in dbContext.GarmentDeliveryOrderItems on f.GarmentDOItemId equals g.Id
                         join h in dbContext.GarmentDeliveryOrders on g.GarmentDOId equals h.Id
                         join j in dbContext.GarmentExternalPurchaseOrders on g.EPOId equals j.Id
                         where a.ExpenditureDate.AddHours(offset).Date >= DateFrom.Date
                               && a.ExpenditureDate.AddHours(offset).Date <= DateTo.Date
                               && a.UnitSenderId == (unit == 0 ? a.UnitSenderId : unit)
                               && b.ProductName == "FABRIC"
                               && (j.PaymentMethod == "CMT" || j.PaymentMethod == "FREE FROM BUYER")

                         select new GarmentRealizationCMTReportViewModel
                         {
                             UENNo = a.UENNo,
                             ProductRemark = b.ProductRemark,
                             Quantity = b.Quantity,
                             EAmountVLS = (decimal)b.Quantity * e.PricePerDealUnit,
                             EAmountIDR = (decimal)b.Quantity * e.PricePerDealUnit * (decimal)h.DOCurrencyRate,
                             RONo = c.RONo,
                             URNNo = d.URNNo,
                             ProductRemark2 = e.ProductRemark,
                             ReceiptQuantity = e.ReceiptQuantity,
                             UAmountVLS = e.ReceiptQuantity * e.PricePerDealUnit,
                             UAmountIDR = e.ReceiptQuantity * e.PricePerDealUnit * (decimal)h.DOCurrencyRate,
                             SupplierName = d.SupplierName,
                             BillNo = h.BillNo,
                             PaymentBill = h.PaymentBill,
                             DONo = h.DONo,
                             UENPrice = e.PricePerDealUnit,
                             DORate = (decimal)h.DOCurrencyRate
                         });

            var QueryGroup = from a in Query
                             group a by new { a.UENNo, a.RONo, a.URNNo, a.BillNo, a.PaymentBill } into groupdata
                             select new GarmentRealizationCMTReportViewModel
                             {
                                 UENNo = groupdata.FirstOrDefault().UENNo,
                                 ProductRemark = groupdata.FirstOrDefault().ProductRemark,
                                 Quantity = groupdata.Sum(x => x.Quantity),
                                 EAmountVLS = groupdata.Sum(x => x.EAmountVLS),
                                 EAmountIDR = groupdata.Sum(x => x.EAmountIDR),
                                 RONo = groupdata.FirstOrDefault().RONo,
                                 URNNo = groupdata.FirstOrDefault().URNNo,
                                 ProductRemark2 = groupdata.FirstOrDefault().ProductRemark2,
                                 ReceiptQuantity = groupdata.Sum(x => x.ReceiptQuantity),
                                 UAmountVLS = groupdata.Sum(x => x.UAmountVLS),
                                 UAmountIDR = groupdata.Sum(x => x.UAmountIDR),
                                 SupplierName = groupdata.FirstOrDefault().SupplierName,
                                 BillNo = groupdata.FirstOrDefault().BillNo,
                                 PaymentBill = groupdata.FirstOrDefault().PaymentBill,
                                 DONo = groupdata.FirstOrDefault().DONo,
                                 UENPrice = groupdata.FirstOrDefault().UENPrice,
                                 DORate = groupdata.FirstOrDefault().DORate,
                             };

            foreach (GarmentRealizationCMTReportViewModel i in QueryGroup)
            {
                var data1 = GetExpenditureGood(i.RONo);

                realizationCMT.Add(new GarmentRealizationCMTReportViewModel
                {
                    UENNo = i.UENNo,
                    ProductRemark = i.ProductRemark,
                    Quantity = i.Quantity,
                    EAmountVLS = i.EAmountVLS,
                    EAmountIDR = i.EAmountIDR,
                    RONo = i.RONo,
                    URNNo = i.URNNo,
                    ProductRemark2 = i.ProductRemark2,
                    ReceiptQuantity = i.ReceiptQuantity,
                    UAmountVLS = i.UAmountVLS,
                    UAmountIDR = i.UAmountIDR,
                    SupplierName = i.SupplierName,
                    BillNo = i.BillNo,
                    PaymentBill = i.PaymentBill,
                    DONo = i.DONo,
                    InvoiceNo = data1 == null ? "-" : data1.FirstOrDefault().Invoice,
                    ExpenditureGoodNo = data1 == null ? "-" : data1.FirstOrDefault().ExpenditureGoodNo,
                    Article = data1 == null ? "-" : data1.FirstOrDefault().Article,
                    UnitQty = data1 == null ? 0 : data1.FirstOrDefault().TotalQuantity,
                    EGAmountIDR = data1 == null ? 0 : (decimal)data1.FirstOrDefault().TotalPrice,
                });
            };

            return realizationCMT.AsQueryable();

        }

        public Tuple<List<GarmentRealizationCMTReportViewModel>, int> GetReport(DateTime? dateFrom, DateTime? dateTo, long unit, int page, int size, string Order, int offset)
        {
            var Query = GetQuery(dateFrom, dateTo, unit, offset);

            var b = Query.ToArray();
            var index = 0;

            foreach (GarmentRealizationCMTReportViewModel a in Query)
            {
                GarmentRealizationCMTReportViewModel dup = Array.Find(b, o => o.InvoiceNo == a.InvoiceNo && o.ExpenditureGoodNo == a.ExpenditureGoodNo);
                if (dup != null)
                {
                    if (dup.Count == 0)
                    {
                        index++;
                        dup.Count = index;
                    }
                }
                a.Count = dup.Count;
            }

            Pageable<GarmentRealizationCMTReportViewModel> pageable = new Pageable<GarmentRealizationCMTReportViewModel>(Query.OrderBy(o => o.InvoiceNo).ThenBy(o => o.ExpenditureGoodNo), page - 1, size);
            List<GarmentRealizationCMTReportViewModel> Data = pageable.Data.ToList<GarmentRealizationCMTReportViewModel>();
            int TotalData = pageable.TotalCount;

            return Tuple.Create(Data, TotalData);
        }

        public MemoryStream GenerateExcel(DateTime? dateFrom, DateTime? dateTo, long unit, int offset, string unitname)
        {
            DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : (DateTime)dateFrom;
            DateTime DateTo = dateTo == null ? DateTime.Now : (DateTime)dateTo;
            var Query = GetQuery(dateFrom, dateTo, unit, offset);
            var headers = new string[] { "No", "No Invoice", "No. BON", "RO", "Artikel", "Qty BJ", "Fabric Cost" };
            var subheaders = new string[] { "No. BON", "Keterangan", "Qty", "Amount Valas", "Amount IDR", "Asal", "No. BON", "Keterangan", "Qty", "Amount Valas", "Amount IDR", "Supplier", "No Nota", "No BON Kecil", "Surat Jalan", "No Kasbon" };
            DataTable result = new DataTable();
            result.Columns.Add(new DataColumn() { ColumnName = "No", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "No Invoice", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "No. BON", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "RO", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Artikel", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Qty BJ", DataType = typeof(double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Fabric Cost", DataType = typeof(double) });
            result.Columns.Add(new DataColumn() { ColumnName = "No. BON Pemakaian", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Keterangan Pemakaian", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Qty Pemakaian", DataType = typeof(double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Amount Valas - Pemakaian", DataType = typeof(double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Amount IDR - Pemakaian", DataType = typeof(double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Asal Pemakaian", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "No. BON Peneriamaan", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Keterangan Peneriamaan", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Qty Peneriamaan", DataType = typeof(double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Amount Valas - Penerimaan", DataType = typeof(double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Amount IDR - Penerimaan", DataType = typeof(double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Supplier", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "No Nota", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "No BON Kecil", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Surat Jalan", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "No Kasbon", DataType = typeof(String) });

            ExcelPackage package = new ExcelPackage();
            if (Query.ToArray().Count() == 0)
            {
                result.Rows.Add("", "", "", "", "", 0, 0, "", "", 0, 0, 0, "", "", "", 0, 0, 0, "", "", "", "", "");
                var sheet = package.Workbook.Worksheets.Add("Data");
                sheet.Cells["A7"].LoadFromDataTable(result, false, OfficeOpenXml.Table.TableStyles.Light1);// to allow column name to be generated properly for empty data as template
            }
            else
            {
                var Qr = Query.ToArray();
                var q = Query.ToList();
                var index = 0;
                //foreach (GarmentRealizationCMTReportViewModel a in q)
                //{
                //    GarmentRealizationCMTReportViewModel dup = Array.Find(Qr, o => o.InvoiceNo == a.InvoiceNo && o.ExpenditureGoodNo == a.ExpenditureGoodNo);
                //    if (dup != null)
                //    {
                //        if (dup.Count == 0)
                //        {
                //            index++;
                //            dup.Count = index;
                //        }
                //    }
                //    a.Count = dup.Count;
                //}
                Query = q.AsQueryable().OrderBy(o => o.InvoiceNo).ThenBy(o => o.ExpenditureGoodNo);
                int indexz = 0;
                foreach (var item in Query)
                {
                    indexz++;
                    result.Rows.Add(indexz, item.InvoiceNo, item.ExpenditureGoodNo, item.RONo, item.Article, item.UnitQty, item.EGAmountIDR, item.UENNo, item.ProductRemark, item.Quantity, item.EAmountVLS,
                                    item.EAmountIDR, item.RONo, item.URNNo, item.ProductRemark2, item.ReceiptQuantity, item.UAmountVLS, item.UAmountIDR, item.SupplierName, item.BillNo, item.PaymentBill, item.DONo, "");
                }

                // bool styling = true;

                foreach (KeyValuePair<DataTable, String> item in new List<KeyValuePair<DataTable, string>>() { new KeyValuePair<DataTable, string>(result, "Territory") })
                {
                    var sheet = package.Workbook.Worksheets.Add(item.Value);
                    #region KopTable
                    sheet.Cells[$"A1:Q1"].Value = "LAPORAN DATA REALISASI CMT GARMENT";
                    sheet.Cells[$"A1:Q1"].Merge = true;
                    sheet.Cells[$"A1:Q1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    sheet.Cells[$"A1:Q1"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    sheet.Cells[$"A1:Q1"].Style.Font.Bold = true;
                    sheet.Cells[$"A2:Q2"].Value = string.Format("Periode Tanggal {0} s/d {1}", DateFrom.ToString("dd MMM yyyy", new CultureInfo("id-ID")), DateTo.ToString("dd MMM yyyy", new CultureInfo("id-ID")));
                    sheet.Cells[$"A2:Q2"].Merge = true;
                    sheet.Cells[$"A2:Q2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    sheet.Cells[$"A2:Q2"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    sheet.Cells[$"A2:Q2"].Style.Font.Bold = true;
                    sheet.Cells[$"A3:Q3"].Value = string.Format("Konfeksi {0}", string.IsNullOrWhiteSpace(unitname) ? "ALL" : unitname);
                    sheet.Cells[$"A3:Q3"].Merge = true;
                    sheet.Cells[$"A3:Q3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    sheet.Cells[$"A3:Q3"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    sheet.Cells[$"A3:Q3"].Style.Font.Bold = true;
                    #endregion


                    sheet.Cells["A8"].LoadFromDataTable(item.Key, false, OfficeOpenXml.Table.TableStyles.Light16);
                    sheet.Cells["H6"].Value = "BON PEMAKAIAN";
                    sheet.Cells["H6:L6"].Merge = true;
                    sheet.Cells["H6:L6"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                    sheet.Cells["M6"].Value = "BON PENERIMAAN";
                    sheet.Cells["M6:V6"].Merge = true;
                    sheet.Cells["M6:V6"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                    foreach (var i in Enumerable.Range(0, 7))
                    {
                        var col = (char)('A' + i);
                        sheet.Cells[$"{col}6"].Value = headers[i];
                        sheet.Cells[$"{col}6:{col}7"].Merge = true;
                        sheet.Cells[$"{col}6:{col}7"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                    }
                    foreach (var i in Enumerable.Range(0, 16))
                    {
                        var col = (char)('H' + i);
                        sheet.Cells[$"{col}7"].Value = subheaders[i];
                        sheet.Cells[$"{col}7"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                    }
                    sheet.Cells["A6:W7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells["A6:W7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    sheet.Cells["A6:W7"].Style.Font.Bold = true;
                    //sheet.Cells["C1:D1"].Merge = true;
                    //sheet.Cells["C1:D1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    //sheet.Cells["E1:F1"].Merge = true;
                    //sheet.Cells["C1:D1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    Dictionary<string, int> counts = new Dictionary<string, int>();
                    Dictionary<string, int> countsType = new Dictionary<string, int>();
                    var docNo = Query.ToArray();
                    //int value;

                    //foreach (var a in Query)
                    //{
                    //    //FactBeacukaiViewModel dup = Array.Find(docNo, o => o.BCType == a.BCType && o.BCNo == a.BCNo);
                    //    if (counts.TryGetValue(a.InvoiceNo + a.ExpenditureGoodNo, out value))
                    //    {
                    //        counts[a.InvoiceNo + a.ExpenditureGoodNo]++;
                    //    }
                    //    else
                    //    {
                    //        counts[a.InvoiceNo + a.ExpenditureGoodNo] = 1;
                    //    }


                    //    //FactBeacukaiViewModel dup1 = Array.Find(docNo, o => o.BCType == a.BCType);
                    //    if (countsType.TryGetValue(a.InvoiceNo, out value))
                    //    {
                    //        countsType[a.InvoiceNo]++;
                    //    }
                    //    else
                    //    {
                    //        countsType[a.InvoiceNo] = 1;
                    //    }
                    //}

                    index = 8;
                    foreach (KeyValuePair<string, int> b in counts)
                    {
                        sheet.Cells["A" + index + ":A" + (index + b.Value - 1)].Merge = true;
                        sheet.Cells["A" + index + ":A" + (index + b.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                        sheet.Cells["D" + index + ":D" + (index + b.Value - 1)].Merge = true;
                        sheet.Cells["D" + index + ":D" + (index + b.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                        sheet.Cells["E" + index + ":E" + (index + b.Value - 1)].Merge = true;
                        sheet.Cells["E" + index + ":E" + (index + b.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                        sheet.Cells["F" + index + ":F" + (index + b.Value - 1)].Merge = true;
                        sheet.Cells["F" + index + ":F" + (index + b.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                        sheet.Cells["F" + index + ":F" + (index + b.Value - 1)].Merge = true;
                        sheet.Cells["F" + index + ":F" + (index + b.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                        sheet.Cells["G" + index + ":G" + (index + b.Value - 1)].Merge = true;
                        sheet.Cells["G" + index + ":G" + (index + b.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                        sheet.Cells["W" + index + ":W" + (index + b.Value - 1)].Merge = true;
                        sheet.Cells["W" + index + ":W" + (index + b.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                        index += b.Value;
                    }

                    index = 8;
                    foreach (KeyValuePair<string, int> c in countsType)
                    {
                        sheet.Cells["B" + index + ":B" + (index + c.Value - 1)].Merge = true;
                        sheet.Cells["B" + index + ":B" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                        index += c.Value;
                    }
                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                }
            }
            MemoryStream stream = new MemoryStream();
            package.SaveAs(stream);
            return stream;
        }

        public List<GarmentExpenditureGoodViewModel> GetExpenditureGood(string RONo)
        {

            string expenditureUri = "expenditure-goods/byRO";

            IHttpClientService httpClient = (IHttpClientService)serviceProvider.GetService(typeof(IHttpClientService));

            var response = httpClient.GetAsync($"{APIEndpoint.GarmentProduction}{expenditureUri}?RONo={RONo}").Result;
            if (response.IsSuccessStatusCode)
            {
                var content = response.Content.ReadAsStringAsync().Result;
                Dictionary<string, object> result = JsonConvert.DeserializeObject<Dictionary<string, object>>(content);

                List<GarmentExpenditureGoodViewModel> viewModel;
                if (result.GetValueOrDefault("data") == null)
                {
                    viewModel = null;
                }
                else
                {
                    var viewModels = JsonConvert.DeserializeObject<List<GarmentExpenditureGoodViewModel>>(result.GetValueOrDefault("data").ToString());
                    viewModel = viewModels.Count() == 0 ? null : viewModels;
                }
                return viewModel;
            }
            else
            {
                return null;
            }
        }
    }
}
