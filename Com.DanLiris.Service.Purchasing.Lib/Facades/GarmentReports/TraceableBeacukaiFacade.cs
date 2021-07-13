using Com.DanLiris.Service.Purchasing.Lib.ViewModels.GarmentReports;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.NewIntegrationViewModel.GarmentExpenditureGood;
using System.Net.Http;
using Newtonsoft.Json;
using Com.DanLiris.Service.Purchasing.Lib.Helpers;
using Com.DanLiris.Service.Purchasing.Lib.Interfaces;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.NewIntegrationViewModel.GarmentFinishingOut;
using System.IO;
using System.Data;
using OfficeOpenXml;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.NewIntegrationViewModel.BeacukaiAdded;
using System.Globalization;
using OfficeOpenXml.Style;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.NewIntegrationViewModel.GarmentCuttingOut;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.NewIntegrationViewModel.GarmentPreparing;

namespace Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentReports
{
    public class TraceableBeacukaiFacade : ITraceableBeacukaiFacade
    {
        private readonly PurchasingDbContext dbContext;
        public readonly IServiceProvider serviceProvider;

        public TraceableBeacukaiFacade(IServiceProvider serviceProvider, PurchasingDbContext dbContext)
        {
            this.serviceProvider = serviceProvider;
            this.dbContext = dbContext;
        }

        #region Masuk

        public List<GarmentExpenditureGoodViewModel> GetExpenditureGood(string RONo)
        {
            var param = new StringContent(JsonConvert.SerializeObject(RONo), Encoding.UTF8, "application/json");
            string expenditureUri = APIEndpoint.GarmentProduction + $"expenditure-goods/traceable-by-ro";

            IHttpClientService httpClient = (IHttpClientService)serviceProvider.GetService(typeof(IHttpClientService));

            var httpResponse = httpClient.SendAsync(HttpMethod.Get, expenditureUri, param).Result;
            if (httpResponse.IsSuccessStatusCode)
            {
                var content = httpResponse.Content.ReadAsStringAsync().Result;
                Dictionary<string, object> result = JsonConvert.DeserializeObject<Dictionary<string, object>>(content);

                List<GarmentExpenditureGoodViewModel> viewModel;
                if (result.GetValueOrDefault("data") == null)
                {
                    viewModel = new List<GarmentExpenditureGoodViewModel>();
                }
                else
                {
                    viewModel = JsonConvert.DeserializeObject<List<GarmentExpenditureGoodViewModel>>(result.GetValueOrDefault("data").ToString());

                }
                return viewModel;
            }
            else
            {
                return new List<GarmentExpenditureGoodViewModel>();
            }
        }

        public List<GarmentCuttingOutViewModel> GetCuttingOut(string RONo)
        {
            var param = new StringContent(JsonConvert.SerializeObject(RONo), Encoding.UTF8, "application/json");
            string expenditureUri = APIEndpoint.GarmentProduction + $"cutting-outs/for-traceable";

            IHttpClientService httpClient = (IHttpClientService)serviceProvider.GetService(typeof(IHttpClientService));

            var httpResponse = httpClient.SendAsync(HttpMethod.Get, expenditureUri, param).Result;
            if (httpResponse.IsSuccessStatusCode)
            {
                var content = httpResponse.Content.ReadAsStringAsync().Result;
                Dictionary<string, object> result = JsonConvert.DeserializeObject<Dictionary<string, object>>(content);

                List<GarmentCuttingOutViewModel> viewModel;
                if (result.GetValueOrDefault("data") == null)
                {
                    viewModel = new List<GarmentCuttingOutViewModel>();
                }
                else
                {
                    viewModel = JsonConvert.DeserializeObject<List<GarmentCuttingOutViewModel>>(result.GetValueOrDefault("data").ToString());

                }
                return viewModel;
            }
            else
            {
                return new List<GarmentCuttingOutViewModel>();
            }
        }

        public List<GarmentFinishingOutViewModel> GetFinishingOut(string RONo)
        {
            var param = new StringContent(JsonConvert.SerializeObject(RONo), Encoding.UTF8, "application/json");
            string expenditureUri = APIEndpoint.GarmentProduction + $"finishing-outs/for-traceable";

            IHttpClientService httpClient = (IHttpClientService)serviceProvider.GetService(typeof(IHttpClientService));

            var httpResponse = httpClient.SendAsync(HttpMethod.Get, expenditureUri, param).Result;
            if (httpResponse.IsSuccessStatusCode)
            {
                var content = httpResponse.Content.ReadAsStringAsync().Result;
                Dictionary<string, object> result = JsonConvert.DeserializeObject<Dictionary<string, object>>(content);

                List<GarmentFinishingOutViewModel> viewModel;
                if (result.GetValueOrDefault("data") == null)
                {
                    viewModel = new List<GarmentFinishingOutViewModel>();
                }
                else
                {
                    viewModel = JsonConvert.DeserializeObject<List<GarmentFinishingOutViewModel>>(result.GetValueOrDefault("data").ToString());

                }
                return viewModel;
            }
            else
            {
                return new List<GarmentFinishingOutViewModel>();
            }
        }

        public List<BeacukaiAddedViewModel> GetPEB(string invoice)
        {
            var param = new StringContent(JsonConvert.SerializeObject(invoice), Encoding.UTF8, "application/json");
            string shippingInvoiceUri = APIEndpoint.CustomsReport + $"customs-reports/getPEB";

            IHttpClientService httpClient = (IHttpClientService)serviceProvider.GetService(typeof(IHttpClientService));

            var httpResponse = httpClient.SendAsync(HttpMethod.Get, shippingInvoiceUri, param).Result;
            if (httpResponse.IsSuccessStatusCode)
            {
                var content = httpResponse.Content.ReadAsStringAsync().Result;
                Dictionary<string, object> result = JsonConvert.DeserializeObject<Dictionary<string, object>>(content);

                List<BeacukaiAddedViewModel> viewModel;
                if (result.GetValueOrDefault("data") == null)
                {
                    viewModel = new List<BeacukaiAddedViewModel>();
                }
                else
                {
                    viewModel = JsonConvert.DeserializeObject<List<BeacukaiAddedViewModel>>(result.GetValueOrDefault("data").ToString());

                }
                return viewModel;
            }
            else
            {
                return new List<BeacukaiAddedViewModel>();
            }
        }

        public Tuple<List<TraceableInBeacukaiViewModel>, int> GetReportTraceableIN(string filter, string tipe, string tipebc)
        {
            //var Query = GetStockQuery(tipebarang, unitcode, dateFrom, dateTo, offset);
            //Query = Query.OrderByDescending(x => x.SupplierName).ThenBy(x => x.Dono);
            List<TraceableInBeacukaiViewModel> Query = GetTraceableInQuery(filter, tipe, tipebc);
            //Query = Query.OrderBy(x => x.ItemCode).ToList();
            int TotalData = Query.Count;
            //int TotalData = Data.Count();
            return Tuple.Create(Query, TotalData);
        }

        public List<TraceableInBeacukaiViewModel> GetTraceableInQuery(string filter, string tipe, string tipebc)
        {
            var Query = tipe == "BCNo" ? (from a in (from aa in dbContext.GarmentBeacukais
                                                     where aa.BeacukaiNo == filter && aa.CustomsType == tipebc
                                                     select aa)
                                          join t in dbContext.GarmentDeliveryOrders on a.Id equals t.CustomsId
                                          join c in dbContext.GarmentDeliveryOrderItems on t.Id equals c.GarmentDOId
                                          join d in dbContext.GarmentDeliveryOrderDetails on c.Id equals d.GarmentDOItemId
                                          join e in dbContext.GarmentUnitReceiptNoteItems on d.Id equals e.DODetailId
                                          join r in dbContext.GarmentUnitReceiptNotes on e.URNId equals r.Id
                                          join f in dbContext.GarmentUnitDeliveryOrderItems on e.Id equals f.URNItemId into unitdoitem
                                          from ff in unitdoitem.DefaultIfEmpty()
                                          join g in dbContext.GarmentUnitDeliveryOrders on ff.UnitDOId equals g.Id into unitdo
                                          from gg in unitdo.DefaultIfEmpty()
                                          join h in dbContext.GarmentUnitExpenditureNoteItems on ff.Id equals h.UnitDOItemId into uenitem
                                          from hh in uenitem.DefaultIfEmpty()
                                          join i in dbContext.GarmentUnitExpenditureNotes on hh.UENId equals i.Id into uen
                                          from ii in uen.DefaultIfEmpty()
                                          where r.URNType == "PEMBELIAN"
                                          select new TraceableInBeacukaiViewModelTemp
                                          {
                                              BCDate = a.BeacukaiDate,
                                              BCNo = a.BeacukaiNo,
                                              BCType = a.CustomsType,
                                              BJQty = 0,
                                              BonNo = a.BillNo,
                                              BUK = ii != null ? ii.UENNo : "buk-",
                                              BUM = r.URNNo,
                                              EksporQty = 0,
                                              QtyBUK = hh != null ? hh.Quantity : 0,
                                              ItemCode = d.ProductCode,
                                              ItemName = d.ProductName,
                                              PO = d.POSerialNumber,
                                              ReceiptQty = Math.Round((double)(e.ReceiptQuantity * e.Conversion),2),
                                              ROJob = gg != null ? gg.RONo : "rojob-",
                                              SatuanBUK = hh != null ? hh.UomUnit : "satbuk-",
                                              SatuanReceipt = e.SmallUomUnit,
                                              SampleQty = ii != null ? (ii.ExpenditureType == "SAMPLE" ? hh.Quantity : 0) : 0
                                          }).Distinct() :
                                         (from a in (from aa in dbContext.GarmentBeacukais
                                                    select aa)
                                         join deliv in dbContext.GarmentDeliveryOrders on a.Id equals deliv.CustomsId
                                         join c in dbContext.GarmentDeliveryOrderItems on deliv.Id equals c.GarmentDOId
                                         join d in dbContext.GarmentDeliveryOrderDetails on c.Id equals d.GarmentDOItemId
                                         join e in dbContext.GarmentUnitReceiptNoteItems on d.Id equals e.DODetailId
                                         join r in dbContext.GarmentUnitReceiptNotes on e.URNId equals r.Id
                                         join f in dbContext.GarmentUnitDeliveryOrderItems on e.Id equals f.URNItemId
                                         join g in dbContext.GarmentUnitDeliveryOrders on f.UnitDOId equals g.Id
                                         join h in dbContext.GarmentUnitExpenditureNoteItems on g.Id equals h.UnitDOItemId
                                         join i in dbContext.GarmentUnitExpenditureNotes on h.UENId equals i.Id
                                         where g.RONo == filter
                                         where r.URNType == "PEMBELIAN"
                                         select new TraceableInBeacukaiViewModelTemp
                                         {
                                             BCDate = a.BeacukaiDate,
                                             BCNo = a.BeacukaiNo,
                                             BCType = a.CustomsType,
                                             BJQty = 0,
                                             BonNo = a.BillNo,
                                             BUK = i.UENNo,
                                             BUM = r.URNNo,
                                             EksporQty = 0,
                                             QtyBUK = h.Quantity,
                                             ItemCode = d.ProductCode,
                                             ItemName = d.ProductName,
                                             PO = d.POSerialNumber,
                                             ReceiptQty = Math.Round((double)(e.ReceiptQuantity * e.Conversion),2),
                                             ROJob = g.RONo,
                                             SatuanBUK = h.UomUnit,
                                             SatuanReceipt = e.SmallUomUnit,
                                             SampleQty = i.ExpenditureType == "SAMPLE" ? h.Quantity : 0
                                         }).Distinct();

            //Query = Query.GroupBy(x => new { x.BCDate, x.BCNo, x.BCType, x.BonNo, x.BUM, x.BUK, x.ItemCode, x.ItemName, x.PO, x.ROJob, x.SatuanBUK, x.SatuanReceipt }, (key, group) => new TraceableInBeacukaiViewModelTemp
            //{
            //    BCDate = key.BCDate,
            //    BCNo = key.BCNo,
            //    BCType = key.BCType,
            //    BJQty = group.Sum(x => x.BJQty),
            //    BonNo = key.BonNo,
            //    BUM = key.BUM,
            //    BUK = key.BUK,
            //    EksporQty = group.Sum(x => x.EksporQty),
            //    QtyBUK = group.Sum(x => x.QtyBUK),
            //    ItemCode = key.ItemCode,
            //    ItemName = key.ItemName,
            //    PO = key.PO,
            //    ReceiptQty = group.Sum(x => x.ReceiptQty),
            //    ROJob = key.ROJob,
            //    SatuanBUK = key.SatuanBUK,
            //    SatuanReceipt = key.SatuanReceipt
            //});

            var Query2 = Query.OrderBy(x => x.BCType).ThenBy(x => x.BCNo).ThenBy(x => x.BCDate).ThenBy(x => x.BonNo).ThenBy(x => x.PO).ThenBy(x => x.ItemCode).ThenBy(x => x.ItemName).ThenBy(x => x.ReceiptQty).ThenBy(x => x.ROJob).ThenBy(x => x.BUK).ToList();

            var ro = string.Join(",", Query2.Select(x => x.ROJob).Distinct().ToList());
            var expendituregoods = GetExpenditureGood(ro);
            var cuttingIn = GetCuttingOut(ro);
            var invoices = string.Join(",", expendituregoods.Select(x => x.Invoice).Distinct().ToList());
            var PEBs = GetPEB(invoices);
            var finouts = GetFinishingOut(ro);

            var Data2 = (from a in Query2
                         join expend in expendituregoods on a.ROJob equals expend.RONo into expendgood
                         from bb in expendgood.DefaultIfEmpty()
                         //join peb in PEBs on bb.Invoice equals peb.BonNo.Trim() into bcout
                         //from cc in bcout.DefaultIfEmpty()
                         select new 
                         {
                             BCDate = a.BCDate,
                             BCNo = a.BCNo,
                             BCType = a.BCType,
                             BJQty = a.BJQty,
                             BonNo = a.BonNo,
                             BUK = a.BUK,
                             BUM = a.BUM,
                             QtyBUK = a.QtyBUK,
                             ItemCode = a.ItemCode,
                             ItemName = a.ItemName,
                             PO = a.PO,
                             ReceiptQty = a.ReceiptQty,
                             ROJob = a.ROJob,
                             SatuanBUK = a.SatuanBUK,
                             SatuanReceipt = a.SatuanReceipt,
                             SampleQty = bb == null ? 0 : bb.ExpenditureType == "SAMPLE" ? bb.TotalQuantity : 0,
                             Invoice = bb == null ? "invo-" : bb.Invoice,
                             //PEB = cc != null ? cc.BCNo : "peb-",
                             //PEBDate = cc != null ? cc.BCDate : new DateTimeOffset(new DateTime(1970, 1, 1)),
                             //EksporQty = cc != null ? cc.Quantity : 0
                         }).OrderBy(x => x.BCType).ThenBy(x => x.BCNo).ThenBy(x => x.BCDate).ThenBy(x => x.BonNo).ThenBy(x => x.ROJob).ThenBy(x => x.PO).ThenBy(x => x.ItemCode).ThenBy(x => x.ItemName).ThenBy(x => x.ReceiptQty).ThenBy(x => x.BUK).Distinct().ToList();

            //var DataTrace = Data2.Select((a, coba) => new
            //{
            //    rown = coba + 1,
            //    BCDate = a.BCDate,
            //    BCNo = a.BCNo,
            //    BCType = a.BCType,
            //    BJQty = a.BJQty,
            //    BonNo = a.BonNo,
            //    BUK = a.BUK,
            //    BUM = a.BUM,
            //    EksporQty = a.EksporQty,
            //    QtyBUK = a.QtyBUK,
            //    ItemCode = a.ItemCode,
            //    ItemName = a.ItemName,
            //    PO = a.PO,
            //    ReceiptQty = a.ReceiptQty,
            //    ROJob = a.ROJob,
            //    SatuanBUK = a.SatuanBUK,
            //    SatuanReceipt = a.SatuanReceipt

            //    //SampleQty = bb == null ? 0 : bb.ExpenditureType == "SAMPLE" ? bb.TotalQuantity : 0,
            //    //Invoice = bb == null ? "" : bb.Invoice,

            //}).ToList();

            List<TraceableInBeacukaiViewModelTemp> traceableIn1 = new List<TraceableInBeacukaiViewModelTemp>();
            List<TraceableInBeacukaiViewModel> traceableIn2 = new List<TraceableInBeacukaiViewModel>();

            var DistinctMasukperPO = Query.Select(x => new {x.ReceiptQty, x.PO }).GroupBy(x=> new { x.PO, x.ReceiptQty }, (key,group) => new {
                key.PO,
                key.ReceiptQty
            }).ToList();
            var DistinctKeluarperPO = Query.Select(x => new {x.QtyBUK, x.PO }).GroupBy(x => new { x.PO, x.QtyBUK }, (key, group) => new
            {
                key.PO,
                key.QtyBUK

            }).ToList();

            var groupMasukperPO = DistinctMasukperPO.GroupBy(x => new { x.PO }, (key, group) => new
            {
                PO = key.PO,
                ReceiptQty = group.Sum(x => x.ReceiptQty)
            }).ToList();

            var groupKeluarperPO = DistinctKeluarperPO.GroupBy(x => new { x.PO }, (key, group) => new
            {
                PO = key.PO,
                QtyBUK = group.Sum(x => x.QtyBUK)
            }).ToList();

            //var ListQtyMasukPerPO = DistinctMasukperPO.GroupBy(x => new { x.PO, x.ReceiptQty }, (key, group) => new
            //{
            //    PO = key.PO,
            //    ReceiptQty = group.Sum(x=>x.ReceiptQty)
            //});

            //var ListQtyKeluarperPO = DistinctKeluarperPO.GroupBy(x => new { x.PO }, (key, group) => new
            //{
            //    PO = key.PO,
            //    QtyBUK = group.Sum(x=>x.QtyBUK)
            //});

            //var Sisa = (from aa in ListQtyMasukPerPO
            //            join bb in ListQtyKeluarperPO on aa.PO equals bb.PO
            //            select new
            //            {
            //                PO = aa.PO,
            //                Sisa = aa.ReceiptQty - bb.QtyBUK
            //            }).ToList();



            foreach (var i in Data2)
            {
                var cutting = cuttingIn.FirstOrDefault(x => x.RONo == i.ROJob);
                var sisa1 = groupMasukperPO.FirstOrDefault(x => x.PO == i.PO);
                var sisa2 = groupKeluarperPO.FirstOrDefault(x => x.PO == i.PO);
                var PEB = PEBs.FirstOrDefault(x => x.BonNo.Trim() == i.Invoice);
                //var PEBQty = PEBs.FirstOrDefault(x => x.BonNo.Trim() == i.Invoice);
                var subconout = finouts.FirstOrDefault(x => x.roJob == i.ROJob && x.finishingInType == "PEMBELIAN");
                var finishingout = finouts.FirstOrDefault(x => x.roJob == i.ROJob && x.finishingTo == "GUDANG JADI");


                var trace1 = new TraceableInBeacukaiViewModelTemp
                {
                    BCDate = i.BCDate,
                    BCNo = i.BCNo,
                    BCType = i.BCType,
                    BJQty = (finishingout == null ? 0 : finishingout.totalQty) + (subconout == null ? 0 : subconout.totalQty),
                    BonNo = i.BonNo,
                    BUK = i.BUK,
                    BUM = i.BUM,
                    Invoice = i.Invoice,
                    ItemCode = i.ItemCode,
                    ItemName = i.ItemName,
                    PEB = PEB != null ? PEB.BCNo : "peb-",
                    PEBDate = PEB != null ? PEB.BCDate : new DateTimeOffset(new DateTime(1970, 1, 1)),
                    PO = i.PO,
                    QtyBUK = i.QtyBUK,
                    ReceiptQty = i.ReceiptQty,
                    ROJob = i.ROJob,
                    SampleQty = 0,
                    SatuanBUK = i.SatuanBUK,
                    SatuanReceipt = i.SatuanReceipt,
                    Sisa = Math.Round(sisa1.ReceiptQty - sisa2.QtyBUK,2),
                    SubkonOutQty = subconout == null ? 0 : subconout.totalQty,
                    ProduksiQty = ((cutting != null && finishingout != null) ? cutting.TotalCuttingOutQuantity - finishingout.totalQty : 0),
                    EksporQty = PEB != null ? PEB.Quantity : 0
                };

                traceableIn1.Add(trace1);

            }

            
            //traceableIn1 = traceableIn1.OrderBy(x => x.ROJob).ThenBy(x => x.Invoice).ToList();


            //traceableIn1 = traceableIn1.OrderBy(x => x.BCType).ThenBy(x => x.BCNo).ThenBy(x => x.BCDate).ThenBy(x => x.BonNo).ThenBy(x => x.ROJob).ThenBy(x => x.PO).ThenBy(x => x.ItemCode)
            //    .ThenBy(x => x.ItemName).ThenBy(x => x.BUM).ThenBy(x => x.ReceiptQty).ThenBy(x => x.SatuanReceipt).ThenBy(x => x.BUK).ThenBy(x => x.QtyBUK).ThenBy(x => x.Sisa).ThenBy(x => x.SatuanBUK).ThenBy(x => x.ProduksiQty)
            //    .ThenBy(x => x.BJQty).ThenBy(x => x.Invoice).ToList();


            var trace21 = traceableIn1.Select((i, order) => new
            {
                rown = order + 1,
                BCDate = i.BCDate,
                BCNo = i.BCNo,
                BCType = i.BCType,
                BJQty = i.BJQty,
                BonNo = i.BonNo,
                BUK = i.BUK,
                BUM = i.BUM,
                //Invoice = i.Invoice,
                ItemCode = i.ItemCode,
                ItemName = i.ItemName,
                //PEB = PEB != null ? PEB.BCNo : "-",
                //PEBDate = new DateTimeOffset(new DateTime(1970, 1, 1)),
                PO = i.PO,
                //ProduksiQty = cutting == null ? 0 : cutting.Sum(x => x.TotalCuttingInQuantity),
                QtyBUK = i.QtyBUK,
                ReceiptQty = i.ReceiptQty,
                ROJob = i.ROJob,
                SampleQty = 0,
                SatuanBUK = i.SatuanBUK,
                SatuanReceipt = i.SatuanReceipt,
                Sisa = i.Sisa,
                SubkonOutQty = i.SubkonOutQty,
                ProduksiQty = i.ProduksiQty
            });

            var expendorder = Data2.OrderBy(x => x.ROJob).ThenBy(x=>x.Invoice).Select((a, order) => new
            {
                rown = order + 1,
                Invoice = a.Invoice,
                PEB = PEBs.FirstOrDefault(x => x.BonNo.Trim() == a.Invoice) == null ? "peb-" : PEBs.FirstOrDefault(x => x.BonNo.Trim() == a.Invoice).BCNo,
                PEBDate = PEBs.FirstOrDefault(x => x.BonNo.Trim() == a.Invoice) == null ? new DateTimeOffset(new DateTime(1970, 1, 1)) : PEBs.FirstOrDefault(x => x.BonNo.Trim() == a.Invoice).BCDate,
                EksporQty = PEBs.FirstOrDefault(x => x.BonNo.Trim() == a.Invoice) == null ? 0 : PEBs.FirstOrDefault(x => x.BonNo.Trim() == a.Invoice).Quantity
            }).ToList();

            var traceable = from a in trace21
                            join r in expendorder on a.rown equals r.rown
                            select new TraceableInBeacukaiViewModelTemp
                            {
                                BCDate = a.BCDate,
                                BCNo = a.BCNo,
                                BCType = a.BCType,
                                BJQty = a.BJQty,
                                BonNo = a.BonNo,
                                BUK = a.BUK,
                                BUM = a.BUM,
                                count = 0,
                                EksporQty = r.EksporQty,
                                //ExType = i.ExType,
                                Invoice = r.Invoice,
                                ItemCode = a.ItemCode,
                                ItemName = a.ItemName,
                                PEB = r.PEB,
                                PEBDate = r.PEBDate,
                                PO = a.PO,
                                ProduksiQty = a.ProduksiQty,
                                QtyBUK = a.QtyBUK,
                                ReceiptQty = a.ReceiptQty,
                                ROJob = a.ROJob,
                                SampleQty = a.SampleQty,
                                SatuanBUK = a.SatuanBUK,
                                SatuanReceipt = a.SatuanReceipt,
                                Sisa = a.Sisa,
                                SubkonOutQty = a.SubkonOutQty
                            };


            var traceableInBeacukaiViews = traceable.ToArray();
            var index = 0;
            foreach (TraceableInBeacukaiViewModelTemp a in traceableInBeacukaiViews)
            {
                TraceableInBeacukaiViewModelTemp dup = Array.Find(traceableInBeacukaiViews, o => o.BCType == a.BCType && o.BCNo == a.BCNo && o.BonNo == a.BonNo);
                if (dup != null)
                {
                    if (dup.count == 0)
                    {
                        index++;
                        dup.count = index;
                    }
                }
                a.count = dup.count;
            }

            foreach(var i in traceableInBeacukaiViews.ToList())
            {
                var trace1 = new TraceableInBeacukaiViewModel
                {
                    BCDate = i.BCDate,
                    BCNo = i.BCNo,
                    BCType = i.BCType,
                    BJQty = i.BJQty == 0 ? "BJQty0" : i.BJQty.ToString(),
                    BonNo = i.BonNo,
                    BUK = i.BUK,
                    BUM = i.BUM,
                    Invoice = i.Invoice,
                    ItemCode = i.ItemCode,
                    ItemName = i.ItemName,
                    PEB = i.PEB,
                    PEBDate = i.PEBDate,
                    PO = i.PO,
                    QtyBUK = i.QtyBUK == 0 ? "QtyBUK0" : i.QtyBUK.ToString(),
                    ReceiptQty = i.ReceiptQty.ToString(),
                    ROJob = i.ROJob,
                    SampleQty = i.SampleQty == 0 ? "SampleQty0" : i.SampleQty.ToString(),
                    SatuanBUK = i.SatuanBUK,
                    SatuanReceipt = i.SatuanReceipt,
                    Sisa = i.Sisa == 0 ? "Sisa0" : i.Sisa.ToString(),
                    SubkonOutQty = i.SubkonOutQty == 0 ? "SubkonOutQty0" : i.SubkonOutQty.ToString(),
                    ProduksiQty = i.ProduksiQty == 0 ? "ProduksiQty0" : i.ProduksiQty.ToString(),
                    count = i.count,
                    EksporQty = i.EksporQty == 0 ? "EksporQty0" : i.EksporQty.ToString(),
                    ExType = ""
                };

                traceableIn2.Add(trace1);
            }
                

            return traceableIn2.Distinct().ToList();

        }

        public MemoryStream GetTraceableInExcel(string filter, string tipe, string tipebc)
        {
            var Query = GetTraceableInQuery(filter, tipe, tipebc);
            //Query.OrderBy(x => x.BCNo);
            DataTable result = new DataTable();
            result.Columns.Add(new DataColumn() { ColumnName = "No", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Jenis BC", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Nomor BC", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Tanggal BC", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "No Bon", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "ROJob", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "PO", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Kode Barang", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Nama Barang", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Jumlah Terima", DataType = typeof(Double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Satuan Terima", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "No BUK", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Jumlah Keluar", DataType = typeof(Double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Sisa", DataType = typeof(Double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Satuan Keluar", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Produksi Qty", DataType = typeof(Double) });
            result.Columns.Add(new DataColumn() { ColumnName = "BJ Qty", DataType = typeof(Double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Invoice", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "BC Keluar", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Tgl BC Keluar", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Ekspor Qty", DataType = typeof(Double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Sample Qty", DataType = typeof(Double) });
            if (Query.ToArray().Count() == 0)
                result.Rows.Add("", "", "", "", "", "", "", "", 0, "", 
                    "No BUK", 0, 0, "", 0, 0, "", "", "", 0, 0); // to allow column name to be generated properly for empty data as template
            else
            {
                var docNo = Query.ToArray();
                var q = Query.ToList();
                var index = 0;
                foreach (TraceableInBeacukaiViewModel a in q)
                {
                    TraceableInBeacukaiViewModel dup = Array.Find(docNo, o => o.BCType == a.BCType && o.BCNo == a.BCNo && o.BonNo == o.BonNo);
                    if (dup != null)
                    {
                        if (dup.count == 0)
                        {
                            index++;
                            dup.count = index;
                        }
                    }
                    a.count = dup.count;
                }
                Query = q;
                //var index = 0;
                foreach (var item in Query)
                {
                    string bcdate = item.BCDate == new DateTimeOffset(new DateTime(1970, 1, 1)) ? "-" : item.BCDate.DateTime.ToString("dd MMM yyyy", new CultureInfo("id-ID"));
                    string pebdate = item.PEBDate == new DateTimeOffset(new DateTime(1970, 1, 1)) ? "-" : item.PEBDate.Value.DateTime.ToString("dd MMM yyyy", new CultureInfo("id-ID"));
                    string rojob = item.ROJob == "rojob-" ? "-" : item.ROJob;
                    string buk = item.BUK == "buk-" ? "-" : item.BUK;
                    string satuanbuk = item.SatuanBUK == "satbuk-" ? "-" : item.SatuanBUK;
                    string invoice = item.Invoice == "invo-" ? "-" : item.Invoice;
                    string PEB = item.PEB == "peb-" ? "-" : item.PEB;
                    double qtybuk = item.QtyBUK == "QtyBUK0" ? 0 : Convert.ToDouble(item.QtyBUK);
                    double sisa = item.Sisa == "Sisa0" ? 0 : Convert.ToDouble(item.Sisa);
                    double produksiQty = item.ProduksiQty == "ProduksiQty0" ? 0 : Convert.ToDouble(item.ProduksiQty);
                    double bJQty = item.BJQty == "BJQty0" ? 0 : Convert.ToDouble(item.BJQty);
                    double eksporQty = item.EksporQty == "EksporQty0" ? 0 : Convert.ToDouble(item.EksporQty);
                    double sampleQty = item.SampleQty == "SampleQty0" ? 0 : Convert.ToDouble(item.SampleQty);
                    result.Rows.Add(item.count, item.BCType, item.BCNo, bcdate, item.BonNo, rojob, 
                        item.PO, item.ItemCode, item.ItemName, 
                        item.ReceiptQty, 
                        item.SatuanReceipt, buk,
                        qtybuk, sisa, satuanbuk, produksiQty, bJQty,
                        invoice, PEB, pebdate, eksporQty, sampleQty);

                }
            }
            //return Excel.CreateExcel(new List<KeyValuePair<DataTable, string>>() { new KeyValuePair<DataTable, string>(result, "Territory") }, true);
            ExcelPackage package = new ExcelPackage();
            bool styling = true;

            foreach (KeyValuePair<DataTable, String> item in new List<KeyValuePair<DataTable, string>>() { new KeyValuePair<DataTable, string>(result, "Territory") })
            {
                var sheet = package.Workbook.Worksheets.Add(item.Value);
                sheet.Cells["A1"].LoadFromDataTable(item.Key, true, (styling == true) ? OfficeOpenXml.Table.TableStyles.Light16 : OfficeOpenXml.Table.TableStyles.None);
                //sheet.Cells["C1:D1"].Merge = true;
                //sheet.Cells["C1:D1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                //sheet.Cells["E1:F1"].Merge = true;
                //sheet.Cells["C1:D1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                Dictionary<string, int> bcnospan = new Dictionary<string, int>();
                Dictionary<string, int> bonspan = new Dictionary<string, int>();
                Dictionary<string, int> rojobspan = new Dictionary<string, int>();
                Dictionary<string, int> pospan = new Dictionary<string, int>();
                Dictionary<string, int> itemcodespan = new Dictionary<string, int>();
                Dictionary<string, int> itemnamespan = new Dictionary<string, int>();
                Dictionary<string, int> qtyreceiptspan = new Dictionary<string, int>();
                Dictionary<string, int> satuanreceiptspan = new Dictionary<string, int>();
                Dictionary<string, int> nobukspan = new Dictionary<string, int>();
                Dictionary<string, int> sisaspan = new Dictionary<string, int>();
                Dictionary<string, int> satuanbukspan = new Dictionary<string, int>();
                Dictionary<string, int> produksiQtyspan = new Dictionary<string, int>();
                Dictionary<string, int> bjquantityspan = new Dictionary<string, int>();
                Dictionary<string, int> invoicespan = new Dictionary<string, int>();
                Dictionary<string, int> pebnospan = new Dictionary<string, int>();
                Dictionary<string, int> pebdatespan = new Dictionary<string, int>();
                Dictionary<string, int> eksporqtyspan = new Dictionary<string, int>();
                Dictionary<string, int> samppleqtyspan = new Dictionary<string, int>();
                var docNo = Query.ToArray();
                int value;
                foreach (var a in Query)
                {
                    //FactBeacukaiViewModel dup = Array.Find(docNo, o => o.BCType == a.BCType && o.BCNo == a.BCNo);
                    if (bcnospan.TryGetValue(a.BCType + a.BCNo + a.BCDate, out value))
                    {
                        bcnospan[a.BCType + a.BCNo + a.BCDate]++;
                    }
                    else
                    {
                        bcnospan[a.BCType + a.BCNo + a.BCDate] = 1;
                    }

                    //FactBeacukaiViewModel dup1 = Array.Find(docNo, o => o.BCType == a.BCType);
                    if (bonspan.TryGetValue(a.BonNo, out value))
                    {
                        bonspan[a.BonNo]++;
                    }
                    else
                    {
                        bonspan[a.BonNo] = 1;
                    }

                    if (rojobspan.TryGetValue(a.ROJob + a.PO, out value))
                    {
                        rojobspan[a.ROJob + a.PO]++;
                    }
                    else
                    {
                        rojobspan[a.ROJob + a.PO] = 1;
                    }
                    if (itemcodespan.TryGetValue(a.ItemCode, out value))
                    {
                        itemcodespan[a.ItemCode]++;
                    }
                    else
                    {
                        itemcodespan[a.ItemCode] = 1;
                    }
                    if (itemnamespan.TryGetValue(a.ItemName, out value))
                    {
                        itemnamespan[a.ItemName]++;
                    }
                    else
                    {
                        itemnamespan[a.ItemName] = 1;
                    }

                    if (qtyreceiptspan.TryGetValue(a.ReceiptQty + "bum" + a.PO + a.ROJob, out value))
                    {
                        qtyreceiptspan[a.ReceiptQty + "bum" + a.PO + a.ROJob]++;
                    }
                    else
                    {
                        qtyreceiptspan[a.ReceiptQty + "bum" + a.PO + a.ROJob] = 1;
                    }
                    if (satuanreceiptspan.TryGetValue(a.SatuanReceipt, out value))
                    {
                        satuanreceiptspan[a.SatuanReceipt]++;
                    }
                    else
                    {
                        satuanreceiptspan[a.SatuanReceipt] = 1;
                    }
                    if (nobukspan.TryGetValue(a.BUK + a.QtyBUK, out value))
                    {
                        nobukspan[a.BUK + a.QtyBUK]++;
                    }
                    else
                    {
                        nobukspan[a.BUK + a.QtyBUK] = 1;
                    }
                    if (sisaspan.TryGetValue(a.Sisa.ToString() + "sisa" + a.PO + a.ROJob, out value))
                    {
                        sisaspan[a.Sisa.ToString() + "sisa" + a.PO + a.ROJob]++;
                    }
                    else
                    {
                        sisaspan[a.Sisa.ToString() + "sisa" + a.PO + a.ROJob] = 1;
                    }
                    if (satuanbukspan.TryGetValue(a.SatuanBUK, out value))
                    {
                        satuanbukspan[a.SatuanBUK]++;
                    }
                    else
                    {
                        satuanbukspan[a.SatuanBUK] = 1;
                    }
                    if (produksiQtyspan.TryGetValue(a.ProduksiQty.ToString() + a.ROJob, out value))
                    {
                        produksiQtyspan[a.ProduksiQty.ToString() + a.ROJob]++;
                    }
                    else
                    {
                        produksiQtyspan[a.ProduksiQty.ToString() + a.ROJob] = 1;
                    }
                    if (bjquantityspan.TryGetValue(a.BJQty.ToString() + a.ROJob, out value))
                    {
                        bjquantityspan[a.BJQty.ToString() + a.ROJob]++;
                    }
                    else
                    {
                        bjquantityspan[a.BJQty.ToString() + a.ROJob] = 1;
                    }
                    if (invoicespan.TryGetValue(a.Invoice + a.ROJob, out value))
                    {
                        invoicespan[a.Invoice + a.ROJob]++;
                    }
                    else
                    {
                        invoicespan[a.Invoice + a.ROJob] = 1;
                    }
                    if (pebnospan.TryGetValue(a.PEB + a.ROJob, out value))
                    {
                        pebnospan[a.PEB + a.ROJob]++;
                    }
                    else
                    {
                        pebnospan[a.PEB + a.ROJob] = 1;
                    }

                    if (pebdatespan.TryGetValue(a.PEBDate.ToString() + a.ROJob, out value))
                    {
                        pebdatespan[a.PEBDate.ToString() + a.ROJob]++;
                    }
                    else
                    {
                        pebdatespan[a.PEBDate.ToString() + a.ROJob] = 1;
                    }

                    if (eksporqtyspan.TryGetValue(a.EksporQty.ToString() + a.ROJob, out value))
                    {
                        eksporqtyspan[a.EksporQty.ToString() + a.ROJob]++;
                    }
                    else
                    {
                        eksporqtyspan[a.EksporQty.ToString() + a.ROJob] = 1;
                    }

                    if (samppleqtyspan.TryGetValue(a.SampleQty.ToString() + a.ROJob, out value))
                    {
                        samppleqtyspan[a.SampleQty.ToString() + a.ROJob]++;
                    }
                    else
                    {
                        samppleqtyspan[a.SampleQty.ToString() + a.ROJob] = 1;
                    }

                    if (pospan.TryGetValue(a.PO + a.ROJob, out value))
                    {
                        pospan[a.PO + a.ROJob]++;
                    }
                    else
                    {
                        pospan[a.PO + a.ROJob] = 1;
                    }
                }

                int index = 2;
                foreach (KeyValuePair<string, int> b in bcnospan)
                {
                    sheet.Cells["A" + index + ":A" + (index + b.Value - 1)].Merge = true;
                    sheet.Cells["A" + index + ":A" + (index + b.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                    sheet.Cells["B" + index + ":B" + (index + b.Value - 1)].Merge = true;
                    sheet.Cells["B" + index + ":B" + (index + b.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                    sheet.Cells["C" + index + ":C" + (index + b.Value - 1)].Merge = true;
                    sheet.Cells["C" + index + ":C" + (index + b.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                    sheet.Cells["D" + index + ":D" + (index + b.Value - 1)].Merge = true;
                    sheet.Cells["D" + index + ":D" + (index + b.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                    index += b.Value;
                }

                index = 2;
                foreach (KeyValuePair<string, int> c in bonspan)
                {
                    sheet.Cells["E" + index + ":E" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["E" + index + ":E" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in rojobspan)
                {
                    sheet.Cells["F" + index + ":F" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["F" + index + ":F" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in pospan)
                {
                    sheet.Cells["G" + index + ":G" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["G" + index + ":G" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in itemcodespan)
                {
                    sheet.Cells["H" + index + ":H" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["H" + index + ":H" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in itemnamespan)
                {
                    sheet.Cells["I" + index + ":I" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["I" + index + ":I" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in qtyreceiptspan)
                {
                    sheet.Cells["J" + index + ":J" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["J" + index + ":J" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in satuanreceiptspan)
                {
                    sheet.Cells["K" + index + ":K" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["K" + index + ":K" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in nobukspan)
                {
                    sheet.Cells["L" + index + ":L" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["L" + index + ":L" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                    sheet.Cells["M" + index + ":M" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["M" + index + ":M" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in sisaspan)
                {
                    sheet.Cells["N" + index + ":N" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["N" + index + ":N" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in satuanbukspan)
                {
                    sheet.Cells["O" + index + ":O" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["O" + index + ":O" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in produksiQtyspan)
                {
                    sheet.Cells["P" + index + ":P" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["P" + index + ":P" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in bjquantityspan)
                {
                    sheet.Cells["Q" + index + ":Q" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["Q" + index + ":Q" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in invoicespan)
                {
                    sheet.Cells["R" + index + ":R" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["R" + index + ":R" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in pebnospan)
                {
                    sheet.Cells["S" + index + ":S" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["S" + index + ":S" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in pebdatespan)
                {
                    sheet.Cells["T" + index + ":T" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["T" + index + ":T" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in eksporqtyspan)
                {
                    sheet.Cells["U" + index + ":U" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["U" + index + ":U" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                index = 2;
                foreach (KeyValuePair<string, int> c in samppleqtyspan)
                {
                    sheet.Cells["V" + index + ":V" + (index + c.Value - 1)].Merge = true;
                    sheet.Cells["V" + index + ":V" + (index + c.Value - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                    index += c.Value;
                }
                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
            }
            MemoryStream stream = new MemoryStream();
            package.SaveAs(stream);
            return stream;
            //return Excel.CreateExcel(new List<KeyValuePair<DataTable, string>>() { new KeyValuePair<DataTable, string>(result, "Territory") }, true);
        }

        #endregion
        #region Keluar
        private List<GarmentExpenditureGoodViewModel> GetRono(int page, int size, string order, string filter)
        {
            IHttpClientService httpClient = (IHttpClientService)serviceProvider.GetService(typeof(IHttpClientService));

            string shippingInvoiceUri = APIEndpoint.GarmentProduction + $"expenditure-goods";
            string queryUri = "?page=" + page + "&size=" + size + "&order=" + order + "&filter=" + filter;
            string uri = shippingInvoiceUri + queryUri;

            var httpResponse = httpClient.GetAsync(uri).Result;
            if (httpResponse.IsSuccessStatusCode)
            {
                var content = httpResponse.Content.ReadAsStringAsync().Result;
                Dictionary<string, object> result = JsonConvert.DeserializeObject<Dictionary<string, object>>(content);

                List<GarmentExpenditureGoodViewModel> viewModel;
                if (result.GetValueOrDefault("data") == null)
                {
                    viewModel = new List<GarmentExpenditureGoodViewModel>();
                }
                else
                {
                    viewModel = JsonConvert.DeserializeObject<List<GarmentExpenditureGoodViewModel>>(result.GetValueOrDefault("data").ToString());

                }
                return viewModel;
            }
            else
            {
                return new List<GarmentExpenditureGoodViewModel>();
            }

            //if (httpClient != null)
            //{
            //    var garmentSupplierUri = APIEndpoint.GarmentProduction + $"expenditure-goods";
            //    string queryUri = "?page=" + page + "&size=" + size + "&order=" + order + "&filter=" + filter;
            //    string uri = garmentSupplierUri + queryUri;
            //    var response = httpClient.GetAsync($"{uri}").Result.Content.ReadAsStringAsync();
            //    Dictionary<string, object> result = JsonConvert.DeserializeObject<Dictionary<string, object>>(response.Result);
            //    List<GarmentExpenditureGoodViewModel> viewModel = JsonConvert.DeserializeObject<List<GarmentExpenditureGoodViewModel>>(result.GetValueOrDefault("data").ToString());
            //    return viewModel;
            //}
            //else
            //{
            //    List<GarmentExpenditureGoodViewModel> viewModel = new List<GarmentExpenditureGoodViewModel>();
            //    return viewModel;
            //}
        }

        public List<BeacukaiAddedViewModel> GetPEBbyBCNo(string bcno)
        {
            //var param = new StringContent(JsonConvert.SerializeObject(invoice), Encoding.UTF8, "application/json");
            string shippingInvoiceUri = APIEndpoint.CustomsReport + $"customs-reports/getPEB/byBCNo?bcno={bcno}";

            IHttpClientService httpClient = (IHttpClientService)serviceProvider.GetService(typeof(IHttpClientService));

            var httpResponse = httpClient.GetAsync(shippingInvoiceUri).Result;
            if (httpResponse.IsSuccessStatusCode)
            {
                var content = httpResponse.Content.ReadAsStringAsync().Result;
                Dictionary<string, object> result = JsonConvert.DeserializeObject<Dictionary<string, object>>(content);

                List<BeacukaiAddedViewModel> viewModel;
                if (result.GetValueOrDefault("data") == null)
                {
                    viewModel = new List<BeacukaiAddedViewModel>();
                }
                else
                {
                    viewModel = JsonConvert.DeserializeObject<List<BeacukaiAddedViewModel>>(result.GetValueOrDefault("data").ToString());

                }
                return viewModel;
            }
            else
            {
                return new List<BeacukaiAddedViewModel>();
            }
        }

        public List<GarmentPreparingViewModel> GetPreparingByRo(string ro)
        {
            var param = new StringContent(JsonConvert.SerializeObject(ro), Encoding.UTF8, "application/json");
            string shippingInvoiceUri = APIEndpoint.GarmentProduction + $"preparings/byRONO?RO={ro}";

            IHttpClientService httpClient = (IHttpClientService)serviceProvider.GetService(typeof(IHttpClientService));

            var httpResponse = httpClient.SendAsync(HttpMethod.Get,shippingInvoiceUri,param).Result;
            if (httpResponse.IsSuccessStatusCode)
            {
                var content = httpResponse.Content.ReadAsStringAsync().Result;
                Dictionary<string, object> result = JsonConvert.DeserializeObject<Dictionary<string, object>>(content);

                List<GarmentPreparingViewModel> viewModel;
                if (result.GetValueOrDefault("data") == null)
                {
                    viewModel = new List<GarmentPreparingViewModel>();
                }
                else
                {
                    viewModel = JsonConvert.DeserializeObject<List<GarmentPreparingViewModel>>(result.GetValueOrDefault("data").ToString());

                }
                return viewModel;
            }
            else
            {
                return new List<GarmentPreparingViewModel>();
            }
        }

        public List<TraceableOutBeacukaiViewModel> getQueryTraceableOut(string bcno)
        {

            var PEB = GetPEBbyBCNo(bcno.Trim());

            var filterexpend = new
            {
                invoice = PEB.Select(x=>x.BonNo.Trim()).FirstOrDefault()
            };

            var expend = GetRono(1, int.MaxValue, "{}", JsonConvert.SerializeObject(filterexpend));

            var Query = (from a in PEB
                         join b in expend on a.BonNo.Trim() equals b.Invoice.Trim()
                         select new TraceableOutBeacukaiViewModel
                         {
                             BCDate = a.BCDate,
                             BCNo = a.BCNo,
                             BCType = a.BCType,
                             BuyerName = b.Buyer.Name,
                             ComodityName = b.Comodity.Name,
                             ExpenditureDate = b.ExpenditureDate,
                             ExpenditureGoodId = b.ExpenditureGoodNo,
                             ExpenditureNo = b.Invoice,
                             Qty = b.TotalQuantity,
                             RO = b.RONo,
                             UnitQtyName = "PCS"

                         }).GroupBy(x => new { x.ExpenditureDate, x.ExpenditureGoodId, x.ComodityName, x.ExpenditureNo, x.BuyerName, x.BCType, x.BCNo, x.BCDate, x.RO, x.UnitQtyName }, (key, group) => new TraceableOutBeacukaiViewModel
                         {
                             BCDate = key.BCDate,
                             BCNo = key.BCNo,
                             BCType = key.BCType,
                             BuyerName = key.BuyerName,
                             ComodityName = key.ComodityName,
                             ExpenditureDate = key.ExpenditureDate,
                             ExpenditureGoodId = key.ExpenditureGoodId,
                             ExpenditureNo = key.ExpenditureNo,
                             Qty = group.Sum(x => x.Qty),
                             RO = key.RO,
                             UnitQtyName = key.UnitQtyName

                         }).ToList();

            var RO = string.Join(",", Query.Select(x => x.RO).Distinct().ToList());
            var listRo = Query.Select(x => x.RO).Distinct().ToList();

            var expend2 = GetPreparingByRo(RO);

            var rinciandetil = (from a in dbContext.GarmentUnitDeliveryOrderItems
                         join b in dbContext.GarmentUnitDeliveryOrders on a.UnitDOId equals b.Id
                         join c in dbContext.GarmentUnitExpenditureNoteItems on a.Id equals c.UnitDOItemId
                         join d in dbContext.GarmentUnitReceiptNoteItems on a.URNItemId equals d.Id
                         join e in dbContext.GarmentDeliveryOrderDetails on d.DODetailId equals e.Id
                         join f in dbContext.GarmentDeliveryOrderItems on e.GarmentDOItemId equals f.Id
                         join g in dbContext.GarmentDeliveryOrders on f.GarmentDOId equals g.Id
                         join h in dbContext.GarmentBeacukais on g.CustomsId equals h.Id
                         where listRo.Contains(b.RONo)
                         select new TraceableOutBeacukaiDetailViewModel
                         {
                             BCDate = h.BeacukaiDate.DateTime,
                             BCNo = h.BeacukaiNo,
                             BCType = h.CustomsType,
                             DestinationJob = b.RONo,
                             ItemCode = d.ProductCode,
                             ItemName = d.ProductName,
                             SmallestQuantity = c.Quantity,
                             UnitQtyName = c.UomUnit
                         }).GroupBy(x => new { x.BCDate, x.BCNo, x.BCType, x.DestinationJob, x.ItemCode, x.ItemName, x.UnitQtyName }, (key, group) => new TraceableOutBeacukaiDetailViewModel
                         {
                             BCDate = key.BCDate,
                             BCNo = key.BCNo,
                             BCType = key.BCType,
                             DestinationJob = key.DestinationJob,
                             ItemCode = key.ItemCode,
                             ItemName = key.ItemName,
                             SmallestQuantity = group.Sum(x => x.SmallestQuantity),
                             UnitQtyName = key.UnitQtyName

                         }).ToList();

            foreach(var g in rinciandetil)
            {
                var QtyExpend = expend2.FirstOrDefault(x => x.RONo == g.DestinationJob && x.ProductCode == g.ItemCode);

                g.SmallestQuantity = QtyExpend == null ? g.SmallestQuantity : QtyExpend.Quantity;
            }

            foreach (var i in Query)
            {
                var rinci = rinciandetil.Where(x => x.DestinationJob == i.RO).ToList();

                i.rincian = rinci;
                
            }


            return Query;
        }

        //public List<TraceableOutBeacukaiDetailViewModel> getQueryDetail(string RO)
        //{
        //    var expend = GetPreparingByRo(RO);
        //    //var expenditems = new List<GarmentPreparingItemViewModel>();
        //    //foreach(var i in expend)
        //    //{
        //    //    foreach(var n in i.Items)
        //    //    {
        //    //        expenditems.Add(new GarmentPreparingItemViewModel
        //    //        {
        //    //            BasicPrice = n.BasicPrice,
        //    //            DesignColor = n.DesignColor,
        //    //            FabricType = n.FabricType,
        //    //            Id = n.Id,
        //    //            LastModifiedBy = n.LastModifiedBy,
        //    //            LastModifiedDate = n.LastModifiedDate,
        //    //            Quantity = n.Quantity,
        //    //            RemainingQuantity = n.RemainingQuantity,
        //    //            ROSource = n.ROSource,
        //    //            UENItemId = n.UENItemId
        //    //        });
        //    //    }
        //    //}

        //    //var Query = (from a in dbContext.GarmentUnitDeliveryOrderItems
        //    //             join b in dbContext.GarmentUnitDeliveryOrders on a.UnitDOId equals b.Id
        //    //             join c in dbContext.GarmentUnitExpenditureNoteItems on a.Id equals c.UnitDOItemId
        //    //             join d in dbContext.GarmentUnitReceiptNoteItems on a.URNItemId equals d.Id
        //    //             join e in dbContext.GarmentDeliveryOrderDetails on d.DODetailId equals e.Id
        //    //             join f in dbContext.GarmentDeliveryOrderItems on e.GarmentDOItemId equals f.Id
        //    //             join g in dbContext.GarmentDeliveryOrders on f.GarmentDOId equals g.Id
        //    //             join h in dbContext.GarmentBeacukais on g.CustomsId equals h.Id
        //    //             where b.RONo == (string.IsNullOrEmpty(RO) ? b.RONo : RO)
        //    //             select new TraceableOutBeacukaiDetailViewModel
        //    //             {
        //    //                 BCDate = h.BeacukaiDate.DateTime,
        //    //                 BCNo = h.BeacukaiNo,
        //    //                 BCType = h.CustomsType,
        //    //                 DestinationJob = b.RONo,
        //    //                 ItemCode = d.ProductCode,
        //    //                 ItemName = d.ProductName,
        //    //                 SmallestQuantity = c.Quantity,
        //    //                 UnitQtyName = c.UomUnit
        //    //             }).GroupBy(x => new { x.BCDate, x.BCNo, x.BCType, x.DestinationJob, x.ItemCode, x.ItemName, x.UnitQtyName }, (key, group) => new TraceableOutBeacukaiDetailViewModel
        //    //             {
        //    //                 BCDate = key.BCDate,
        //    //                 BCNo = key.BCNo,
        //    //                 BCType = key.BCType,
        //    //                 DestinationJob = key.DestinationJob,
        //    //                 ItemCode = key.ItemCode,
        //    //                 ItemName = key.ItemName,
        //    //                 SmallestQuantity = group.Sum(x => x.SmallestQuantity),
        //    //                 UnitQtyName = key.UnitQtyName

        //    //             }).ToList();

        //    var Query = (from a in expend
        //                 join c in dbContext.GarmentUnitExpenditureNoteItems on a.UENItemId equals c.Id
        //                 join d in dbContext.GarmentUnitReceiptNoteItems on c.URNItemId equals d.Id
        //                 join e in dbContext.GarmentDeliveryOrderDetails on d.DODetailId equals e.Id
        //                 join f in dbContext.GarmentDeliveryOrderItems on e.GarmentDOItemId equals f.Id
        //                 join g in dbContext.GarmentDeliveryOrders on f.GarmentDOId equals g.Id
        //                 join h in dbContext.GarmentBeacukais on g.CustomsId equals h.Id
        //                 //where b.RONo == (string.IsNullOrEmpty(RO) ? b.RONo : RO)
        //                 select new TraceableOutBeacukaiDetailViewModel
        //                 {
        //                     BCDate = h.BeacukaiDate.DateTime,
        //                     BCNo = h.BeacukaiNo,
        //                     BCType = h.CustomsType,
        //                     DestinationJob = RO,
        //                     ItemCode = d.ProductCode,
        //                     ItemName = d.ProductName,
        //                     SmallestQuantity = a.Quantity - a.RemainingQuantity,
        //                     UnitQtyName = c.UomUnit
        //                 }).GroupBy(x => new { x.BCDate, x.BCNo, x.BCType, x.DestinationJob, x.ItemCode, x.ItemName, x.UnitQtyName }, (key, group) => new TraceableOutBeacukaiDetailViewModel
        //                 {
        //                     BCDate = key.BCDate,
        //                     BCNo = key.BCNo,
        //                     BCType = key.BCType,
        //                     DestinationJob = key.DestinationJob,
        //                     ItemCode = key.ItemCode,
        //                     ItemName = key.ItemName,
        //                     SmallestQuantity = group.Sum(x => x.SmallestQuantity),
        //                     UnitQtyName = key.UnitQtyName

        //                 }).ToList();



        //    return Query;
        //}


        public MemoryStream GetExceltraceOut(string bcno)
        {
            var index2 = 0;
            var query = getQueryTraceableOut(bcno);
            var satuan = "-";

            DataTable result = new DataTable();

            result.Columns.Add(new DataColumn() { ColumnName = "no", DataType = typeof(string) });
            result.Columns.Add(new DataColumn() { ColumnName = "tanggal keluar", DataType = typeof(string) });
            result.Columns.Add(new DataColumn() { ColumnName = "no. bon", DataType = typeof(string) });
            result.Columns.Add(new DataColumn() { ColumnName = "nama barang", DataType = typeof(string) });
            result.Columns.Add(new DataColumn() { ColumnName = "jumlah barang", DataType = typeof(double) });
            result.Columns.Add(new DataColumn() { ColumnName = "satuan", DataType = typeof(string) });
            result.Columns.Add(new DataColumn() { ColumnName = "no invoice", DataType = typeof(string) });
            result.Columns.Add(new DataColumn() { ColumnName = "buyer", DataType = typeof(string) });
            result.Columns.Add(new DataColumn() { ColumnName = "jenis", DataType = typeof(string) });
            result.Columns.Add(new DataColumn() { ColumnName = "nomor", DataType = typeof(string) });
            result.Columns.Add(new DataColumn() { ColumnName = "tanggal", DataType = typeof(string) });
            result.Columns.Add(new DataColumn() { ColumnName = "RO", DataType = typeof(string) });

            DataTable result2 = new DataTable();
            result2.Columns.Add(new DataColumn() { ColumnName = "Nomor22", DataType = typeof(String) });
            result2.Columns.Add(new DataColumn() { ColumnName = "No RO", DataType = typeof(String) });
            result2.Columns.Add(new DataColumn() { ColumnName = "Kode Barang", DataType = typeof(String) });
            result2.Columns.Add(new DataColumn() { ColumnName = "Nama Barang 2", DataType = typeof(String) });
            result2.Columns.Add(new DataColumn() { ColumnName = "Jumlah Pemakaian", DataType = typeof(Double) });
            result2.Columns.Add(new DataColumn() { ColumnName = "Satuan 2", DataType = typeof(String) });
            result2.Columns.Add(new DataColumn() { ColumnName = "jumlah budget", DataType = typeof(String) });
            result2.Columns.Add(new DataColumn() { ColumnName = "Jenis BC", DataType = typeof(String) });
            result2.Columns.Add(new DataColumn() { ColumnName = "No.", DataType = typeof(String) });
            result2.Columns.Add(new DataColumn() { ColumnName = "Tanggal 2", DataType = typeof(String) });


            //if (query.ToArray().Count() == 0)
            //    result.Rows.Add("", "", "", "", 0, "", "", "", "", "", ""); // to allow column name to be generated properly for empty data as template
            //else
            //{

                var index = 0;
                foreach (var item in query)
                {
                    index++;
                    string ExpenditureDate = item.ExpenditureDate == new DateTimeOffset(new DateTime(1970, 1, 1)) ? "-" : Convert.ToDateTime(item.ExpenditureDate).ToString("dd MMM yyyy", new CultureInfo("id-ID"));
                    string BCDate = item.BCDate == new DateTimeOffset(new DateTime(1970, 1, 1)) ? "-" : Convert.ToDateTime(item.BCDate).ToString("dd MMM yyyy", new CultureInfo("id-ID"));
                    result.Rows.Add(index, ExpenditureDate, item.ExpenditureGoodId, item.ComodityName, item.Qty, item.UnitQtyName, item.ExpenditureNo, item.BuyerName, item.BCType, item.BCNo, BCDate, item.RO);
                    
                    foreach(var detail in item.rincian)
                    {
                        index2++;
                        string BCDate2 = detail.BCDate == new DateTimeOffset(new DateTime(1970, 1, 1)) ? "-" : Convert.ToDateTime(item.BCDate).ToString("dd MMM yyyy", new CultureInfo("id-ID"));
                        result2.Rows.Add(index2, detail.DestinationJob, detail.ItemCode, detail.ItemName, detail.SmallestQuantity, detail.UnitQtyName, detail.BCType, detail.BCNo, BCDate2);

                    }

                }


            //}


            //foreach (var detail in query)
            //{
            //    //var querydetail = getQueryDetail(detail.RO);




            //    //if (querydetail.ToArray().Count() == 0)
            //    //    result2.Rows.Add("", "", "", "", 0, "", "", "", ""); // to allow column name to be generated properly for empty data as template
            //    //else
            //    //{


            //        foreach (var item in querydetail)
            //        {
            //            index2++;
            //            string BCDate = item.BCDate == new DateTimeOffset(new DateTime(1970, 1, 1)) ? "-" : Convert.ToDateTime(item.BCDate).ToString("dd MMM yyyy", new CultureInfo("id-ID"));
            //            result2.Rows.Add(index2, item.DestinationJob, item.ItemCode, item.ItemName, item.SmallestQuantity, item.UnitQtyName, item.BCType, item.BCNo, BCDate);

            //        }
            //    //}

            //}


            ExcelPackage package = new ExcelPackage();



            var sheet = package.Workbook.Worksheets.Add("Data");


            var Tittle = new string[] { "Monitoring Pengeluaran Hasil Produksi" };
            var headers = new string[] { "No", "Tanggal Keluar", "No BON", "Nama Barang", "Jumlah Barang", "Satuan", "No. Invoice", "Buyer", "Dokumen", "Dokumen1", "Dokumen2", "RO" };
            var subHeaders = new string[] { "Jenis", "Nomor", "Tanggal" };

            //for (int i = 0; i < headers.Length; i++)
            //{
            //    result.Columns.Add(new DataColumn() { ColumnName = headers[i], DataType = typeof(string) });
            //}



            sheet.Cells["A5"].LoadFromDataTable(result, false, OfficeOpenXml.Table.TableStyles.Light16);

            sheet.Cells["A2"].Value = Tittle[0];
            sheet.Cells["A2:L2"].Merge = true;

            sheet.Cells["I3"].Value = headers[8];
            sheet.Cells["I3:K3"].Merge = true;

            foreach (var i in Enumerable.Range(0, 8))
            {
                var col = (char)('A' + i);
                sheet.Cells[$"{col}3"].Value = headers[i];
                sheet.Cells[$"{col}3:{col}4"].Merge = true;
            }

            foreach (var i in Enumerable.Range(0, 3))
            {
                var col = (char)('I' + i);
                sheet.Cells[$"{col}4"].Value = subHeaders[i];
            }

            foreach (var i in Enumerable.Range(0, 1))
            {
                var col = (char)('L' + i);
                sheet.Cells[$"{col}3"].Value = headers[i + 11];
                sheet.Cells[$"{col}3:{col}4"].Merge = true;
            }
            sheet.Cells["A1:L4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["A1:L4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells["A1:L4"].Style.Font.Bold = true;
            //-----------


            var countdata = query.Count();

            sheet.Cells[$"A{countdata + 11}"].LoadFromDataTable(result2, false, OfficeOpenXml.Table.TableStyles.Light16);

            var Tittle1 = new string[] { "PERINCIAN PEMAKAIAN BAHAN BAKU DAN BAHAN PENOLONG" };
            var headers1 = new string[] { "No", "No. RO", "Kode Barang", "Nama Barang", "Jumlah Pemakaian", "Satuan", "Dokumen", "Dokumen1", "Dokumen2" };
            var subHeaders1 = new string[] { "Jenis", "Nomor", "Tanggal" };

            sheet.Cells[$"A{countdata + 8}"].Value = Tittle1[0];
            sheet.Cells[$"A{countdata + 8}:I{countdata + 8}"].Merge = true;

            sheet.Cells[$"G{countdata + 9}"].Value = headers1[6];
            sheet.Cells[$"G{countdata + 9}:I{countdata + 9}"].Merge = true;

            foreach (var i in Enumerable.Range(0, 6))
            {
                var col = (char)('A' + i);
                sheet.Cells[$"{col}{countdata + 9}"].Value = headers1[i];
                sheet.Cells[$"{col}{countdata + 9}:{col}{countdata + 10}"].Merge = true;
            }

            foreach (var i in Enumerable.Range(0, 3))
            {
                var col = (char)('G' + i);
                sheet.Cells[$"{col}{countdata + 10}"].Value = subHeaders1[i];
            }


            sheet.Cells[$"A{countdata + 8}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            sheet.Cells[$"A{countdata + 8}:I{countdata + 10}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[$"A{countdata + 8}:I{countdata + 10}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells[$"A{countdata + 8}:I{countdata + 10}"].Style.Font.Bold = true;

            var widths = new int[] { 5, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20 };
            foreach (var i in Enumerable.Range(0, headers.Length))
            {
                sheet.Column(i + 1).Width = widths[i];
            }

            MemoryStream stream = new MemoryStream();
            package.SaveAs(stream);
            return stream;
        }

        #endregion
    }
}
