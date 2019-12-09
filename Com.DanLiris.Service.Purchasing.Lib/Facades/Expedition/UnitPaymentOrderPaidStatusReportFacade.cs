﻿using Com.DanLiris.Service.Purchasing.Lib.Helpers;
using Com.DanLiris.Service.Purchasing.Lib.Helpers.ReadResponse;
using Com.DanLiris.Service.Purchasing.Lib.Interfaces;
using Com.DanLiris.Service.Purchasing.Lib.Models.Expedition;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.Expedition;
using Com.Moonlay.NetCore.Lib;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using Com.DanLiris.Service.Purchasing.Lib.Enums;
using System.Text;

namespace Com.DanLiris.Service.Purchasing.Lib.Facades.Expedition
{
    public class UnitPaymentOrderPaidStatusReportFacade : IUnitPaymentOrderPaidStatusReportFacade
    {
        private readonly PurchasingDbContext dbContext;
        private readonly DbSet<PurchasingDocumentExpedition> dbSet;
        public UnitPaymentOrderPaidStatusReportFacade(PurchasingDbContext dbContext)
        {
            this.dbContext = dbContext;
            this.dbSet = this.dbContext.Set<PurchasingDocumentExpedition>();
        }



        public IQueryable<UnitPaymentOrderPaidStatusViewModel> GetQuery(string UnitPaymentOrderNo, string SupplierCode, string DivisionCode, string Status, DateTimeOffset? DateFromDue, DateTimeOffset? DateToDue, DateTimeOffset? DateFrom, DateTimeOffset? DateTo, int Offset)
        {

            DateTimeOffset dateFrom = DateFrom == null ? new DateTime(1970, 1, 1) : (DateTimeOffset)DateFrom;
            DateTimeOffset dateTo = DateTo == null ? new DateTime(2100, 1, 1) : (DateTimeOffset)DateTo;

            DateTimeOffset dateFromDue = DateFromDue == null ? new DateTime(1970, 1, 1) : (DateTimeOffset)DateFromDue;
            DateTimeOffset dateToDue = DateToDue == null ? new DateTime(2100, 1, 1) : (DateTimeOffset)DateToDue;

            var foo = new ExpeditionPosition();


            var Query1 = (from a in dbContext.UnitPaymentOrders

                          join d in dbContext.PurchasingDocumentExpeditions on a.UPONo equals d.UnitPaymentOrderNo into m
                          from d in m.DefaultIfEmpty()
                          join b in dbContext.UnitPaymentOrderItems on a.Id equals b.UPOId into n
                          from b in n.DefaultIfEmpty()
                          join c in dbContext.UnitPaymentOrderDetails on b.Id equals c.UPOItemId into o
                          from c in o.DefaultIfEmpty()
                          join f in dbContext.UnitReceiptNotes on b.URNId equals f.Id

                          where
                  a.IsDeleted == false
                  && b.IsDeleted == false
                  && c.IsDeleted == false
                  && d.IsDeleted == false
                  // && d.Position == 0 ? d.Position== 0: d.Position == d.Position
                  && a.UPONo == (string.IsNullOrWhiteSpace(UnitPaymentOrderNo) ? a.UPONo : UnitPaymentOrderNo)
                  && a.SupplierCode == (string.IsNullOrWhiteSpace(SupplierCode) ? a.SupplierCode : SupplierCode)
                  && a.DivisionCode == (string.IsNullOrWhiteSpace(DivisionCode) ? a.DivisionCode : DivisionCode)
                  && a.DueDate.AddHours(Offset).Date >= dateFromDue.Date
                  && a.DueDate.AddHours(Offset).Date <= dateToDue.Date
                  && a.Date.AddHours(Offset).Date >= dateFrom.Date
                  && a.Date.AddHours(Offset).Date <= dateTo.Date

                          orderby d.Id descending

                          select new //UnitPaymentOrderPaidStatusViewModel
                          {


                              /* UnitPaymentOrderNo = a.UPONo,
                                 UPODate = a.Date,
                                 DueDate = a.DueDate,
                                 InvoiceNo = a.InvoiceNo,
                                 SupplierName = a.SupplierName,
                                 DivisionName = a.DivisionName,
                                 PaymentMethod = a.PaymentMethod,
                               DPP = c.PriceTotal,
                                 PPH = a.UseIncomeTax == true ? (a.IncomeTaxRate * c.PriceTotal) / 100 : 0,
                                 PPN = a.UseVat == true ? (c.PriceTotal * 10) / 100 : 0,
                                 Currency = a.CurrencyCode,
                                 BankExpenditureNoteDate = d.BankExpenditureNoteDate,
                                 BankExpenditureNoteNo = d.BankExpenditureNoteNo,
                                 BankExpenditureNotePPHDate = d.BankExpenditureNotePPHDate,
                                 BankExpenditureNotePPHNo = d.BankExpenditureNotePPHNo,
                                 UnitName = f.UnitName,
                                 IsPaid = d.IsPaid,
                                 IsPaidPPH = d.IsPaidPPH,
                                 UseIncomeTax = a.UseIncomeTax,
                                 UseVat = a.UseVat,
                                 CategoryName = a.CategoryName */

                              a.UPONo,
                              a.Date,
                              a.DueDate,
                              a.InvoiceNo,
                              a.SupplierName,
                              a.DivisionName,
                              a.PaymentMethod,
                              c.PriceTotal,
                              a.IncomeTaxRate,
                              a.CurrencyCode,
                              d.BankExpenditureNoteDate,
                              d.BankExpenditureNoteNo,
                              d.BankExpenditureNotePPHDate,
                              d.BankExpenditureNotePPHNo,
                              f.UnitName,
                              d.IsPaid,
                              d.IsPaidPPH,
                              a.UseIncomeTax,
                              a.UseVat,
                              a.CategoryName,


                              position = d.Position == null ? ExpeditionPosition.INVALID : d.Position



                          });

            var Query = (from query in Query1
                         group query by new { query.UnitName, query.DivisionName, query.UPONo } into groupdata







                         select new UnitPaymentOrderPaidStatusViewModel
                         {
                             UnitPaymentOrderNo = groupdata.Key.UPONo,
                             UPODate = groupdata.First().Date,
                             DueDate = groupdata.First().DueDate,
                             InvoiceNo = groupdata.First().InvoiceNo,
                             SupplierName = groupdata.First().SupplierName,
                             DivisionName = groupdata.Key.DivisionName,
                             PaymentMethod = groupdata.First().PaymentMethod,




                             DPP = (groupdata.Count(a => a.position == ExpeditionPosition.SEND_TO_PURCHASING_DIVISION)) != 0 &&
                                   (groupdata.Count(a => a.position != ExpeditionPosition.SEND_TO_PURCHASING_DIVISION)) != 0
                                    ?

                                    groupdata.Where(a => a.position != ExpeditionPosition.SEND_TO_PURCHASING_DIVISION).Sum(a => a.PriceTotal) :

                                    groupdata.Sum(a => a.PriceTotal),

                             PPH = groupdata.First().UseIncomeTax == true ? (
                                    (groupdata.Count(a => a.position == ExpeditionPosition.SEND_TO_PURCHASING_DIVISION)) != 0 &&
                                   (groupdata.Count(a => a.position != ExpeditionPosition.SEND_TO_PURCHASING_DIVISION)) != 0
                                    ?

                                    groupdata.Where(a => a.position != ExpeditionPosition.SEND_TO_PURCHASING_DIVISION).Sum(a => (a.IncomeTaxRate * a.PriceTotal) / 100) :

                                    groupdata.Sum(a => (a.IncomeTaxRate * a.PriceTotal) / 100)

                                    ) : 0,

                             PPN = groupdata.First().UseVat == true ? (

                                   (groupdata.Count(a => a.position == ExpeditionPosition.SEND_TO_PURCHASING_DIVISION)) != 0 &&
                                   (groupdata.Count(a => a.position != ExpeditionPosition.SEND_TO_PURCHASING_DIVISION)) != 0
                                    ?

                                    groupdata.Where(a => a.position != ExpeditionPosition.SEND_TO_PURCHASING_DIVISION).Sum(a => (a.PriceTotal * 10) / 100) :

                                    groupdata.Sum(a => (a.PriceTotal * 10) / 100)

                             ) : 0,

                             // groupdata.First().PriceTotal :

                             //(groupdata.Count(a => a.position == ExpeditionPosition.SEND_TO_PURCHASING_DIVISION)) != 0 &&
                             //(groupdata.Count(a => a.position != ExpeditionPosition.SEND_TO_PURCHASING_DIVISION)) == 0
                             //? groupdata.Sum(a => a.PriceTotal) :
                             //groupdata.Where(a => a.position == ExpeditionPosition.SEND_TO_PURCHASING_DIVISION).Sum(a => a.PriceTotal),



                             //PPH = groupdata.First().UseIncomeTax == true ? groupdata.Sum(a => (a.IncomeTaxRate * a.PriceTotal) / 100) : 0,
                             //PPN = groupdata.First().UseVat == true ? groupdata.Sum(a => (a.PriceTotal * 10) / 100) : 0,

                             //DPP =  groupdata.Sum(a => a.PriceTotal),







                             //groupdata.First().position != ExpeditionPosition.SEND_TO_PURCHASING_DIVISION ? groupdata.Sum(a => (a.IncomeTaxRate * a.PriceTotal) / 100) :

                             //(groupdata.First().IncomeTaxRate * groupdata.First().PriceTotal) / 100) : 0,



                             //groupdata.First().position != ExpeditionPosition.SEND_TO_PURCHASING_DIVISION ? groupdata.Sum(a => (a.PriceTotal * 10) / 100) :
                             //(groupdata.First().PriceTotal * 10) / 100

                             //  ) : 0,
                             Currency = groupdata.First().CurrencyCode,
                             BankExpenditureNoteDate = groupdata.First().BankExpenditureNoteDate,
                             BankExpenditureNoteNo = groupdata.First().BankExpenditureNoteNo,
                             BankExpenditureNotePPHDate = groupdata.First().BankExpenditureNotePPHDate,
                             BankExpenditureNotePPHNo = groupdata.First().BankExpenditureNotePPHNo,
                             UnitName = groupdata.Key.UnitName,
                             IsPaid = groupdata.First().IsPaid,
                             IsPaidPPH = groupdata.First().IsPaidPPH,
                             UseIncomeTax = groupdata.First().UseIncomeTax,
                             UseVat = groupdata.First().UseVat,
                             CategoryName = groupdata.First().CategoryName,
                             //  Position = groupdata.First().Position
                         });

            return Query;
        }

        public ReadResponse<object> GetReport(int Size, int Page, string Order, string UnitPaymentOrderNo, string SupplierCode, string DivisionCode, string Status, DateTimeOffset? DateFromDue, DateTimeOffset? DateToDue, DateTimeOffset? DateFrom, DateTimeOffset? DateTo, int Offset)
        {
            var Query = GetQuery(UnitPaymentOrderNo, SupplierCode, DivisionCode, Status, DateFromDue, DateToDue, DateFrom, DateTo, Offset);

            Dictionary<string, string> OrderDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Order);

            if (!string.IsNullOrWhiteSpace(Status))

            {
                if (Status.Equals("LUNAS"))
                {
                    Query = Query.Where(p => p.IsPaid == true && p.IsPaidPPH == false && p.UseVat == true && p.UseIncomeTax == false ||
                                            p.IsPaid == true && p.IsPaidPPH == true ||
                                            p.IsPaid == true && p.IsPaidPPH == false && p.UseIncomeTax == false && p.UseVat == false

                    );
                }
                else if (Status.Equals("SUDAH BAYAR DPP+PPN"))
                {
                    Query = Query.Where(p => p.IsPaid == true && p.IsPaidPPH == false && p.UseIncomeTax == true && p.UseVat == false && p.BankExpenditureNoteNo != null || p.IsPaid == true && p.IsPaidPPH == false && p.UseIncomeTax == true && p.UseVat == true);
                }
                else if (Status.Equals("SUDAH BAYAR PPH"))
                {
                    Query = Query.Where(p => p.IsPaidPPH == true && p.IsPaid == false && p.UseIncomeTax == true && p.UseVat == false || p.IsPaidPPH == true && p.IsPaid == false && p.UseIncomeTax == true && p.UseVat == true);
                }
                else if (Status.Equals("BELUM BAYAR"))
                {
                    Query = Query.Where(p => p.IsPaid == false && p.IsPaidPPH == false);
                }
              

            }




            Pageable<UnitPaymentOrderPaidStatusViewModel> pageable = new Pageable<UnitPaymentOrderPaidStatusViewModel>(Query, Page - 1, Size);
            List<UnitPaymentOrderPaidStatusViewModel> Data = pageable.Data.ToList<UnitPaymentOrderPaidStatusViewModel>();
            int TotalData = pageable.TotalCount;

            List<object> list = new List<object>();
            foreach (var datum in Data)
            {



                list.Add(new UnitPaymentOrderPaidStatusViewModel
                {
                    UnitPaymentOrderNo = datum.UnitPaymentOrderNo,
                    UPODate = datum.UPODate,
                    DueDate = datum.DueDate,
                    InvoiceNo = datum.InvoiceNo,
                    SupplierName = datum.SupplierName,
                    DivisionName = datum.DivisionName,
                    PaymentMethod = datum.PaymentMethod,
                    Status =
                                  (datum.IsPaid == true && datum.IsPaidPPH == false && datum.UseVat == true && datum.UseIncomeTax == false) ||

                                  (datum.IsPaid == true && datum.IsPaidPPH == true) ? "LUNAS" :
                                  //Sudah Bayar DPP+PPN
                                  (datum.IsPaid == true && datum.IsPaidPPH == false && datum.UseVat == true && datum.UseIncomeTax == true) ||
                                  (datum.IsPaid == true && datum.IsPaidPPH == false && datum.UseVat == false && datum.UseIncomeTax == true && datum.BankExpenditureNoteNo != null) ? "SUDAH BAYAR DPP+PPN" :
                                  //Sudah Bayar PPH               
                                  (datum.IsPaidPPH == true && datum.IsPaid == false && datum.UseVat == true && datum.UseIncomeTax == true) ||
                                  (datum.IsPaidPPH == true && datum.IsPaid == false && datum.UseVat == false && datum.UseIncomeTax == true && datum.BankExpenditureNoteNo == null) ? "SUDAH BAYAR PPH"
                                  : (datum.Position == 6) ? "BELUM BAYAR" : "BELUM BAYAR",


                    DPP = (datum.UseIncomeTax == true) ? datum.DPP - datum.PPH : datum.DPP,
                    PPH = datum.PPH,
                    PPN = datum.PPN,

                    TotalPaid =
                            (datum.UseVat == true && datum.UseIncomeTax == false) ? (datum.DPP) + datum.PPN :
                            (datum.DPP - datum.PPH) + datum.PPH + datum.PPN,
                    /*  (datum.IsPaid == true && datum.IsPaidPPH == false && datum.UseVat == true && datum.UseIncomeTax == true) ||
                      (datum.IsPaid == true && datum.IsPaidPPH == false && datum.UseVat == false && datum.UseIncomeTax == true) ? (datum.DPP + datum.PPN) : 
                      (datum.IsPaid == false && datum.IsPaidPPH == true && datum.UseVat == true && datum.UseIncomeTax == true) ||
                      (datum.IsPaid == false && datum.IsPaidPPH == true && datum.UseVat == false && datum.UseIncomeTax == true) ? datum.PPH:
                      (datum.DPP - datum.PPH)+datum.PPH + datum.PPN, */

                    Currency = datum.Currency,
                    BankExpenditureNotePPHDate = datum.BankExpenditureNotePPHDate,
                    BankExpenditureNotePPHNo = datum.BankExpenditureNotePPHNo,
                    BankExpenditureNoteNo = datum.BankExpenditureNoteNo,
                    BankExpenditureNoteDate = datum.BankExpenditureNoteDate,
                    statuspph = datum.UseIncomeTax == true ? "IYA" : "TIDAK",
                    statusppn = datum.UseVat == true ? "IYA" : "TIDAK",
                    CategoryName = datum.CategoryName,
                    UnitName = datum.UnitName
                });
            }

            return new ReadResponse<object>(list, TotalData, OrderDictionary);
        }

        //public ReadResponse<object> GetReport(int Size, int Page, string Order, string UnitPaymentOrderNo, string SupplierCode, string DivisionCode, string Status, DateTimeOffset? DateFrom, DateTimeOffset? DateTo, int Offset)
        //{
        //    IQueryable<PurchasingDocumentExpedition> Query = this.dbContext.PurchasingDocumentExpeditions;

        //    if (DateFrom == null || DateTo == null)
        //    {
        //        Query = Query
        //          .Where(p => p.IsDeleted == false &&
        //              p.UnitPaymentOrderNo == (UnitPaymentOrderNo == null ? p.UnitPaymentOrderNo : UnitPaymentOrderNo) &&
        //              p.SupplierCode == (SupplierCode == null ? p.SupplierCode : SupplierCode) &&
        //              p.DivisionCode == (DivisionCode == null ? p.DivisionCode : DivisionCode)
        //          );
        //    }
        //    else
        //    {
        //        Query = Query
        //           .Where(p => p.IsDeleted == false &&
        //                  p.UnitPaymentOrderNo == (UnitPaymentOrderNo == null ? p.UnitPaymentOrderNo : UnitPaymentOrderNo) &&
        //                  p.SupplierCode == (SupplierCode == null ? p.SupplierCode : SupplierCode) &&
        //                  p.DivisionCode == (DivisionCode == null ? p.DivisionCode : DivisionCode) &&
        //                  p.DueDate.Date >= DateFrom.Value.Date &&
        //                  p.DueDate.Date <= DateTo.Value.Date
        //           );
        //    }

        //    Query = Query
        //        .Select(s => new PurchasingDocumentExpedition
        //        {
        //            Id = s.Id,
        //            UnitPaymentOrderNo = s.UnitPaymentOrderNo,
        //            UPODate = s.UPODate,
        //            DueDate = s.DueDate,
        //            InvoiceNo = s.InvoiceNo,
        //            SupplierCode = s.SupplierCode,
        //            SupplierName = s.SupplierName,
        //            DivisionCode = s.DivisionCode,
        //            DivisionName = s.DivisionName,
        //            PaymentMethod = s.PaymentMethod,
        //            IsPaid = s.IsPaid,
        //            IsPaidPPH = s.IsPaidPPH,
        //            TotalPaid = s.TotalPaid,
        //            IncomeTax = s.IncomeTax,
        //            Vat = s.Vat,
        //            Currency = s.Currency,
        //            BankExpenditureNoteDate = s.BankExpenditureNoteDate,
        //            BankExpenditureNoteNo = s.BankExpenditureNoteNo,
        //            BankExpenditureNotePPHDate = s.BankExpenditureNotePPHDate,
        //            BankExpenditureNotePPHNo = s.BankExpenditureNotePPHNo,
        //            LastModifiedUtc = s.LastModifiedUtc
        //        });

        //    if(!string.IsNullOrWhiteSpace(Status))
        //    {
        //        if (Status.Equals("LUNAS"))
        //        {
        //            Query = Query.Where(p => p.IsPaid == true && p.IsPaidPPH == true);
        //        }
        //        else if (Status.Equals("SUDAH BAYAR DPP+PPN"))
        //        {
        //            Query = Query.Where(p => p.IsPaid == true && p.IsPaidPPH == false);
        //        }
        //        else if (Status.Equals("SUDAH BAYAR PPH"))
        //        {
        //            Query = Query.Where(p => p.IsPaidPPH == true && p.IsPaid == false);
        //        }
        //        else if (Status.Equals("BELUM BAYAR"))
        //        {
        //            Query = Query.Where(p => p.IsPaid == false && p.IsPaidPPH == false);
        //        }
        //    }

        //    Dictionary<string, string> OrderDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Order);
        //    Query = QueryHelper<PurchasingDocumentExpedition>.ConfigureOrder(Query, OrderDictionary);

        //    Pageable<PurchasingDocumentExpedition> pageable = new Pageable<PurchasingDocumentExpedition>(Query, Page - 1, Size);
        //    List<PurchasingDocumentExpedition> Data = pageable.Data.ToList<PurchasingDocumentExpedition>();
        //    int TotalData = pageable.TotalCount;

        //    string[] BENNo = Data.Where(p => p.BankExpenditureNoteNo != null).Select(p => p.BankExpenditureNoteNo).ToArray();
        //    string[] BENPPHNo = Data.Where(p => p.BankExpenditureNotePPHNo != null).Select(p => p.BankExpenditureNotePPHNo).ToArray();

        //    var BEN = this.dbContext.BankExpenditureNotes.Where(p => BENNo.Contains(p.DocumentNo)).ToList();
        //    var BENPPH = this.dbContext.PPHBankExpenditureNotes.Where(p => BENPPHNo.Contains(p.No)).ToList();

        //    List<object> list = new List<object>();

        //    foreach (var datum in Data)
        //    {
        //        var Bank = BEN.SingleOrDefault(p => p.DocumentNo == datum.BankExpenditureNoteNo);
        //        var BankPPH = BENPPH.SingleOrDefault(p => p.No == datum.BankExpenditureNotePPHNo);

        //        list.Add(new UnitPaymentOrderPaidStatusViewModel
        //        {
        //            UnitPaymentOrderNo = datum.UnitPaymentOrderNo,
        //            UPODate = datum.UPODate,
        //            DueDate = datum.DueDate,
        //            InvoiceNo = datum.InvoiceNo,
        //            SupplierName = datum.SupplierName,
        //            DivisionName = datum.DivisionName,
        //            PaymentMethod = datum.PaymentMethod,
        //            Status = (datum.IsPaid == true && datum.IsPaidPPH == true) ? "LUNAS" : (datum.IsPaid == true && datum.IsPaidPPH == false) ? "SUDAH BAYAR DPP+PPB" : (datum.IsPaidPPH == true && datum.IsPaid == false) ? "SUDAH BAYAR PPH" : "BELUM BAYAR",
        //            DPP = datum.TotalPaid - datum.Vat,
        //            IncomeTax = datum.IncomeTax,
        //            Vat = datum.Vat,
        //            TotalPaid = datum.TotalPaid,
        //            Currency = datum.Currency,
        //            BankExpenditureNotePPHDate = datum.BankExpenditureNotePPHDate,
        //            PPHBank = BankPPH != null ? string.Concat(BankPPH.BankAccountName, " - ", BankPPH.BankName, " - ", BankPPH.BankAccountNumber, " - ", BankPPH.Currency) : null,
        //            BankExpenditureNoteDate = datum.BankExpenditureNoteDate,
        //            Bank = Bank != null ? string.Concat(Bank.BankAccountName, " - ", Bank.BankName, " - ", Bank.BankAccountNumber, " - ", Bank.BankCurrencyCode) : null,
        //        });
        //    }

        //    return new ReadResponse<object>(list, TotalData, OrderDictionary);
        //}
    }
}
