﻿using Com.DanLiris.Service.Purchasing.Lib.Helpers;
using Com.DanLiris.Service.Purchasing.Lib.Helpers.ReadResponse;
using Com.DanLiris.Service.Purchasing.Lib.Interfaces;
using Com.DanLiris.Service.Purchasing.Lib.Models.Expedition;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.Expedition;
using Com.Moonlay.NetCore.Lib;
using Microsoft.EntityFrameworkCore;
using MongoDB.Bson;
using MongoDB.Driver;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Com.DanLiris.Service.Purchasing.Lib.Facades.Expedition
{
    public class UnitPaymentOrderUnpaidReportFacade : IUnitPaymentOrderUnpaidReportFacade
    {
        private readonly PurchasingDbContext dbContext;
        private readonly DbSet<PurchasingDocumentExpedition> dbSet;

        IMongoCollection<BsonDocument> collection;

        public UnitPaymentOrderUnpaidReportFacade(PurchasingDbContext dbContext)
        {
            this.dbContext = dbContext;
            this.dbSet = this.dbContext.Set<PurchasingDocumentExpedition>();

            MongoDbContext mongoDbContext = new MongoDbContext();
            this.collection = mongoDbContext.UnitPaymentOrder;
        }

        public List<UnitPaymentOrderUnpaidViewModel> GetReportMongo(string no, string supplierCode, DateTimeOffset? dateFrom, DateTimeOffset? dateTo)
        {
            string query = "{'$and' : [{_deleted : false}";
            if (no != null)
            {
                query += ",{ no : '" + no + "'}";
            }
            if (supplierCode != null)
            {
                query += ",{ supplier.code : '" + supplierCode + "'}";
            }
            if (dateFrom != null && dateTo != null)
            {
                query += ",{'$and' : [{ date : {'$gte' : new Date('" + dateFrom.ToString() + "')}}, " + "{date : {'$lte' : new Date('" + dateTo.ToString() + "')}}]}";
            }
            query += "]}";
            
            List<BsonDocument> bsonData = collection.Find(query).Project(Builders<BsonDocument>.Projection
                .Include("no").Include("date").Include("currency").Include("supplier").Include("invoceNo").Include("dueDate").Include("items")).ToList();
            List<UnitPaymentOrderUnpaidViewModel> listData = new List<UnitPaymentOrderUnpaidViewModel>();
            foreach( var data in bsonData)
            {
                var unitReceiptNotes = data["items"].AsBsonArray.AsQueryable().Select(m => m["unitReceiptNote"].AsBsonValue).ToList();
                foreach(var unitReceiptNote in unitReceiptNotes)
                {
                    var itemsCount = unitReceiptNote["items"].AsBsonArray.Count;
                    var product = unitReceiptNote["items"].AsBsonArray.AsQueryable().Select(x => x["product"].AsBsonValue["name"].ToString()).ToList();
                    var qty = unitReceiptNote["items"].AsBsonArray.AsQueryable().Select(x => x["deliveredQuantity"].ToDouble()).ToList();
                    for (int i = 0; i < itemsCount; i++)
                    {
                        listData.Add(new UnitPaymentOrderUnpaidViewModel
                        {
                            UnitPaymentOrderNo = data["no"].ToString(),
                            UPODate = data["date"].ToUniversalTime(),
                            Currency = data["currency"].AsBsonValue["code"].ToString(),
                            SupplierName = data["supplier"].AsBsonValue["name"].ToString(),
                            InvoiceNo = data["invoceNo"].ToString(),
                            DueDate = data["dueDate"].ToUniversalTime(),
                            ProductName = product[i],
                            Quantity = qty[i],
                            UnitName = unitReceiptNote.AsBsonDocument.Contains("unitId") ? unitReceiptNote["unit"].AsBsonValue["name"].ToString() : "Not Exist",
                        });
                    }       
                    
                }
            }

            return listData;
        }

        public ReadResponse GetReport(int Size, int Page, string Order, string UnitPaymentOrderNo, string SupplierCode, DateTimeOffset? DateFrom, DateTimeOffset? DateTo, int Offset)
        {

            IQueryable<PurchasingDocumentExpedition> Query = this.dbContext.PurchasingDocumentExpeditions;
            if(DateFrom == null || DateTo == null)
            {
                DateTo = DateTimeOffset.Now.AddHours(Offset);
                DateFrom = DateTimeOffset.Now.AddHours(Offset).AddMonths(-1);
            }

            Query = Query
                   .Where(p => p.IsDeleted == false &&
                          p.UnitPaymentOrderNo == (UnitPaymentOrderNo == null ? p.UnitPaymentOrderNo : UnitPaymentOrderNo) &&
                          p.SupplierCode == (SupplierCode == null ? p.SupplierCode : SupplierCode) &&
                          p.DueDate.Date >= DateFrom.Value.Date &&
                          p.DueDate.Date <= DateTo.Value.Date
                   );

            Query = Query
                .Select(s => new PurchasingDocumentExpedition
                {
                    Id = s.Id,
                    UnitPaymentOrderNo = s.UnitPaymentOrderNo,
                    UPODate = s.UPODate,
                    DueDate = s.DueDate,
                    InvoiceNo = s.InvoiceNo,
                    SupplierCode = s.SupplierCode,
                    SupplierName = s.SupplierName,
                    IsPaid = s.IsPaid,
                    IsPaidPPH = s.IsPaidPPH,
                    TotalPaid = s.TotalPaid,
                    IncomeTax = s.IncomeTax,
                    Vat = s.Vat,
                    Currency = s.Currency,
                    LastModifiedUtc = s.LastModifiedUtc
                });
            
            Dictionary<string, string> OrderDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Order);
            Query = QueryHelper<PurchasingDocumentExpedition>.ConfigureOrder(Query, OrderDictionary);
            List<UnitPaymentOrderUnpaidViewModel> list = new List<UnitPaymentOrderUnpaidViewModel>();

            list = (from datum in Query
                   select new UnitPaymentOrderUnpaidViewModel
                   {
                       UnitPaymentOrderNo = datum.UnitPaymentOrderNo,
                       UPODate = datum.UPODate,
                       InvoiceNo = datum.InvoiceNo,
                       SupplierName = datum.SupplierName,
                       Currency = datum.Currency,
                       IncomeTax = datum.IncomeTax,
                       DueDate = datum.DueDate,
                       DPPVat = datum.TotalPaid,
                       TotalPaid = datum.TotalPaid + datum.Vat
                   }).ToList();

            List<UnitPaymentOrderUnpaidViewModel> dataMongo = GetReportMongo(UnitPaymentOrderNo, SupplierCode, DateFrom, DateTo);

            var resultQuery = (from a in dataMongo
                                join b in list on a.UnitPaymentOrderNo equals b.UnitPaymentOrderNo into ab
                                from b in ab.DefaultIfEmpty()
                                select new UnitPaymentOrderUnpaidViewModel
                                {
                                    UnitPaymentOrderNo = a.UnitPaymentOrderNo,
                                    UPODate = a.UPODate,
                                    InvoiceNo = a.InvoiceNo,
                                    SupplierName = a.SupplierName,
                                    Currency = a.Currency,
                                    IncomeTax = b == null ? 0 : b.IncomeTax,
                                    DPPVat = b == null ? 0 : b.DPPVat,
                                    DueDate = a.DueDate,
                                    ProductName = a.ProductName,
                                    Quantity = a.Quantity,
                                    UnitName = a.UnitName,
                                    TotalPaid = 0
                                }).AsQueryable();
            
            Pageable<UnitPaymentOrderUnpaidViewModel> pageable = new Pageable<UnitPaymentOrderUnpaidViewModel>(resultQuery, Page - 1, Size);
            List<UnitPaymentOrderUnpaidViewModel> Data = pageable.Data.ToList<UnitPaymentOrderUnpaidViewModel>();
            int TotalData = pageable.TotalCount;

            return new ReadResponse(Data.ToList<object>(), TotalData, OrderDictionary);
        }
    }
}
