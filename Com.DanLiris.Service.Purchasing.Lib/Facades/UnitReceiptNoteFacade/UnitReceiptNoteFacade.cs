﻿using Com.DanLiris.Service.Purchasing.Lib.Helpers;
using Com.DanLiris.Service.Purchasing.Lib.Interfaces;
using Com.DanLiris.Service.Purchasing.Lib.Models.DeliveryOrderModel;
using Com.DanLiris.Service.Purchasing.Lib.Models.ExternalPurchaseOrderModel;
using Com.DanLiris.Service.Purchasing.Lib.Models.InternalPurchaseOrderModel;
using Com.DanLiris.Service.Purchasing.Lib.Models.PurchaseRequestModel;
using Com.DanLiris.Service.Purchasing.Lib.Models.UnitReceiptNoteModel;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.IntegrationViewModel;
using Com.Moonlay.Models;
using Com.Moonlay.NetCore.Lib;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace Com.DanLiris.Service.Purchasing.Lib.Facades.UnitReceiptNoteFacade
{
    public class UnitReceiptNoteFacade
    {
        private string USER_AGENT = "Facade";

        private readonly PurchasingDbContext dbContext;
        public readonly IServiceProvider serviceProvider;
        private readonly DbSet<UnitReceiptNote> dbSet;

        public UnitReceiptNoteFacade(IServiceProvider serviceProvider, PurchasingDbContext dbContext)
        {
            this.serviceProvider = serviceProvider;
            this.dbContext = dbContext;
            this.dbSet = dbContext.Set<UnitReceiptNote>();
        }

        public Tuple<List<UnitReceiptNote>, int, Dictionary<string, string>> Read(int Page = 1, int Size = 25, string Order = "{}", string Keyword = null, string Filter = "{}")
        {
            IQueryable<UnitReceiptNote> Query = this.dbSet;

            List<string> searchAttributes = new List<string>()
            {
                "URNNo", "UnitName", "SupplierName", "DONo","Items.PRNo"
            };

            Query = QueryHelper<UnitReceiptNote>.ConfigureSearch(Query, searchAttributes, Keyword);

            Query = Query.Select(s => new UnitReceiptNote
            {
                Id = s.Id,
                UId = s.UId,
                URNNo = s.URNNo,
                ReceiptDate = s.ReceiptDate,
                UnitName = s.UnitName,
                DivisionName = s.DivisionName,
                SupplierName = s.SupplierName,
                DONo = s.DONo,
                CreatedBy = s.CreatedBy,
                LastModifiedUtc = s.LastModifiedUtc,
                Items = s.Items.ToList()
            });

            

            Dictionary<string, string> FilterDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Filter);
            Query = QueryHelper<UnitReceiptNote>.ConfigureFilter(Query, FilterDictionary);

            Dictionary<string, string> OrderDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Order);
            Query = QueryHelper<UnitReceiptNote>.ConfigureOrder(Query, OrderDictionary);

            Pageable<UnitReceiptNote> pageable = new Pageable<UnitReceiptNote>(Query, Page - 1, Size);
            List<UnitReceiptNote> Data = pageable.Data.ToList<UnitReceiptNote>();
            int TotalData = pageable.TotalCount;

            return Tuple.Create(Data, TotalData, OrderDictionary);
        }

        public UnitReceiptNote ReadById(int id)
        {
            var a = this.dbSet.Where(p => p.Id == id)
                .Include(p => p.Items)
                .FirstOrDefault();
            return a;
        }


        async Task<string> GenerateNo(UnitReceiptNote model)
        {
            string Year = model.ReceiptDate.ToString("yy");
            string Month = model.ReceiptDate.ToString("MM");


            string no = $"{Year}-{Month}-{model.URNNo}-{model.UnitCode}-";
            int Padding = 3;

            var lastNo = await this.dbSet.Where(w => w.URNNo.StartsWith(no) && !w.IsDeleted).OrderByDescending(o => o.URNNo).FirstOrDefaultAsync();

            if (lastNo == null)
            {
                return no + "1".PadLeft(Padding, '0');
            }
            else
            {
                int lastNoNumber = Int32.Parse(lastNo.URNNo.Replace(no, "")) + 1;
                return no + lastNoNumber.ToString().PadLeft(Padding, '0');
            }
        }

        public async Task<int> Create(UnitReceiptNote m, string user)
        {
            int Created = 0;

            using (var transaction = this.dbContext.Database.BeginTransaction())
            {
                try
                {
                    EntityExtension.FlagForCreate(m, user, "Facade");

                    m.URNNo = await GenerateNo(m);

                    foreach (var item in m.Items)
                    {

                        EntityExtension.FlagForCreate(item, user, "Facade");
                        ExternalPurchaseOrderDetail externalPurchaseOrderDetail = this.dbContext.ExternalPurchaseOrderDetails.FirstOrDefault(s => s.Id == item.EPODetailId);
                        PurchaseRequestItem prItem = this.dbContext.PurchaseRequestItems.FirstOrDefault(s => s.Id == externalPurchaseOrderDetail.PRItemId);
                        InternalPurchaseOrderItem poItem = this.dbContext.InternalPurchaseOrderItems.FirstOrDefault(s => s.Id == externalPurchaseOrderDetail.POItemId);
                        DeliveryOrderDetail doDetail = dbContext.DeliveryOrderDetails.FirstOrDefault(s => s.Id == item.DODetailId);
                        item.PRItemId = doDetail.PRItemId;
                        item.PricePerDealUnit = externalPurchaseOrderDetail.PricePerDealUnit;
                        doDetail.ReceiptQuantity += item.ReceiptQuantity;
                        externalPurchaseOrderDetail.ReceiptQuantity += item.ReceiptQuantity;
                        if (externalPurchaseOrderDetail.DOQuantity >= externalPurchaseOrderDetail.DealQuantity)
                        {
                            if (externalPurchaseOrderDetail.ReceiptQuantity < externalPurchaseOrderDetail.DealQuantity)
                            {
                                //prItem.Status = "Barang sudah diterima Unit parsial";
                                poItem.Status = "Barang sudah diterima Unit parsial";
                            }
                            else
                            {
                                //prItem.Status = "Barang sudah diterima Unit semua";
                                poItem.Status = "Barang sudah diterima Unit semua";
                            }
                        }
                        else
                        {
                            //prItem.Status = "Barang sudah diterima Unit parsial";
                            poItem.Status = "Barang sudah diterima Unit parsial";
                        }
                        
                    }
                    if (m.IsStorage == true)
                    {
                        insertStorage(m, user, "IN");
                    }
                    this.dbSet.Add(m);
                    Created = await dbContext.SaveChangesAsync();
                    transaction.Commit();
                }
                catch (Exception e)
                {
                    transaction.Rollback();
                    throw new Exception(e.Message);
                }
            }

            return Created;
        }


        public async Task<int> Update(int id, UnitReceiptNote unitReceiptNote, string user)
        {
            int Updated = 0;

            using (var transaction = this.dbContext.Database.BeginTransaction())
            {
                try
                {
                    var m = this.dbSet.AsNoTracking()
                        .Include(d => d.Items)
                        .Single(pr => pr.Id == id && !pr.IsDeleted);

                    if (m != null && !id.Equals(unitReceiptNote.Id))
                    {
                        if (m.IsStorage == true)
                        {
                            insertStorage(m, user, "OUT");
                        }
                        if (unitReceiptNote.IsStorage == false)
                        {
                            unitReceiptNote.StorageCode = null;
                            unitReceiptNote.StorageId = null;
                            unitReceiptNote.StorageName = null;
                        }
                        EntityExtension.FlagForUpdate(unitReceiptNote, user, USER_AGENT);

                        foreach (var item in unitReceiptNote.Items)
                        {
                            if (item.Id == 0)
                            {
                                EntityExtension.FlagForCreate(item, user, USER_AGENT);
                                ExternalPurchaseOrderDetail externalPurchaseOrderDetail = this.dbContext.ExternalPurchaseOrderDetails.FirstOrDefault(s => s.Id == item.EPODetailId);
                                PurchaseRequestItem prItem = this.dbContext.PurchaseRequestItems.FirstOrDefault(s => s.Id == externalPurchaseOrderDetail.PRItemId);
                                InternalPurchaseOrderItem poItem = this.dbContext.InternalPurchaseOrderItems.FirstOrDefault(s => s.Id == externalPurchaseOrderDetail.POItemId);
                                DeliveryOrderDetail doDetail = dbContext.DeliveryOrderDetails.FirstOrDefault(s => s.Id == item.DODetailId);
                                item.PRItemId = doDetail.PRItemId;
                                item.PricePerDealUnit = externalPurchaseOrderDetail.PricePerDealUnit;
                                doDetail.ReceiptQuantity += item.ReceiptQuantity;
                                externalPurchaseOrderDetail.ReceiptQuantity += item.ReceiptQuantity;
                                if (externalPurchaseOrderDetail.DOQuantity >= externalPurchaseOrderDetail.DealQuantity)
                                {
                                    if (externalPurchaseOrderDetail.ReceiptQuantity < externalPurchaseOrderDetail.DealQuantity)
                                    {
                                        //prItem.Status = "Barang sudah diterima Unit parsial";
                                        poItem.Status = "Barang sudah diterima Unit parsial";
                                    }
                                    else
                                    {
                                        //prItem.Status = "Barang sudah diterima Unit semua";
                                        poItem.Status = "Barang sudah diterima Unit semua";
                                    }
                                }
                                else
                                {
                                    //prItem.Status = "Barang sudah diterima Unit parsial";
                                    poItem.Status = "Barang sudah diterima Unit parsial";
                                }
                            }
                            EntityExtension.FlagForUpdate(item, user, USER_AGENT);
                        }

                        this.dbContext.Update(unitReceiptNote);

                        foreach (var itemExist in m.Items)
                        {
                            var a = itemExist;
                            ExternalPurchaseOrderDetail epoDetail = this.dbContext.ExternalPurchaseOrderDetails.FirstOrDefault(s => s.Id == itemExist.EPODetailId);
                            //PurchaseRequestItem purchaseRequestItem = this.dbContext.PurchaseRequestItems.FirstOrDefault(s => s.Id == externalPurchaseOrderDetail.PRItemId);
                            InternalPurchaseOrderItem purchaseOrderItem = this.dbContext.InternalPurchaseOrderItems.FirstOrDefault(s => s.Id == epoDetail.POItemId);
                            DeliveryOrderDetail sjDetail = dbContext.DeliveryOrderDetails.FirstOrDefault(s => s.Id == itemExist.DODetailId);
                            itemExist.PRItemId = sjDetail.PRItemId;
                            itemExist.PricePerDealUnit = epoDetail.PricePerDealUnit;
                            UnitReceiptNoteItem unitReceiptNoteItem = unitReceiptNote.Items.FirstOrDefault(i => i.Id.Equals(itemExist.Id));
                            if (unitReceiptNoteItem == null)
                            {
                                EntityExtension.FlagForDelete(itemExist, user, USER_AGENT);
                                this.dbContext.UnitReceiptNoteItems.Update(itemExist);
                                sjDetail.ReceiptQuantity -= itemExist.ReceiptQuantity;
                                epoDetail.ReceiptQuantity -= itemExist.ReceiptQuantity;
                                if (epoDetail.ReceiptQuantity == 0)
                                {
                                    if (epoDetail.DOQuantity > 0 && epoDetail.DOQuantity >= epoDetail.DealQuantity)
                                    {
                                        //prItem.Status = "Barang sudah diterima Unit semua";
                                        purchaseOrderItem.Status = "Barang sudah datang semua";
                                    }
                                    else if (epoDetail.DOQuantity > 0 && epoDetail.DOQuantity < epoDetail.DealQuantity)
                                    {
                                        //prItem.Status = "Barang sudah diterima Unit parsial";
                                        purchaseOrderItem.Status = "Barang sudah datang parsial";
                                    }
                                }
                                else if (epoDetail.ReceiptQuantity > 0)
                                {
                                    if (epoDetail.DOQuantity >= epoDetail.DealQuantity)
                                    {
                                        if (epoDetail.ReceiptQuantity < epoDetail.DealQuantity)
                                        {
                                            //prItem.Status = "Barang sudah diterima Unit parsial";
                                            purchaseOrderItem.Status = "Barang sudah diterima Unit parsial";
                                        }
                                        else if (epoDetail.ReceiptQuantity >= epoDetail.DealQuantity)
                                        {
                                            //prItem.Status = "Barang sudah diterima Unit semua";
                                            purchaseOrderItem.Status = "Barang sudah diterima Unit semua";
                                        }
                                        else if (epoDetail.DOQuantity < epoDetail.DealQuantity)
                                        {
                                            purchaseOrderItem.Status = "Barang sudah diterima Unit parsial";
                                        }
                                    }
                                }
                            }
                            else
                            {
                                EntityExtension.FlagForUpdate(itemExist, user, USER_AGENT);

                                sjDetail.ReceiptQuantity -= itemExist.ReceiptQuantity;
                                epoDetail.ReceiptQuantity -= itemExist.ReceiptQuantity;
                                itemExist.PRItemId = sjDetail.PRItemId;
                                itemExist.PricePerDealUnit = epoDetail.PricePerDealUnit;
                                sjDetail.ReceiptQuantity += unitReceiptNoteItem.ReceiptQuantity;
                                epoDetail.ReceiptQuantity += unitReceiptNoteItem.ReceiptQuantity;
                                if (epoDetail.ReceiptQuantity == 0)
                                {
                                    if (epoDetail.DOQuantity > 0 && epoDetail.DOQuantity >= epoDetail.DealQuantity)
                                    {
                                        //prItem.Status = "Barang sudah diterima Unit semua";
                                        purchaseOrderItem.Status = "Barang sudah datang semua";
                                    }
                                    else if (epoDetail.DOQuantity > 0 && epoDetail.DOQuantity < epoDetail.DealQuantity)
                                    {
                                        //prItem.Status = "Barang sudah diterima Unit parsial";
                                        purchaseOrderItem.Status = "Barang sudah datang semua";
                                    }
                                }
                                else if (epoDetail.ReceiptQuantity > 0)
                                {
                                    if (epoDetail.DOQuantity >= epoDetail.DealQuantity)
                                    {
                                        if (epoDetail.ReceiptQuantity < epoDetail.DealQuantity)
                                        {
                                            //prItem.Status = "Barang sudah diterima Unit parsial";
                                            purchaseOrderItem.Status = "Barang sudah diterima Unit parsial";
                                        }
                                        else if (epoDetail.ReceiptQuantity >= epoDetail.DealQuantity)
                                        {
                                            //prItem.Status = "Barang sudah diterima Unit semua";
                                            purchaseOrderItem.Status = "Barang sudah diterima Unit semua";
                                        }
                                        else if (epoDetail.DOQuantity < epoDetail.DealQuantity)
                                        {
                                            purchaseOrderItem.Status = "Barang sudah diterima Unit parsial";
                                        }
                                    }
                                }
                            }

                        }

                        if (unitReceiptNote.IsStorage == true)
                        {
                            insertStorage(unitReceiptNote, user, "IN");
                        }

                        Updated = await dbContext.SaveChangesAsync();
                        transaction.Commit();
                    }
                    else
                    {
                        throw new Exception("Invalid Id");
                    }
                }
                catch (Exception e)
                {
                    transaction.Rollback();
                    throw new Exception(e.Message);
                }
            }

            return Updated;
        }

        public int Delete(int id, string user)
        {
            int Deleted = 0;

            using (var transaction = this.dbContext.Database.BeginTransaction())
            {
                try
                {
                    var m = this.dbSet
                        .Include(d => d.Items)
                        .SingleOrDefault(pr => pr.Id == id && !pr.IsDeleted);

                    EntityExtension.FlagForDelete(m, user, USER_AGENT);

                    foreach (var item in m.Items)
                    {
                        EntityExtension.FlagForDelete(item, user, USER_AGENT);

                        ExternalPurchaseOrderDetail externalPurchaseOrderDetail = this.dbContext.ExternalPurchaseOrderDetails.FirstOrDefault(s => s.Id == item.EPODetailId);
                        PurchaseRequestItem prItem = this.dbContext.PurchaseRequestItems.FirstOrDefault(s => s.Id == externalPurchaseOrderDetail.PRItemId);
                        InternalPurchaseOrderItem poItem = this.dbContext.InternalPurchaseOrderItems.FirstOrDefault(s => s.Id == externalPurchaseOrderDetail.POItemId);
                        DeliveryOrderDetail doDetail = dbContext.DeliveryOrderDetails.FirstOrDefault(s => s.Id == item.DODetailId);

                        doDetail.ReceiptQuantity -= item.ReceiptQuantity;
                        externalPurchaseOrderDetail.ReceiptQuantity -= item.ReceiptQuantity;
                        if (externalPurchaseOrderDetail.ReceiptQuantity == 0)
                        {
                            if (externalPurchaseOrderDetail.DOQuantity > 0 && externalPurchaseOrderDetail.DOQuantity >= externalPurchaseOrderDetail.DealQuantity)
                            {
                                //prItem.Status = "Barang sudah diterima Unit semua";
                                poItem.Status = "Barang sudah datang semua";
                            }
                            else if (externalPurchaseOrderDetail.DOQuantity > 0 && externalPurchaseOrderDetail.DOQuantity < externalPurchaseOrderDetail.DealQuantity)
                            {
                                //prItem.Status = "Barang sudah diterima Unit parsial";
                                poItem.Status = "Barang sudah datang parsial";
                            }
                        }
                        else if (externalPurchaseOrderDetail.ReceiptQuantity > 0)
                        {
                            if (externalPurchaseOrderDetail.DOQuantity >= externalPurchaseOrderDetail.DealQuantity)
                            {
                                if (externalPurchaseOrderDetail.ReceiptQuantity < externalPurchaseOrderDetail.DealQuantity)
                                {
                                    //prItem.Status = "Barang sudah diterima Unit parsial";
                                    poItem.Status = "Barang sudah diterima Unit parsial";
                                }
                                else if (externalPurchaseOrderDetail.ReceiptQuantity >= externalPurchaseOrderDetail.DealQuantity)
                                {
                                    //prItem.Status = "Barang sudah diterima Unit semua";
                                    poItem.Status = "Barang sudah diterima Unit semua";
                                }
                                else if(externalPurchaseOrderDetail.DOQuantity < externalPurchaseOrderDetail.DealQuantity)
                                {
                                    poItem.Status = "Barang sudah diterima Unit parsial";
                                }
                            }
                        }
                    }
                    if (m.IsStorage == true)
                    {
                        insertStorage(m, user, "OUT");
                    }

                    Deleted = dbContext.SaveChanges();
                    transaction.Commit();
                }
                catch (Exception e)
                {
                    transaction.Rollback();
                    throw new Exception(e.Message);
                }
            }

            return Deleted;
        }

        public void insertStorage( UnitReceiptNote unitReceiptNote, string user,string type)
        {
                List<object> items =new List<object>();
                foreach(var item in unitReceiptNote.Items)
                {
                    items.Add(new {
                        productId=item.ProductId,
                        productcode = item.ProductCode,
                        productname = item.ProductName,
                        uomId=item.UomId,
                        uom=item.Uom,
                        quantity=item.ReceiptQuantity,
                        remark=item.ProductRemark
                    });
                }
                var data = new
                {
                    storageId=unitReceiptNote.StorageId,
                    storagecode = unitReceiptNote.StorageCode,
                    storagename = unitReceiptNote.StorageName,
                    referenceNo=unitReceiptNote.URNNo,
                    referenceType= "Bon Terima Unit - " + unitReceiptNote.UnitName,
                    type=type,
                    remark=unitReceiptNote.Remark,
                    date=unitReceiptNote.ReceiptDate,
                    items=items
                };
                string inventoryUri = "inventory-documents";
                IHttpClientService httpClient = (IHttpClientService)this.serviceProvider.GetService(typeof(IHttpClientService));
                var response = httpClient.PostAsync($"{APIEndpoint.Inventory}{inventoryUri}", new StringContent(JsonConvert.SerializeObject(data).ToString(), Encoding.UTF8, General.JsonMediaType)).Result;
                response.EnsureSuccessStatusCode();
                
        }

        public Tuple<List<UnitReceiptNote>, int, Dictionary<string, string>> ReadBySupplierUnit(int Page = 1, int Size = 25, string Order = "{}", string Keyword = null, string Filter = "{}")
        {
            IQueryable<UnitReceiptNote> Query = this.dbSet;

            List<string> searchAttributes = new List<string>()
            {
                "URNNo", "DONo",
            };
            Query = QueryHelper<UnitReceiptNote>.ConfigureSearch(Query, searchAttributes, Keyword);

            Dictionary<string, string> FilterDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Filter);

            Query = Query
                .Where(w =>
                    w.IsDeleted == false &&
                    w.IsPaid == false &&
                    dbContext.DeliveryOrderItems.Any(doItem =>
                        doItem.DOId == w.DOId &&
                        dbContext.ExternalPurchaseOrders.Any(epo =>
                            epo.Id == doItem.EPOId &&
                            epo.DivisionId.Equals(FilterDictionary.GetValueOrDefault("DivisionId") ?? "") &&
                            epo.SupplierId.Equals(FilterDictionary.GetValueOrDefault("SupplierId") ?? "") &&
                            epo.PaymentMethod.Equals(FilterDictionary.GetValueOrDefault("PaymentMethod") ?? "") &&
                            epo.CurrencyCode.Equals(FilterDictionary.GetValueOrDefault("CurrencyCode") ?? "") &&
                            string.Concat("", epo.IncomeTaxId).Equals(FilterDictionary.GetValueOrDefault("IncomeTaxId") ?? "") &&
                            epo.UseVat == Boolean.Parse(FilterDictionary.GetValueOrDefault("UseVat") ?? "false") &&
                            epo.Items.Any(epoItem => 
                                dbContext.InternalPurchaseOrders.Any(po =>
                                    po.Id == epoItem.POId &&
                                    po.CategoryId.Equals(FilterDictionary.GetValueOrDefault("CategoryId") ?? "")
                                )
                            )
                        )
                    )
                )
                .Select(s => new UnitReceiptNote
                {
                    Id = s.Id,
                    UId = s.UId,
                    URNNo = s.URNNo,
                    ReceiptDate = s.ReceiptDate,
                    UnitName = s.UnitName,
                    DivisionName = s.DivisionName,
                    SupplierName = s.SupplierName,
                    DONo = s.DONo,
                    CreatedBy = s.CreatedBy,
                    LastModifiedUtc = s.LastModifiedUtc,
                    Items = s.Items.ToList()
                });

            Dictionary<string, string> OrderDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Order);
            Query = QueryHelper<UnitReceiptNote>.ConfigureOrder(Query, OrderDictionary);

            Pageable<UnitReceiptNote> pageable = new Pageable<UnitReceiptNote>(Query, Page - 1, Size);
            List<UnitReceiptNote> Data = pageable.Data.ToList<UnitReceiptNote>();
            int TotalData = pageable.TotalCount;

            return Tuple.Create(Data, TotalData, OrderDictionary);
        }
    }
}
