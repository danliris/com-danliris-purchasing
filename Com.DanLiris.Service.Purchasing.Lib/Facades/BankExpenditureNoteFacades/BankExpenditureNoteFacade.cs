﻿using Com.DanLiris.Service.Purchasing.Lib.Enums;
using Com.DanLiris.Service.Purchasing.Lib.Helpers;
using Com.DanLiris.Service.Purchasing.Lib.Helpers.ReadResponse;
using Com.DanLiris.Service.Purchasing.Lib.Interfaces;
using Com.DanLiris.Service.Purchasing.Lib.Models.BankExpenditureNoteModel;
using Com.DanLiris.Service.Purchasing.Lib.Models.Expedition;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.BankExpenditureNote;
using Com.Moonlay.Models;
using Com.Moonlay.NetCore.Lib;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Com.DanLiris.Service.Purchasing.Lib.Facades.BankExpenditureNoteFacades
{
    public class BankExpenditureNoteFacade : IBankExpenditureNoteFacade, IReadByIdable<BankExpenditureNoteModel>
    {
        private readonly PurchasingDbContext dbContext;
        private readonly DbSet<BankExpenditureNoteModel> dbSet;
        private readonly DbSet<BankExpenditureNoteDetailModel> detailDbSet;
        private readonly IBankDocumentNumberGenerator bankDocumentNumberGenerator;

        private readonly string USER_AGENT = "Facade";

        public BankExpenditureNoteFacade(PurchasingDbContext dbContext, IBankDocumentNumberGenerator bankDocumentNumberGenerator)
        {
            this.dbContext = dbContext;
            this.bankDocumentNumberGenerator = bankDocumentNumberGenerator;
            dbSet = dbContext.Set<BankExpenditureNoteModel>();
            detailDbSet = dbContext.Set<BankExpenditureNoteDetailModel>();
        }

        public async Task<BankExpenditureNoteModel> ReadById(int id)
        {
            return await this.dbContext.BankExpenditureNotes
                .AsNoTracking()
                    .Include(p => p.Details)
                        .ThenInclude(p => p.Items)
                .Where(d => d.Id.Equals(id) && d.IsDeleted.Equals(false))
                .FirstOrDefaultAsync();
        }

        public ReadResponse Read(int Page = 1, int Size = 25, string Order = "{}", string Keyword = null, string Filter = "{}")
        {
            IQueryable<BankExpenditureNoteModel> Query = this.dbSet;

            Query = Query
                .Select(s => new BankExpenditureNoteModel
                {
                    Id = s.Id,
                    CreatedUtc = s.CreatedUtc,
                    LastModifiedUtc = s.LastModifiedUtc,
                    BankName = s.BankName,
                    BankAccountName = s.BankAccountName,
                    BankAccountNumber =s.BankAccountNumber,
                    DocumentNo = s.DocumentNo,
                    SupplierName = s.SupplierName,
                    GrandTotal = s.GrandTotal,
                    BankCurrencyCode = s.BankCurrencyCode,
                    Details = s.Details.Where(w => w.BankExpenditureNoteId == s.Id).ToList()
                });

            List<string> searchAttributes = new List<string>()
            {
                "DocumentNo", "BankName", "SupplierName","BankCurrencyCode"
            };

            Query = QueryHelper<BankExpenditureNoteModel>.ConfigureSearch(Query, searchAttributes, Keyword);

            Dictionary<string, string> FilterDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Filter);
            Query = QueryHelper<BankExpenditureNoteModel>.ConfigureFilter(Query, FilterDictionary);

            Dictionary<string, string> OrderDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Order);
            Query = QueryHelper<BankExpenditureNoteModel>.ConfigureOrder(Query, OrderDictionary);

            Pageable<BankExpenditureNoteModel> pageable = new Pageable<BankExpenditureNoteModel>(Query, Page - 1, Size);
            List<BankExpenditureNoteModel> Data = pageable.Data.ToList();

            List<object> list = new List<object>();
            list.AddRange(
               Data.Select(s => new
               {
                   s.Id,
                   s.DocumentNo,
                   s.CreatedUtc,
                   s.BankName,
                   s.BankAccountName,
                   s.BankAccountNumber,
                   s.SupplierName,
                   s.GrandTotal,
                   s.BankCurrencyCode,
                   Details = s.Details.Select(sl => new { sl.SupplierName, sl.UnitPaymentOrderNo }).ToList(),
               }).ToList()
            );

            int TotalData = pageable.TotalCount;

            return new ReadResponse(list, TotalData, OrderDictionary);
        }

        public async Task<int> Update(int id, BankExpenditureNoteModel model, string username)
        {
            int Updated = 0;

            using (var transaction = this.dbContext.Database.BeginTransaction())
            {
                try
                {
                    EntityExtension.FlagForUpdate(model, username, USER_AGENT);
                    dbContext.Entry(model).Property(x => x.GrandTotal).IsModified = true;
                    dbContext.Entry(model).Property(x => x.LastModifiedAgent).IsModified = true;
                    dbContext.Entry(model).Property(x => x.LastModifiedBy).IsModified = true;
                    dbContext.Entry(model).Property(x => x.LastModifiedUtc).IsModified = true;

                    foreach (var detail in model.Details)
                    {
                        if (detail.Id == 0)
                        {
                            EntityExtension.FlagForCreate(detail, username, USER_AGENT);
                            dbContext.BankExpenditureNoteDetails.Add(detail);

                            PurchasingDocumentExpedition pde = new PurchasingDocumentExpedition
                            {
                                Id = (int)detail.UnitPaymentOrderId,
                                IsPaid = true,
                                BankExpenditureNoteNo = model.DocumentNo,
                                BankExpenditureNoteDate = model.Date
                            };

                            EntityExtension.FlagForUpdate(pde, username, USER_AGENT);
                            //dbContext.Attach(pde);
                            dbContext.Entry(pde).Property(x => x.IsPaid).IsModified = true;
                            dbContext.Entry(pde).Property(x => x.BankExpenditureNoteNo).IsModified = true;
                            dbContext.Entry(pde).Property(x => x.BankExpenditureNoteDate).IsModified = true;
                            dbContext.Entry(pde).Property(x => x.LastModifiedAgent).IsModified = true;
                            dbContext.Entry(pde).Property(x => x.LastModifiedBy).IsModified = true;
                            dbContext.Entry(pde).Property(x => x.LastModifiedUtc).IsModified = true;

                            foreach (var item in detail.Items)
                            {
                                EntityExtension.FlagForCreate(item, username, USER_AGENT);
                            }
                        }
                    }

                    foreach (var detail in dbContext.BankExpenditureNoteDetails.AsNoTracking().Where(p => p.BankExpenditureNoteId == model.Id))
                    {
                        BankExpenditureNoteDetailModel detailModel = model.Details.FirstOrDefault(prop => prop.Id.Equals(detail.Id));

                        if (detailModel == null)
                        {
                            EntityExtension.FlagForDelete(detail, username, USER_AGENT);

                            foreach (var item in dbContext.BankExpenditureNoteItems.AsNoTracking().Where(p => p.BankExpenditureNoteDetailId == detail.Id))
                            {
                                EntityExtension.FlagForDelete(item, username, USER_AGENT);
                                dbContext.BankExpenditureNoteItems.Update(item);
                            }

                            dbContext.BankExpenditureNoteDetails.Update(detail);

                            PurchasingDocumentExpedition pde = new PurchasingDocumentExpedition
                            {
                                Id = (int)detail.UnitPaymentOrderId,
                                IsPaid = false,
                                BankExpenditureNoteNo = null,
                                BankExpenditureNoteDate = null
                            };

                            EntityExtension.FlagForUpdate(pde, username, USER_AGENT);
                            //dbContext.Attach(pde);
                            dbContext.Entry(pde).Property(x => x.IsPaid).IsModified = true;
                            dbContext.Entry(pde).Property(x => x.BankExpenditureNoteNo).IsModified = true;
                            dbContext.Entry(pde).Property(x => x.BankExpenditureNoteDate).IsModified = true;
                            dbContext.Entry(pde).Property(x => x.LastModifiedAgent).IsModified = true;
                            dbContext.Entry(pde).Property(x => x.LastModifiedBy).IsModified = true;
                            dbContext.Entry(pde).Property(x => x.LastModifiedUtc).IsModified = true;
                        }
                    }

                    Updated = await dbContext.SaveChangesAsync();
                    transaction.Commit();
                }
                catch (Exception e)
                {
                    transaction.Rollback();
                    throw new Exception(e.Message);
                }
            }

            return Updated;
        }

        public async Task<int> Create(BankExpenditureNoteModel model, string username)
        {
            int Created = 0;

            using (var transaction = dbContext.Database.BeginTransaction())
            {
                try
                {
                    EntityExtension.FlagForCreate(model, username, USER_AGENT);

                    model.DocumentNo = await bankDocumentNumberGenerator.GenerateDocumentNumber("K", model.BankCode, username);

                    foreach (var detail in model.Details)
                    {
                        EntityExtension.FlagForCreate(detail, username, USER_AGENT);

                        PurchasingDocumentExpedition pde = new PurchasingDocumentExpedition
                        {
                            Id = (int)detail.UnitPaymentOrderId,
                            IsPaid = true,
                            BankExpenditureNoteNo = model.DocumentNo,
                            BankExpenditureNoteDate = model.Date
                        };

                        EntityExtension.FlagForUpdate(pde, username, USER_AGENT);
                        dbContext.Attach(pde);
                        dbContext.Entry(pde).Property(x => x.IsPaid).IsModified = true;
                        dbContext.Entry(pde).Property(x => x.BankExpenditureNoteNo).IsModified = true;
                        dbContext.Entry(pde).Property(x => x.BankExpenditureNoteDate).IsModified = true;
                        dbContext.Entry(pde).Property(x => x.LastModifiedAgent).IsModified = true;
                        dbContext.Entry(pde).Property(x => x.LastModifiedBy).IsModified = true;
                        dbContext.Entry(pde).Property(x => x.LastModifiedUtc).IsModified = true;

                        foreach (var item in detail.Items)
                        {
                            EntityExtension.FlagForCreate(item, username, USER_AGENT);
                        }
                    }

                    dbSet.Add(model);
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

        public async Task<int> Delete(int Id, string username)
        {
            int Count = 0;

            if (dbSet.Count(p => p.Id == Id && p.IsDeleted == false).Equals(0))
            {
                return 0;
            }

            using (var transaction = dbContext.Database.BeginTransaction())
            {
                try
                {
                    BankExpenditureNoteModel bankExpenditureNote = dbContext.BankExpenditureNotes.Single(p => p.Id == Id);

                    ICollection<BankExpenditureNoteDetailModel> Details = new List<BankExpenditureNoteDetailModel>(dbContext.BankExpenditureNoteDetails.Where(p => p.BankExpenditureNoteId.Equals(Id)));

                    foreach (var detail in Details)
                    {
                        ICollection<BankExpenditureNoteItemModel> Items = new List<BankExpenditureNoteItemModel>(dbContext.BankExpenditureNoteItems.Where(p => p.BankExpenditureNoteDetailId.Equals(detail.Id)));

                        foreach (var item in Items)
                        {
                            EntityExtension.FlagForDelete(item, username, USER_AGENT);
                            dbContext.BankExpenditureNoteItems.Update(item);
                        }

                        EntityExtension.FlagForDelete(detail, username, USER_AGENT);
                        dbContext.BankExpenditureNoteDetails.Update(detail);

                        PurchasingDocumentExpedition pde = new PurchasingDocumentExpedition
                        {
                            Id = (int)detail.UnitPaymentOrderId,
                            IsPaid = false,
                            BankExpenditureNoteNo = null,
                            BankExpenditureNoteDate = null
                        };

                        EntityExtension.FlagForUpdate(pde, username, USER_AGENT);
                        //dbContext.Attach(pde);
                        dbContext.Entry(pde).Property(x => x.IsPaid).IsModified = true;
                        dbContext.Entry(pde).Property(x => x.BankExpenditureNoteNo).IsModified = true;
                        dbContext.Entry(pde).Property(x => x.BankExpenditureNoteDate).IsModified = true;
                        dbContext.Entry(pde).Property(x => x.LastModifiedAgent).IsModified = true;
                        dbContext.Entry(pde).Property(x => x.LastModifiedBy).IsModified = true;
                        dbContext.Entry(pde).Property(x => x.LastModifiedUtc).IsModified = true;
                    }

                    EntityExtension.FlagForDelete(bankExpenditureNote, username, USER_AGENT);
                    dbSet.Update(bankExpenditureNote);
                    Count = await dbContext.SaveChangesAsync();

                    transaction.Commit();
                }
                catch (DbUpdateConcurrencyException e)
                {
                    transaction.Rollback();
                    throw new Exception(e.Message);
                }
                catch (Exception e)
                {
                    transaction.Rollback();
                    throw new Exception(e.Message);
                }
            }

            return Count;
        }

        public ReadResponse GetAllByPosition(int Page, int Size, string Order, string Keyword, string Filter)
        {
            IQueryable<PurchasingDocumentExpedition> Query = dbContext.PurchasingDocumentExpeditions;

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
                    CategoryCode = s.CategoryCode,
                    CategoryName = s.CategoryName,
                    DivisionCode = s.DivisionCode,
                    DivisionName = s.DivisionName,
                    TotalPaid = s.TotalPaid,
                    Currency = s.Currency,
                    Position = s.Position,
                    VerifyDate = s.VerifyDate,
                    Vat = s.Vat,
                    IsPaid = s.IsPaid,
                    PaymentMethod = s.PaymentMethod,
                    Items = s.Items.Where(w => w.PurchasingDocumentExpeditionId == s.Id).ToList(),
                    LastModifiedUtc = s.LastModifiedUtc
                });

            List<string> searchAttributes = new List<string>()
            {
                "UnitPaymentOrderNo", "SupplierName", "DivisionName", "SupplierCode", "InvoiceNo"
            };

            Query = QueryHelper<PurchasingDocumentExpedition>.ConfigureSearch(Query, searchAttributes, Keyword);

            if (Filter.Contains("verificationFilter"))
            {
                Filter = "{}";
                List<ExpeditionPosition> positions = new List<ExpeditionPosition> { ExpeditionPosition.SEND_TO_PURCHASING_DIVISION, ExpeditionPosition.SEND_TO_ACCOUNTING_DIVISION, ExpeditionPosition.SEND_TO_CASHIER_DIVISION };
                Query = Query.Where(p => positions.Contains(p.Position));
            }

            Dictionary<string, string> FilterDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Filter);
            Query = QueryHelper<PurchasingDocumentExpedition>.ConfigureFilter(Query, FilterDictionary);

            Dictionary<string, string> OrderDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Order);
            Query = QueryHelper<PurchasingDocumentExpedition>.ConfigureOrder(Query, OrderDictionary);

            Pageable<PurchasingDocumentExpedition> pageable = new Pageable<PurchasingDocumentExpedition>(Query, Page - 1, Size);
            List<PurchasingDocumentExpedition> Data = pageable.Data.ToList();
            int TotalData = pageable.TotalCount;

            List<object> list = new List<object>();
            list.AddRange(Data.Select(s => new
            {
                UnitPaymentOrderId = s.Id,
                s.UnitPaymentOrderNo,
                s.UPODate,
                s.DueDate,
                s.InvoiceNo,
                s.SupplierCode,
                s.SupplierName,
                s.CategoryCode,
                s.CategoryName,
                s.DivisionCode,
                s.DivisionName,
                s.Vat,
                s.IsPaid,
                s.TotalPaid,
                s.Currency,
                s.PaymentMethod,
                Items = s.Items.Select(sl => new
                {
                    UnitPaymentOrderItemId = sl.Id,
                    sl.UnitId,
                    sl.UnitCode,
                    sl.UnitName,
                    sl.ProductId,
                    sl.ProductCode,
                    sl.ProductName,
                    sl.Quantity,
                    sl.Uom,
                    sl.Price
                }).ToList()
            }));

            return new ReadResponse(list, TotalData, OrderDictionary);
        }

        public ReadResponse GetReport(int Size, int Page, string DocumentNo, string UnitPaymentOrderNo, string InvoiceNo, string SupplierCode, string DivisionCode, string PaymentMethod, DateTimeOffset? DateFrom, DateTimeOffset? DateTo, int Offset)
        {
            IQueryable<BankExpenditureNoteReportViewModel> Query;

            if (DateFrom == null || DateTo == null)
            {
                Query = (from a in dbContext.BankExpenditureNotes
                         join b in dbContext.BankExpenditureNoteDetails on a.Id equals b.BankExpenditureNoteId
                         join c in dbContext.PurchasingDocumentExpeditions on b.UnitPaymentOrderId equals c.Id
                         where c.InvoiceNo == (InvoiceNo ?? c.InvoiceNo)
                            && c.SupplierCode == (SupplierCode ?? c.SupplierCode)
                            && c.UnitPaymentOrderNo == (UnitPaymentOrderNo ?? c.UnitPaymentOrderNo)
                            && c.DivisionCode == (DivisionCode ?? c.DivisionCode)
                            && !c.PaymentMethod.ToUpper().Equals("CASH")
                            && c.IsPaid
                            && c.PaymentMethod == (PaymentMethod ?? c.PaymentMethod)
                         where a.DocumentNo == (DocumentNo ?? a.DocumentNo)
                         orderby a.DocumentNo
                         select new BankExpenditureNoteReportViewModel
                         {
                             DocumentNo = a.DocumentNo,
                             Currency = a.BankCurrencyCode,
                             Date = a.Date,
                             SupplierName = c.SupplierName,
                             CategoryName = c.CategoryName == null ? "-" : c.CategoryName,
                             DivisionName = c.DivisionName,
                             PaymentMethod = c.PaymentMethod,
                             UnitPaymentOrderNo = b.UnitPaymentOrderNo,
                             BankName = string.Concat(a.BankAccountName, " - ", a.BankName, " - ", a.BankAccountNumber, " - ", a.BankCurrencyCode),
                             DPP = c.TotalPaid - c.Vat,
                             VAT = c.Vat,
                             TotalPaid = c.TotalPaid,
                             InvoiceNumber = c.InvoiceNo
                         }
                      );
            }
            else
            {
                Query = (from a in dbContext.BankExpenditureNotes
                         join b in dbContext.BankExpenditureNoteDetails on a.Id equals b.BankExpenditureNoteId
                         join c in dbContext.PurchasingDocumentExpeditions on b.UnitPaymentOrderId equals c.Id
                         where c.InvoiceNo == (InvoiceNo ?? c.InvoiceNo)
                            && c.SupplierCode == (SupplierCode ?? c.SupplierCode)
                            && c.UnitPaymentOrderNo == (UnitPaymentOrderNo ?? c.UnitPaymentOrderNo)
                            && c.DivisionCode == (DivisionCode ?? c.DivisionCode)
                            && !c.PaymentMethod.ToUpper().Equals("CASH")
                            && c.IsPaid
                            && c.PaymentMethod == (PaymentMethod ?? c.PaymentMethod)
                         where a.DocumentNo == (DocumentNo ?? a.DocumentNo) && a.Date.AddHours(Offset).Date >= DateFrom.Value.Date && a.Date.AddHours(Offset).Date <= DateTo.Value.Date
                         orderby a.DocumentNo
                         select new BankExpenditureNoteReportViewModel
                         {
                             DocumentNo = a.DocumentNo,
                             Currency = a.BankCurrencyCode,
                             Date = a.Date,
                             SupplierName = c.SupplierName,
                             CategoryName = c.CategoryName == null ? "-" : c.CategoryName,
                             DivisionName = c.DivisionName,
                             PaymentMethod = c.PaymentMethod,
                             UnitPaymentOrderNo = b.UnitPaymentOrderNo,
                             BankName = string.Concat(a.BankAccountName, " - ", a.BankName, " - ", a.BankAccountNumber, " - ", a.BankCurrencyCode),
                             DPP = c.TotalPaid - c.Vat,
                             VAT = c.Vat,
                             TotalPaid = c.TotalPaid,
                             InvoiceNumber = c.InvoiceNo
                         }
                      );
            }

            Pageable<BankExpenditureNoteReportViewModel> pageable = new Pageable<BankExpenditureNoteReportViewModel>(Query, Page - 1, Size);
            List<object> data = pageable.Data.ToList<object>();

            return new ReadResponse(data, pageable.TotalCount, new Dictionary<string, string>());
        }
    }
}
