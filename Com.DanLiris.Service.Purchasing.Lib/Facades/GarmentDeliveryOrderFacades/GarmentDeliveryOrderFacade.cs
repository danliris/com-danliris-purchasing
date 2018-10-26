﻿using Com.DanLiris.Service.Purchasing.Lib.Helpers;
using Com.DanLiris.Service.Purchasing.Lib.Interfaces;
using Com.DanLiris.Service.Purchasing.Lib.Models.GarmentDeliveryOrderModel;
using Com.DanLiris.Service.Purchasing.Lib.Utilities;
using Com.Moonlay.Models;
using Com.Moonlay.NetCore.Lib;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentDeliveryOrderFacades
{
    public class GarmentDeliveryOrderFacade : IGarmentDeliveryOrderFacade
    {
        private string USER_AGENT = "Facade";

        private readonly PurchasingDbContext dbContext;
        private readonly DbSet<GarmentDeliveryOrder> dbSet;

        public GarmentDeliveryOrderFacade(PurchasingDbContext dbContext)
        {
            this.dbContext = dbContext;
            dbSet = dbContext.Set<GarmentDeliveryOrder>();
        }

        public Tuple<List<GarmentDeliveryOrder>, int, Dictionary<string, string>> Read(int Page = 1, int Size = 25, string Order = "{}", string Keyword = null, string Filter = "{}")
        {
            IQueryable<GarmentDeliveryOrder> Query = this.dbSet.Include(m => m.Items);

            List<string> searchAttributes = new List<string>()
            {
                "DONo", "SupplierName", "Items.EPONo"
            };

            Query = QueryHelper<GarmentDeliveryOrder>.ConfigureSearch(Query, searchAttributes, Keyword);

            Dictionary<string, string> FilterDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Filter);
            Query = QueryHelper<GarmentDeliveryOrder>.ConfigureFilter(Query, FilterDictionary);

            Dictionary<string, string> OrderDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Order);
            Query = QueryHelper<GarmentDeliveryOrder>.ConfigureOrder(Query, OrderDictionary);

            Pageable<GarmentDeliveryOrder> pageable = new Pageable<GarmentDeliveryOrder>(Query, Page - 1, Size);
            List<GarmentDeliveryOrder> Data = pageable.Data.ToList();
            int TotalData = pageable.TotalCount;

            return Tuple.Create(Data, TotalData, OrderDictionary);
        }

        public GarmentDeliveryOrder ReadById(int id)
        {
            var model = dbSet.Where(m => m.Id == id)
                .Include(m => m.Items)
                    .ThenInclude(i => i.Details)
                .FirstOrDefault();
            return model;
        }

        public async Task<int> Create(GarmentDeliveryOrder m, string user, int clientTimeZoneOffset = 7)
        {
            int Created = 0;

            using (var transaction = this.dbContext.Database.BeginTransaction())
            {
                try
                {
                    EntityExtension.FlagForCreate(m, user, USER_AGENT);

                    m.IsClosed = false;

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

        public async Task<int> Update(int id, GarmentDeliveryOrder m, string user, int clientTimeZoneOffset = 7)
        {
            int Updated = 0;

            using (var transaction = dbContext.Database.BeginTransaction())
            {
                try
                {
                    var oldM = this.dbSet.AsNoTracking()
                               .SingleOrDefault(pr => pr.Id == id && !pr.IsDeleted);

                    if (oldM != null && oldM.Id == id)
                    {
                        EntityExtension.FlagForUpdate(m, user, USER_AGENT);

                        dbSet.Update(m);

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

        public async Task<int> Delete(int id, string user)
        {
            int Deleted = 0;

            using (var transaction = this.dbContext.Database.BeginTransaction())
            {
                try
                {
                    var model = this.dbSet
                                .SingleOrDefault(m => m.Id == id && !m.IsDeleted);

                    EntityExtension.FlagForDelete(model, user, USER_AGENT);

                    Deleted = await dbContext.SaveChangesAsync();

                    await dbContext.SaveChangesAsync();
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

        public IQueryable<GarmentDeliveryOrder> ReadBySupplier(string Keyword, string Filter)
        {
            IQueryable<GarmentDeliveryOrder> Query = this.dbSet;

            List<string> searchAttributes = new List<string>()
            {
                "DONo"
            };

            Query = QueryHelper<GarmentDeliveryOrder>.ConfigureSearch(Query, searchAttributes, Keyword); // kalo search setelah Select dengan .Where setelahnya maka case sensitive, kalo tanpa .Where tidak masalah
            Dictionary<string, string> FilterDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Filter);
            Query = QueryHelper<GarmentDeliveryOrder>.ConfigureFilter(Query, FilterDictionary);
            Dictionary<string, string> OrderDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>("{}");

            if (OrderDictionary.Count > 0 && OrderDictionary.Keys.First().Contains("."))
            {
                string Key = OrderDictionary.Keys.First();
                string SubKey = Key.Split(".")[1];
                string OrderType = OrderDictionary[Key];

                Query = Query.Include(m => m.Items)
                    .ThenInclude(i => i.Details);
            }
            else
            {
                Query = QueryHelper<GarmentDeliveryOrder>.ConfigureOrder(Query, OrderDictionary).Include(m => m.Items)
                    .ThenInclude(i => i.Details);
            }

            return Query;
        }
    }
}