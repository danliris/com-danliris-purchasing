﻿using Com.DanLiris.Service.Purchasing.Lib.Configs.Expedition;
using Com.DanLiris.Service.Purchasing.Lib.Models.Expedition;
using Com.Moonlay.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;

namespace Com.DanLiris.Service.Purchasing.Lib
{
    public class PurchasingDbContext : BaseDbContext
    {
        public PurchasingDbContext(DbContextOptions<PurchasingDbContext> options) : base(options)
        {
        }

        public DbSet<PurchasingDocumentExpedition> PurchasingDocumentExpeditions { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.ApplyConfiguration(new PurchasingDocumentExpeditionConfig());
        }
    }
}
