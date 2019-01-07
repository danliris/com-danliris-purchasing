﻿using Com.DanLiris.Service.Purchasing.Mongo.Lib.MongoModels;
using Microsoft.Extensions.Options;
using MongoDB.Driver;

namespace Com.DanLiris.Service.Purchasing.Mongo.Lib
{
    public class MongoDbMigrationContext : IMongoDbContext
    {
        private readonly IMongoDatabase _db;

        public MongoDbMigrationContext(IOptions<MongoDbSettings> options, IMongoClient client)
        {
            _db = client.GetDatabase(options.Value.Database);
        }

        public IMongoCollection<PurchaseRequestMongo> PurchaseRequests => _db.GetCollection<PurchaseRequestMongo>("purchase-requests");
    }
}
