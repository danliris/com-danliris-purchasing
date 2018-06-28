﻿using Com.DanLiris.Service.Purchasing.Lib;
using Com.DanLiris.Service.Purchasing.Lib.Facades.Expedition;
using Com.DanLiris.Service.Purchasing.Lib.Helpers.ReadResponse;
using Com.DanLiris.Service.Purchasing.Lib.Interfaces;
using Com.DanLiris.Service.Purchasing.Lib.Models.Expedition;
using Com.DanLiris.Service.Purchasing.Lib.Services;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.ExpeditionDataUtil;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.PPHBankExpenditureNoteDataUtil;
using Com.DanLiris.Service.Purchasing.Test.Helpers;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using Moq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using Xunit;

namespace Com.DanLiris.Service.Purchasing.Test.Facades.PPHBankExpenditureNoteTest
{
    public class BasicTest
    {
        private const string ENTITY = "PPHBankExpenditureNote";
        private PurchasingDocumentAcceptanceDataUtil pdaDataUtil;

        [MethodImpl(MethodImplOptions.NoInlining)]
        public string GetCurrentMethod()
        {
            StackTrace st = new StackTrace();
            StackFrame sf = st.GetFrame(1);

            return string.Concat(sf.GetMethod().Name, "_", ENTITY);
        }

        private PurchasingDbContext _dbContext(string testName)
        {
            DbContextOptionsBuilder<PurchasingDbContext> optionsBuilder = new DbContextOptionsBuilder<PurchasingDbContext>();
            optionsBuilder
                .UseInMemoryDatabase(testName)
                .ConfigureWarnings(w => w.Ignore(InMemoryEventId.TransactionIgnoredWarning));

            PurchasingDbContext dbContext = new PurchasingDbContext(optionsBuilder.Options);

            return dbContext;
        }

        private PPHBankExpenditureNoteDataUtil _dataUtil(PPHBankExpenditureNoteFacade facade, string testName)
        {
            var serviceProvider = new Mock<IServiceProvider>();
            serviceProvider
                .Setup(x => x.GetService(typeof(IdentityService)))
                .Returns(new IdentityService() { Token = "Token", Username = "Test" });

            serviceProvider
                .Setup(x => x.GetService(typeof(IHttpClientService)))
                .Returns(new HttpClientTestService());

            
            PurchasingDocumentExpeditionFacade pdeFacade = new PurchasingDocumentExpeditionFacade(serviceProvider.Object, _dbContext(testName));
            SendToVerificationDataUtil stvDataUtil = new SendToVerificationDataUtil(pdeFacade);
            pdaDataUtil = new PurchasingDocumentAcceptanceDataUtil(pdeFacade, stvDataUtil);

            return new PPHBankExpenditureNoteDataUtil(facade, pdaDataUtil);
        }

        [Fact]
        public async void Should_Success_Get_Data()
        {
            PPHBankExpenditureNoteFacade facade = new PPHBankExpenditureNoteFacade(_dbContext(GetCurrentMethod()));
            await _dataUtil(facade, GetCurrentMethod()).GetTestData();
            ReadResponse Response = facade.Read();
            Assert.NotEqual(Response.Data.Count, 0);
        }

        [Fact]
        public async void Should_Success_Get_Unit_Payment_Order()
        {
            PPHBankExpenditureNoteFacade facade = new PPHBankExpenditureNoteFacade(_dbContext(GetCurrentMethod()));
            _dataUtil(facade, GetCurrentMethod());
            PurchasingDocumentExpedition model = await pdaDataUtil.GetCashierTestData();
           
            var Response = facade.GetUnitPaymentOrder(null, null, model.IncomeTaxName, model.IncomeTaxRate, model.Currency);
            Assert.NotEqual(Response.Count, 0);
        }

        [Fact]
        public async void Should_Success_Get_Unit_Payment_Order_With_Date()
        {
            PPHBankExpenditureNoteFacade facade = new PPHBankExpenditureNoteFacade(_dbContext(GetCurrentMethod()));
            _dataUtil(facade, GetCurrentMethod());
            PurchasingDocumentExpedition model = await pdaDataUtil.GetCashierTestData();

            var Response = facade.GetUnitPaymentOrder(model.DueDate, model.DueDate, model.IncomeTaxName, model.IncomeTaxRate, model.Currency);
            Assert.NotEqual(Response.Count, 0);
        }

        [Fact]
        public async void Should_Success_Get_Data_By_Id()
        {
            PPHBankExpenditureNoteFacade facade = new PPHBankExpenditureNoteFacade(_dbContext(GetCurrentMethod()));
            PPHBankExpenditureNote model = await _dataUtil(facade, GetCurrentMethod()).GetTestData();
            var Response = facade.ReadById((int)model.Id);
            Assert.NotNull(Response);
        }

        [Fact]
        public async void Should_Success_Create_Data()
        {
            PPHBankExpenditureNoteFacade facade = new PPHBankExpenditureNoteFacade(_dbContext(GetCurrentMethod()));
            PPHBankExpenditureNote model = _dataUtil(facade, GetCurrentMethod()).GetNewData();
            var Response = await facade.Create(model, "Unit Test");
            Assert.NotEqual(Response, 0);
        }

        [Fact]
        public async void Should_Success_Update_Data()
        {
            PPHBankExpenditureNoteFacade facade = new PPHBankExpenditureNoteFacade(_dbContext(GetCurrentMethod()));
            PPHBankExpenditureNote model = await _dataUtil(facade, GetCurrentMethod()).GetTestData();

            PPHBankExpenditureNoteItem modelItem = _dataUtil(facade, GetCurrentMethod()).GetItemNewData();
            model.Items.Clear();
            model.Items.Add(modelItem);
            var Response = await facade.Update((int)model.Id, model, "Unit Test");
            Assert.NotEqual(Response, 0);
        }

        [Fact]
        public async void Should_Success_Delete_Data()
        {
            PPHBankExpenditureNoteFacade facade = new PPHBankExpenditureNoteFacade(_dbContext(GetCurrentMethod()));
            PPHBankExpenditureNote Data = await _dataUtil(facade, GetCurrentMethod()).GetTestData();
            int AffectedRows = await facade.Delete(Data.Id, "Test");
            Assert.True(AffectedRows > 0);
        }
    }
}
