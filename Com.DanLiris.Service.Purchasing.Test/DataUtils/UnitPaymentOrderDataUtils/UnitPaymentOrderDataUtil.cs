﻿using Com.DanLiris.Service.Purchasing.Lib.Facades;
using Com.DanLiris.Service.Purchasing.Lib.Models.UnitPaymentOrderModel;
using Com.DanLiris.Service.Purchasing.Test.DataUtils.UnitReceiptNoteDataUtils;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Com.DanLiris.Service.Purchasing.Test.DataUtils.UnitPaymentOrderDataUtils
{
    public class UnitPaymentOrderDataUtil
    {
        private UnitReceiptNoteDataUtil unitReceiptNoteDataUtil;
        private readonly UnitPaymentOrderFacade facade;

        public UnitPaymentOrderDataUtil(UnitReceiptNoteDataUtil unitReceiptNoteDataUtil, UnitPaymentOrderFacade facade)
        {
            this.unitReceiptNoteDataUtil = unitReceiptNoteDataUtil;
            this.facade = facade;
        }

        public UnitPaymentOrder GetNewData()
        {
            Lib.Models.UnitReceiptNoteModel.UnitReceiptNote unitReceiptNote = Task.Run(() => this.unitReceiptNoteDataUtil.GetTestData("Unit Test")).Result;

            List<UnitPaymentOrderDetail> unitPaymentOrderDetails = new List<UnitPaymentOrderDetail>();
            foreach (var item in unitReceiptNote.Items)
            {
                unitPaymentOrderDetails.Add(new UnitPaymentOrderDetail
                {
                    URNItemId = item.Id,

                    EPODetailId = item.EPODetailId,
                    PRId = item.PRId,
                    PRNo = item.PRNo,
                    PRItemId = item.PRItemId,

                    ProductId = item.ProductId,
                    ProductCode = item.ProductCode,
                    ProductName = item.ProductName,

                    ReceiptQuantity = item.ReceiptQuantity,

                    UomId = item.UomId,
                    UomUnit = item.Uom,

                    PricePerDealUnit = item.PricePerDealUnit,
                    PriceTotal = item.PricePerDealUnit * item.ReceiptQuantity,
                    QuantityCorrection= item.ReceiptQuantity,

                    ProductRemark = item.ProductRemark
                });
            }

            List<UnitPaymentOrderItem> unitPaymentOrderItems = new List<UnitPaymentOrderItem>
            {
                new UnitPaymentOrderItem
                {
                    URNId = unitReceiptNote.Id,
                    URNNo = unitReceiptNote.URNNo,

                    DOId = unitReceiptNote.DOId,
                    DONo = unitReceiptNote.DONo,
                    Details = unitPaymentOrderDetails
                }
            };

            UnitPaymentOrder unitPaymentOrder = new UnitPaymentOrder
            {
                DivisionId = "DivisionId",
                DivisionCode = "DivisionCode",
                DivisionName = "DivisionName",

                SupplierId = "SupplierId",
                SupplierCode = "SupplierCode",
                SupplierName = "SupplierName",

                Date = new DateTimeOffset(),

                CategoryId = "CategoryId ",
                CategoryCode = "CategoryCode",
                CategoryName = "CategoryName",

                CurrencyId = "CurrencyId",
                CurrencyCode = "CurrencyCode",
                CurrencyRate = 5,

                PaymentMethod = "KREDIT",

                InvoiceNo = "INV000111",
                InvoiceDate = new DateTimeOffset(),
                PibNo = null,

                UseIncomeTax = false,
                IncomeTaxId = null,
                IncomeTaxName = null,
                IncomeTaxRate = 0,
                IncomeTaxNo = null,
                IncomeTaxDate = null,

                UseVat = false,
                VatNo = null,
                VatDate = new DateTimeOffset(),
                Position = 1,
                Remark = null,

                DueDate = new DateTimeOffset(), // ???

                Items = unitPaymentOrderItems
            };
            return unitPaymentOrder;
        }

        public async Task<UnitPaymentOrder> GetTestData()
        {
            var data = GetNewData();
            await facade.Create(data, "Unit Test", false);
            return data;
        }
    }
}
