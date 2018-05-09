﻿using Com.DanLiris.Service.Purchasing.Lib.Enums;
using Com.Moonlay.Models;
using System;
using System.ComponentModel.DataAnnotations;
using System.Collections.Generic;
using System.Linq;

namespace Com.DanLiris.Service.Purchasing.Lib.Models.Expedition
{
    public class PurchasingDocumentExpedition : StandardEntity, IValidatableObject
    {
        public string UnitPaymentOrderNo { get; set; }
        public string Supplier { get; set; }
        public string Division { get; set; }
        public ExpeditionPosition Position { get; set; }
        public string SendToVerificationDivisionBy { get; set; }
        public DateTimeOffset? SendToVerificationDivisionDate { get; set; }
        public string VerificationDivisionBy { get; set; }
        public DateTimeOffset? VerificationDivisionDate { get; set; }
        public string SendToCashierDivisionBy { get; set; }
        public DateTimeOffset? SendToCashierDivisionDate { get; set; }
        public string SendToFinanceDivisionBy { get; set; }
        public DateTimeOffset? SendToFinanceDivisionDate { get; set; }
        public string SendToPurchasingDivisionBy { get; set; }
        public DateTimeOffset? SendToPurchasingDivisionDate { get; set; }
        public string CashierDivisionBy { get; set; }
        public DateTimeOffset? CashierDivisionDate { get; set; }
        public string FinanceDivisionBy { get; set; }
        public DateTimeOffset? FinanceDivisionDate { get; set; }

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            PurchasingDbContext dbContext = (PurchasingDbContext)validationContext.GetService(typeof(PurchasingDbContext));

            if (dbContext.PurchasingDocumentExpeditions.Count(p => p._IsDeleted.Equals(false) && p.Id != this.Id && p.UnitPaymentOrderNo.Equals(this.UnitPaymentOrderNo)) > 0) /* Unique */
            {
                yield return new ValidationResult($"Unit Payment Order No {this.UnitPaymentOrderNo} is already exists", new List<string> { "UnitPaymentOrdersCollection" });
            }
        }
    }
}
