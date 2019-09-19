﻿
using Com.DanLiris.Service.Purchasing.Lib.Helpers;
using Com.DanLiris.Service.Purchasing.Lib.Interfaces;
using Com.DanLiris.Service.Purchasing.Lib.Models.GarmentDeliveryOrderModel;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.GarmentDailyPurchasingReportViewModel;
using Com.Moonlay.NetCore.Lib;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace Com.DanLiris.Service.Purchasing.Lib.Facades.GarmentDailyPurchasingReportFacade
{
    public class GarmentDailyPurchasingReportFacade : IGarmentDailyPurchasingReportFacade
    {
        private readonly PurchasingDbContext dbContext;
        public readonly IServiceProvider serviceProvider;
        private readonly DbSet<GarmentDeliveryOrder> dbSet;

        public GarmentDailyPurchasingReportFacade(IServiceProvider serviceProvider, PurchasingDbContext dbContext)
        {
            this.serviceProvider = serviceProvider;
            this.dbContext = dbContext;
            this.dbSet = dbContext.Set<GarmentDeliveryOrder>();
        }
        #region GarmentDailyPurchasingAll
        public IEnumerable<GarmentDailyPurchasingReportViewModel> GetGarmentDailyPurchasingReportQuery(string unitName, bool supplierType, string supplierName, DateTime? dateFrom, DateTime? dateTo, int offset)
        {
            DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : (DateTime)dateFrom;
            DateTime DateTo = dateTo == null ? DateTime.Now : (DateTime)dateTo;

            IQueryable<GarmentDailyPurchasingTempViewModel> d1 = from a in dbContext.GarmentDeliveryOrders
                                                               join b in dbContext.GarmentDeliveryOrderItems on a.Id equals b.GarmentDOId
                                                               join c in dbContext.GarmentDeliveryOrderDetails on b.Id equals c.GarmentDOItemId
                                                               join d in dbContext.GarmentBeacukais on a.CustomsId equals d.Id
                                                               join e in dbContext.GarmentExternalPurchaseOrders on b.EPOId equals e.Id
                                                               join f in dbContext.GarmentInternalPurchaseOrders on c.POId equals f.Id
                                                                 where c.DOQuantity != 0
                                                                 && c.UnitId == (string.IsNullOrWhiteSpace(unitName) ? c.UnitId : unitName)
                                                                 && e.SupplierImport == supplierType
                                                                 && (string.IsNullOrWhiteSpace(supplierName) ? true : (supplierName == "DAN LIRIS" ? a.SupplierCode.Substring(0, 2) == "DL" : a.SupplierCode.Substring(0, 2) != "DL"))
                                                                 && d.BeacukaiDate.AddHours(offset).Date >= DateFrom.Date && d.BeacukaiDate.AddHours(offset).Date <= DateTo.Date

                                                                 select new GarmentDailyPurchasingTempViewModel
                                                               {
                                                                   SuplName = a.SupplierName,
                                                                   UnitName = f.UnitName,
                                                                   BCNo = a.BillNo,
                                                                   BonKecil = a.PaymentBill,
                                                                   DONo = a.DONo,
                                                                   INNo = a.InternNo,
                                                                   ProductName = c.ProductName,
                                                                   JnsBrg = c.CodeRequirment,
                                                                   Quantity = (decimal)c.DOQuantity,
                                                                   Satuan = c.UomUnit,
                                                                   Kurs = (double)a.DOCurrencyRate,
                                                                   Amount = c.PriceTotal,
                                                               };

            IQueryable<GarmentDailyPurchasingTempViewModel> d2 = from gc in dbContext.GarmentCorrectionNotes
                                                               join gci in dbContext.GarmentCorrectionNoteItems on gc.Id equals gci.GCorrectionId
                                                               join ipo in dbContext.GarmentInternalPurchaseOrders on gci.POId equals ipo.Id
                                                               join gdd in dbContext.GarmentDeliveryOrderDetails on gci.DODetailId equals gdd.Id
                                                               join gdi in dbContext.GarmentDeliveryOrderItems on gdd.GarmentDOItemId equals gdi.Id
                                                               join gdo in dbContext.GarmentDeliveryOrders on gdi.GarmentDOId equals gdo.Id
                                                               join epo in dbContext.GarmentExternalPurchaseOrders on gci.EPOId equals epo.Id
                                                               where gci.Quantity != 0
                                                               && ipo.UnitId == (string.IsNullOrWhiteSpace(unitName) ? ipo.UnitId : unitName)
                                                               && epo.SupplierImport == supplierType
                                                               && (string.IsNullOrWhiteSpace(supplierName) ? true : (supplierName == "DAN LIRIS" ? gc.SupplierCode.Substring(0, 2) == "DL" : gc.SupplierCode.Substring(0, 2) != "DL"))
                                                               && gc.CorrectionDate.AddHours(offset).Date >= DateFrom.Date && gc.CorrectionDate.AddHours(offset).Date <= DateTo.Date

                                                               select new GarmentDailyPurchasingTempViewModel
                                                               {                               
                                                                   SuplName = gc.SupplierName,
                                                                   UnitName = ipo.UnitName,
                                                                   BCNo = gc.CorrectionNo,
                                                                   BonKecil = gdo.PaymentBill,
                                                                   DONo = gc.DONo,
                                                                   INNo = gdo.InternNo,
                                                                   ProductName = gci.ProductName,
                                                                   JnsBrg = gdd.CodeRequirment,
                                                                   Quantity = (decimal)gci.Quantity,
                                                                   Satuan = gci.UomIUnit,
                                                                   Kurs = (double)gdo.DOCurrencyRate,
                                                                   Amount = (double)(gci.PriceTotalAfter - gci.PriceTotalBefore),
                                                               };

            IQueryable<GarmentDailyPurchasingTempViewModel> d3 = from gc in dbContext.GarmentCorrectionNotes
                                                               join gci in dbContext.GarmentCorrectionNoteItems on gc.Id equals gci.GCorrectionId
                                                               join ipo in dbContext.GarmentInternalPurchaseOrders on gci.POId equals ipo.Id
                                                               join gdd in dbContext.GarmentDeliveryOrderDetails on gci.DODetailId equals gdd.Id
                                                               join gdi in dbContext.GarmentDeliveryOrderItems on gdd.GarmentDOItemId equals gdi.Id
                                                               join gdo in dbContext.GarmentDeliveryOrders on gdi.GarmentDOId equals gdo.Id
                                                               join epo in dbContext.GarmentExternalPurchaseOrders on gci.EPOId equals epo.Id
                                                               where gci.Quantity != 0
                                                               && ipo.UnitId == (string.IsNullOrWhiteSpace(unitName) ? ipo.UnitId : unitName)
                                                               && epo.SupplierImport == supplierType
                                                               && gc.UseVat == true
                                                               && gc.NKPN != null 
                                                               && (string.IsNullOrWhiteSpace(supplierName) ? true : (supplierName == "DAN LIRIS" ? gc.SupplierCode.Substring(0, 2) == "DL" : gc.SupplierCode.Substring(0, 2) != "DL"))
                                                               && gc.CorrectionDate.AddHours(offset).Date >= DateFrom.Date && gc.CorrectionDate.AddHours(offset).Date <= DateTo.Date

                                                               select new GarmentDailyPurchasingTempViewModel
                                                               {
                                                                   SuplName = gc.SupplierName,
                                                                   UnitName = ipo.UnitName,
                                                                   BCNo = gc.NKPN,
                                                                   BonKecil = gdo.PaymentBill,
                                                                   DONo = gc.DONo,
                                                                   INNo = gdo.InternNo,
                                                                   ProductName = gci.ProductName,
                                                                   JnsBrg = gdd.CodeRequirment,
                                                                   Quantity = (decimal)gci.Quantity,
                                                                   Satuan = gci.UomIUnit,
                                                                   Kurs = (double)gdo.DOCurrencyRate,
                                                                   Amount = (double)(gci.PriceTotalAfter - gci.PriceTotalBefore) * 10/100,
                                                               };

            IQueryable<GarmentDailyPurchasingTempViewModel> d4 = from gc in dbContext.GarmentCorrectionNotes
                                                               join gci in dbContext.GarmentCorrectionNoteItems on gc.Id equals gci.GCorrectionId
                                                               join ipo in dbContext.GarmentInternalPurchaseOrders on gci.POId equals ipo.Id
                                                               join gdd in dbContext.GarmentDeliveryOrderDetails on gci.DODetailId equals gdd.Id
                                                               join gdi in dbContext.GarmentDeliveryOrderItems on gdd.GarmentDOItemId equals gdi.Id
                                                               join gdo in dbContext.GarmentDeliveryOrders on gdi.GarmentDOId equals gdo.Id
                                                               join epo in dbContext.GarmentExternalPurchaseOrders on gci.EPOId equals epo.Id
                                                               where gci.Quantity != 0
                                                               && ipo.UnitId == (string.IsNullOrWhiteSpace(unitName) ? ipo.UnitId : unitName)
                                                               && epo.SupplierImport == supplierType
                                                               && gc.UseIncomeTax == true
                                                               && gc.NKPH != null
                                                               && (string.IsNullOrWhiteSpace(supplierName) ? true : (supplierName == "DAN LIRIS" ? gc.SupplierCode.Substring(0, 2) == "DL" : gc.SupplierCode.Substring(0, 2) != "DL"))
                                                               && gc.CorrectionDate.AddHours(offset).Date >= DateFrom.Date && gc.CorrectionDate.AddHours(offset).Date <= DateTo.Date

                                                               select new GarmentDailyPurchasingTempViewModel
                                                               {
                                                                   SuplName = gc.SupplierName,
                                                                   UnitName = ipo.UnitName,
                                                                   BCNo = gc.NKPH,
                                                                   BonKecil = gdo.PaymentBill,
                                                                   DONo = gc.DONo,
                                                                   INNo = gdo.InternNo,
                                                                   ProductName = gci.ProductName,
                                                                   JnsBrg = gdd.CodeRequirment,
                                                                   Quantity = (decimal)gci.Quantity,
                                                                   Satuan = gci.UomIUnit,
                                                                   Kurs = (double)gdo.DOCurrencyRate,
                                                                   Amount = (double)(gci.PriceTotalAfter - gci.PriceTotalBefore) * ((double)gc.IncomeTaxRate/100),
                                                               };

            IQueryable<GarmentDailyPurchasingTempViewModel> d5 = from inv in dbContext.GarmentInvoices
                                                                 join invi in dbContext.GarmentInvoiceItems on inv.Id equals invi.InvoiceId
                                                                 join invd in dbContext.GarmentInvoiceDetails on invi.Id equals invd.InvoiceItemId
                                                                 join gdd in dbContext.GarmentDeliveryOrderDetails on invd.DODetailId equals gdd.Id
                                                                 join gdi in dbContext.GarmentDeliveryOrderItems on gdd.GarmentDOItemId equals gdi.Id
                                                                 join gdo in dbContext.GarmentDeliveryOrders on  gdi.GarmentDOId equals gdo.Id
                                                                 join epo in dbContext.GarmentExternalPurchaseOrders on invd.EPOId equals epo.Id
                                                                 join ipo in dbContext.GarmentInternalPurchaseOrders on invd.IPOId equals ipo.Id
                                                                 where invd.DOQuantity != 0
                                                                 && ipo.UnitId == (string.IsNullOrWhiteSpace(unitName) ? ipo.UnitId : unitName)
                                                                 && epo.SupplierImport == supplierType
                                                                 && inv.IsPayTax == true
                                                                 && inv.UseVat == true
                                                                 && inv.NPN != null
                                                                 && (string.IsNullOrWhiteSpace(supplierName) ? true : (supplierName == "DAN LIRIS" ? inv.SupplierCode.Substring(0, 2) == "DL" : inv.SupplierCode.Substring(0, 2) != "DL"))
                                                                 && inv.InvoiceDate.AddHours(offset).Date >= DateFrom.Date && inv.InvoiceDate.AddHours(offset).Date <= DateTo.Date

                                                                 select new GarmentDailyPurchasingTempViewModel
                                                               {
                                                                   SuplName = inv.SupplierName,
                                                                   UnitName = ipo.UnitName,
                                                                   BCNo = inv.NPN,
                                                                   BonKecil = gdo.PaymentBill,
                                                                   DONo = gdo.DONo,
                                                                   INNo = gdo.InternNo,
                                                                   ProductName = invd.ProductName,
                                                                   JnsBrg = gdd.CodeRequirment,
                                                                   Quantity = (decimal)invd.DOQuantity,
                                                                   Satuan = invd.UomUnit,
                                                                   Kurs = (double)gdo.DOCurrencyRate,
                                                                   Amount = invd.DOQuantity * invd.PricePerDealUnit * 10 / 100,
                                                               };
            IQueryable<GarmentDailyPurchasingTempViewModel> d6 = from inv in dbContext.GarmentInvoices
                                                               join invi in dbContext.GarmentInvoiceItems on inv.Id equals invi.InvoiceId
                                                               join invd in dbContext.GarmentInvoiceDetails on invi.Id equals invd.InvoiceItemId
                                                               join gdd in dbContext.GarmentDeliveryOrderDetails on invd.DODetailId equals gdd.Id
                                                               join gdi in dbContext.GarmentDeliveryOrderItems on gdd.GarmentDOItemId equals gdi.Id
                                                               join gdo in dbContext.GarmentDeliveryOrders on gdi.GarmentDOId equals gdo.Id
                                                               join epo in dbContext.GarmentExternalPurchaseOrders on invd.EPOId equals epo.Id
                                                               join ipo in dbContext.GarmentInternalPurchaseOrders on invd.IPOId equals ipo.Id
                                                               where invd.DOQuantity != 0
                                                               && ipo.UnitId == (string.IsNullOrWhiteSpace(unitName) ? ipo.UnitId : unitName)
                                                               && epo.SupplierImport == supplierType
                                                               && inv.IsPayTax == true
                                                               && inv.UseIncomeTax == true
                                                               && inv.NPH != null
                                                               && (string.IsNullOrWhiteSpace(supplierName) ? true : (supplierName == "DAN LIRIS" ? inv.SupplierCode.Substring(0, 2) == "DL" : inv.SupplierCode.Substring(0, 2) != "DL"))
                                                               && inv.InvoiceDate.AddHours(offset).Date >= DateFrom.Date && inv.InvoiceDate.AddHours(offset).Date <= DateTo.Date

                                                               select new GarmentDailyPurchasingTempViewModel
                                                               {
                                                                   SuplName = inv.SupplierName,
                                                                   UnitName = ipo.UnitName,
                                                                   BCNo = inv.NPH,
                                                                   BonKecil = gdo.PaymentBill,
                                                                   DONo = gdo.DONo,
                                                                   INNo = gdo.InternNo,
                                                                   ProductName = invd.ProductName,
                                                                   JnsBrg = gdd.CodeRequirment,
                                                                   Quantity = (decimal)invd.DOQuantity,
                                                                   Satuan = invd.UomUnit,
                                                                   Kurs = (double)gdo.DOCurrencyRate,
                                                                   Amount = invd.DOQuantity * invd.PricePerDealUnit * inv.IncomeTaxRate / 100,
                                                               };
            List<GarmentDailyPurchasingTempViewModel> CombineData = d1.Union(d2).Union(d3).Union(d4).Union(d5).Union(d6).ToList();
            
            var Query = from data in CombineData
                        group data by new { data.SuplName, data.BCNo, data.BonKecil, data.DONo, data.INNo, data.UnitName, data.ProductName, data.Satuan, data.JnsBrg } into groupData
                        select new GarmentDailyPurchasingReportViewModel 
                        {
                            SupplierName = groupData.Key.SuplName,
                            UnitName = groupData.Key.UnitName,
                            BillNo = groupData.Key.BCNo,
                            PaymentBill = groupData.Key.BonKecil,
                            DONo = groupData.Key.DONo,
                            InternNo = groupData.Key.INNo,
                            ProductName = groupData.Key.ProductName,
                            CodeRequirement = groupData.Key.JnsBrg,
                            UOMUnit = groupData.Key.Satuan,
                            Quantity = groupData.Sum(s => (double)s.Quantity),
                            Amount = Math.Round(groupData.Sum(s =>s.Amount*s.Kurs) ,2),
                        };
            return Query.AsQueryable();
        }

        public Tuple<List<GarmentDailyPurchasingReportViewModel>, int> GetGDailyPurchasingReport(string unitName, bool supplierType, string supplierName, DateTime? dateFrom, DateTime? dateTo, int offset)
        {
            List<GarmentDailyPurchasingReportViewModel> result = GetGarmentDailyPurchasingReportQuery(unitName, supplierType, supplierName, dateFrom, dateTo, offset).ToList();
            return Tuple.Create(result, result.Count);
        }
               
        public MemoryStream GenerateExcelGDailyPurchasingReport(string unitName, bool supplierType, string supplierName, DateTime? dateFrom, DateTime? dateTo, int offset)
          {
            Tuple<List<GarmentDailyPurchasingReportViewModel>, int> Data = this.GetGDailyPurchasingReport(unitName, supplierType, supplierName, dateFrom, dateTo, offset);

            DataTable result = new DataTable();
            result.Columns.Add(new DataColumn() { ColumnName = "Nomor", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Supplier", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Nama Unit", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Nomor Nota", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Nomor Bon Kecil", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Nomor Surat Jalan", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Nota Intern", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Nama Barang", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Jumlah", DataType = typeof(string) });
            result.Columns.Add(new DataColumn() { ColumnName = "Satuan", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Bahan Embalase (Rp)", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Bahan Pendukung (Rp)", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Bahan Baku (Rp)", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Proses (Rp)", DataType = typeof(String) });

            List<(string, Enum, Enum)> mergeCells = new List<(string, Enum, Enum)>() { };

            if (Data.Item2 == 0)
            {
                result.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "", "", ""); // to allow column name to be generated properly for empty data as template
            }
            else
            {
                Dictionary<string, List<GarmentDailyPurchasingReportViewModel>> dataBySupplier = new Dictionary<string, List<GarmentDailyPurchasingReportViewModel>>();
                Dictionary<string, double> subTotalBESupplier = new Dictionary<string, double>();
                Dictionary<string, double> subTotalBPSupplier = new Dictionary<string, double>();
                Dictionary<string, double> subTotalBBSupplier = new Dictionary<string, double>();
                Dictionary<string, double> subTotalPRCSupplier = new Dictionary<string, double>();

                foreach (GarmentDailyPurchasingReportViewModel data in Data.Item1)
                {       
                    string SupplierName = data.SupplierName;
                    double Amount1 = 0, Amount2 = 0, Amount3 = 0, Amount4 = 0;

                    switch (data.CodeRequirement)
                    {
                        case "BE":
                            Amount1 = data.Amount;
                            Amount2 = 0;
                            Amount3 = 0;
                            Amount4 = 0;
                            break;
                        case "BP":
                            Amount1 = 0;
                            Amount2 = data.Amount;
                            Amount3 = 0;
                            Amount4 = 0;
                            break;

                        case "BB":
                            Amount1 = 0;
                            Amount2 = 0;
                            Amount3 = data.Amount;
                            Amount4 = 0;
                            break;
                        default:
                            Amount1 = 0;
                            Amount2 = 0;
                            Amount3 = 0;
                            Amount4 = data.Amount;
                            break;
                    }

                    if (!dataBySupplier.ContainsKey(SupplierName)) dataBySupplier.Add(SupplierName, new List<GarmentDailyPurchasingReportViewModel> { });
                    dataBySupplier[SupplierName].Add(new GarmentDailyPurchasingReportViewModel
                    {

                        SupplierName = data.SupplierName,
                        UnitName = data.UnitName,
                        BillNo = data.BillNo,
                        PaymentBill = data.PaymentBill,
                        DONo = data.DONo,
                        InternNo = data.InternNo,
                        ProductName = data.ProductName,
                        CodeRequirement = data.CodeRequirement,
                        UOMUnit = data.UOMUnit,
                        Quantity = data.Quantity,
                        Amount = Amount1,
                        Amount1 = Amount2,
                        Amount2 = Amount3,
                        Amount3 = Amount4,
                    });

                    if (!subTotalBESupplier.ContainsKey(SupplierName))
                    {
                        subTotalBESupplier.Add(SupplierName, 0);
                    };

                    if (!subTotalBPSupplier.ContainsKey(SupplierName))
                    {
                        subTotalBPSupplier.Add(SupplierName, 0);
                    };

                    if (!subTotalBBSupplier.ContainsKey(SupplierName))
                    {
                        subTotalBBSupplier.Add(SupplierName, 0);
                    };

                    if (!subTotalPRCSupplier.ContainsKey(SupplierName))
                    {
                        subTotalPRCSupplier.Add(SupplierName, 0);
                    };

                    subTotalBESupplier[SupplierName] += Amount1;
                    subTotalBPSupplier[SupplierName] += Amount2;
                    subTotalBBSupplier[SupplierName] += Amount3;
                    subTotalPRCSupplier[SupplierName] += Amount4;
                }

                double totalBE = 0;
                double totalBP = 0;
                double totalBB = 0;
                double totalPRC = 0;

                int rowPosition = 1;

                foreach (KeyValuePair<string, List<GarmentDailyPurchasingReportViewModel>> SupplName in dataBySupplier)
                {
                    string splCode = "";
                    int index = 0;
                    foreach (GarmentDailyPurchasingReportViewModel data in SupplName.Value)
                    {
                        index++;
                        result.Rows.Add(index, data.SupplierName, data.UnitName, data.BillNo, data.PaymentBill, data.DONo, data.InternNo, data.ProductName, data.Quantity, data.UOMUnit, Math.Round(data.Amount, 2), Math.Round(data.Amount1, 2), Math.Round(data.Amount2, 2), Math.Round(data.Amount3, 2));
                        rowPosition += 1;
                        splCode = data.SupplierName;
                    }

                    result.Rows.Add("SUB TOTAL", "", "", "", "", "", "", "", splCode, "", Math.Round(subTotalBESupplier[SupplName.Key], 2), Math.Round(subTotalBPSupplier[SupplName.Key], 2), Math.Round(subTotalBBSupplier[SupplName.Key], 2), Math.Round(subTotalPRCSupplier[SupplName.Key], 2));
                   
                    rowPosition += 1;
                    mergeCells.Add(($"A{rowPosition}:D{rowPosition}", OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom));

                    totalBE += subTotalBESupplier[SupplName.Key];
                    totalBP += subTotalBPSupplier[SupplName.Key];
                    totalBB += subTotalBBSupplier[SupplName.Key];
                    totalPRC += subTotalPRCSupplier[SupplName.Key];
                }

                result.Rows.Add("TOTAL", "", "", "", "", "", "", "", "", "", Math.Round(totalBE, 2), Math.Round(totalBP, 2), Math.Round(totalBB, 2), Math.Round(totalPRC, 2));
            
                rowPosition += 1;
                 mergeCells.Add(($"A{rowPosition}:D{rowPosition}", OfficeOpenXml.Style.ExcelHorizontalAlignment.Right, OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom));
            }

            return Excel.CreateExcel(new List<(DataTable, string, List<(string, Enum, Enum)>)>() { (result, "Report", mergeCells) }, true);
        }
        #endregion
    }
}
