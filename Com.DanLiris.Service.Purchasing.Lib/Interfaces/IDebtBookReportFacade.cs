﻿
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.DebtBookReportViewModels;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.GarmentReports;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace Com.DanLiris.Service.Purchasing.Lib.Interfaces
{
    public interface IDebtBookReportFacade 
    {
        Tuple<List<DebtBookReportViewModel>, int> GetDebtBookReport(int month, int year, bool? suppliertype, string suppliername);
        MemoryStream GenerateExcelDebtReport(int month, int year, bool? suppliertype, string suppliername);
        
    }
}
