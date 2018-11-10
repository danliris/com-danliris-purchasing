﻿using Com.DanLiris.Service.Purchasing.Lib.Models.GarmentCorrectionNoteModel;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Com.DanLiris.Service.Purchasing.Lib.Interfaces
{
    public interface IGarmentCorrectionNoteFacade
    {
        Tuple<List<GarmentCorrectionNote>, int, Dictionary<string, string>> Read(int Page = 1, int Size = 25, string Order = "{}", string Keyword = null, string Filter = "{}");
        GarmentCorrectionNote ReadById(long id);
        Task<int> Create(GarmentCorrectionNote garmentCorrectionNote, string user, int clientTimeZoneOffset = 7);
    }
}
