using CsvHelper.Configuration;

namespace FormatExcelFile
{
    public class CurrentExcelFile
    {
        public string Flag1 { get; set; }
        public string Flag2 { get; set; }
        public string Flag3 { get; set; }
        public string Flag4 { get; set; }
        public string Flag5 { get; set; }
        public string Name1 { get; set; }
        public string Name2 { get; set; }
        public string Name3 { get; set; }
        public string Name4 { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public string Address4 { get; set; }
        public string City1 { get; set; }
        public string City2 { get; set; }
        public string City3 { get; set; }
        public string City4 { get; set; }
        public string State1 { get; set; }
        public string State2 { get; set; }
        public string State3 { get; set; }
        public string State4 { get; set; }
        public string Zip1 { get; set; }
        public string Zip2 { get; set; }
        public string Zip3 { get; set; }
        public string Zip4 { get; set; }
    }

    public class ExcelMap : ClassMap<CurrentExcelFile>
    {
        public ExcelMap()
        {
            #region Flags
            Map(l => l.Flag1).Name(nameof(CurrentExcelFile.Flag1));
            Map(l => l.Flag2).Name(nameof(CurrentExcelFile.Flag2));
            Map(l => l.Flag3).Name(nameof(CurrentExcelFile.Flag3));
            Map(l => l.Flag4).Name(nameof(CurrentExcelFile.Flag4));
            Map(l => l.Flag5).Name(nameof(CurrentExcelFile.Flag5));
            #endregion
            #region Names
            Map(l => l.Name1).Name(nameof(CurrentExcelFile.Name1));
            Map(l => l.Name2).Name(nameof(CurrentExcelFile.Name2));
            Map(l => l.Name3).Name(nameof(CurrentExcelFile.Name3));
            Map(l => l.Name4).Name(nameof(CurrentExcelFile.Name4));
            #endregion
            #region Addreess
            Map(l => l.Address1).Name(nameof(CurrentExcelFile.Address1));
            Map(l => l.Address2).Name(nameof(CurrentExcelFile.Address2));
            Map(l => l.Address3).Name(nameof(CurrentExcelFile.Address3));
            Map(l => l.Address4).Name(nameof(CurrentExcelFile.Address4));
            #endregion 
            #region Cities
            Map(l => l.City1).Name(nameof(CurrentExcelFile.City1));
            Map(l => l.City2).Name(nameof(CurrentExcelFile.City2));
            Map(l => l.City3).Name(nameof(CurrentExcelFile.City3));
            Map(l => l.City4).Name(nameof(CurrentExcelFile.City4));
            #endregion    
            #region States
            Map(l => l.State1).Name(nameof(CurrentExcelFile.State1));
            Map(l => l.State2).Name(nameof(CurrentExcelFile.State2));
            Map(l => l.State3).Name(nameof(CurrentExcelFile.State3));
            Map(l => l.State4).Name(nameof(CurrentExcelFile.State4));
            #endregion   
            #region ZIP
            Map(l => l.Zip1).Name(nameof(CurrentExcelFile.Zip1));
            Map(l => l.Zip2).Name(nameof(CurrentExcelFile.Zip2));
            Map(l => l.Zip3).Name(nameof(CurrentExcelFile.Zip3));
            Map(l => l.Zip4).Name(nameof(CurrentExcelFile.Zip4));
            #endregion
        }
    }
}
