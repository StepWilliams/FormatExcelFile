using CsvHelper.Configuration;

namespace FormatExcelFile
{
    public class NewExcelFile
    {
        public string Flag1 { get; set; }
        public string Flag2 { get; set; }
        public string Flag3 { get; set; }
        public string Flag4 { get; set; }
        public string Flag5 { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Zip { get; set; }
    }
    public class NewExcelMap : ClassMap<NewExcelFile>
    {
        public NewExcelMap()
        {
            #region Flags
            Map(l => l.Flag1).Name(nameof(NewExcelFile.Flag1));
            Map(l => l.Flag2).Name(nameof(NewExcelFile.Flag2));
            Map(l => l.Flag3).Name(nameof(NewExcelFile.Flag3));
            Map(l => l.Flag4).Name(nameof(NewExcelFile.Flag4));
            Map(l => l.Flag5).Name(nameof(NewExcelFile.Flag5));
            #endregion
            #region Names
            Map(l => l.Name).Name(nameof(NewExcelFile.Name));
            #endregion
            #region Addreess
            Map(l => l.Address).Name(nameof(NewExcelFile.Address));
            #endregion
            #region Cities
            Map(l => l.City).Name(nameof(NewExcelFile.City));
            #endregion
            #region States
            Map(l => l.State).Name(nameof(NewExcelFile.State));
            #endregion
            #region ZIP
            Map(l => l.Zip).Name(nameof(NewExcelFile.Zip));
            #endregion
        }
    }

}
