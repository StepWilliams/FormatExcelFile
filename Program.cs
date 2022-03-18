using CsvHelper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FormatExcelFile
{
    public class Program
    {
        public static string location = string.Empty;
        public static string fileName = "MOCK_DATA.csv";
        public static string newFileName = "MOCK_DATA_FILTERED.csv";
        public static void Main(string[] args)
        {
            Console.WriteLine("Interview test for Stephanie Williams B.");
            Console.WriteLine("Enter file location.");
            location = Console.ReadLine();
         
            //Current CSV file
            List<CurrentExcelFile> list = ReadCSVFile(Path.Combine(location, fileName));
            //List with new CSV data 
            List<NewExcelFile> list2 = FormatExcelData(list);
            //Create new CSV file
            WriteFile(list2, newFileName);

            Console.WriteLine("New file name: MOCK_DATA_FILTERED.csv");
        }

        /// <summary>
        /// Read the current excel file and map it into CurrentExcelFile model
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>List with CurrentExcelFile data from Excel</returns>
        public static List<CurrentExcelFile> ReadCSVFile(string fileName)
        {
            try
            {
                List<CurrentExcelFile> currentExcelFiles = new List<CurrentExcelFile>();
                //Get file info
                FileInfo inputFile = new FileInfo(fileName);
                using (var reader = new StreamReader(inputFile.FullName))
                using (var csvReader = new CsvReader(reader, System.Globalization.CultureInfo.CurrentCulture))
                {
                    //Register the mapper
                    csvReader.Context.RegisterClassMap<ExcelMap>();
                    //Get excel data into CurrentExcelFile model
                    currentExcelFiles = csvReader.GetRecords<CurrentExcelFile>().ToList();
                }

                return currentExcelFiles;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// Format current excek data into a new list of NewExcelFile
        /// </summary>
        /// <param name="currentData"></param>
        /// <returns>List with NewExcelFile</returns>
        public static List<NewExcelFile> FormatExcelData(List<CurrentExcelFile> currentData)
        {
            try
            {
                List<NewExcelFile> newData = new List<NewExcelFile>();
                //read every item from current data excel
                foreach (var item in currentData)
                {   
                    //count to help to recolect data based on the number (1,2,3,4)
                    int count = 1;
                    while (count <= 4)
                    {
                        //Creates the name of the property to get in a dinamic way based on the counter
                        string name = string.Concat(nameof(NewExcelFile.Name), count);
                        //Gets the value from the specific propertie
                        string nameValue = item.GetType().GetProperty(name).GetValue(item, null).ToString();

                        string address = string.Concat(nameof(NewExcelFile.Address), count);
                        string addressValue = item.GetType().GetProperty(address).GetValue(item, null).ToString();

                        string city = string.Concat(nameof(NewExcelFile.City), count);
                        string cityValue = item.GetType().GetProperty(city).GetValue(item, null).ToString();

                        string state = string.Concat(nameof(NewExcelFile.State), count);
                        string stateValue = item.GetType().GetProperty(state).GetValue(item, null).ToString();

                        string zip = string.Concat(nameof(NewExcelFile.Zip), count);
                        string zipValue = item.GetType().GetProperty(zip).GetValue(item, null).ToString();

                        //Creates an NewExcelFile object with collected data
                        NewExcelFile data = new NewExcelFile
                        {
                            Flag1 = item.Flag1,
                            Flag2 = item.Flag2,
                            Flag3 = item.Flag3,
                            Flag4 = item.Flag4,
                            Flag5 = item.Flag5,
                            Name = nameValue,
                            Address = addressValue,
                            City = cityValue,
                            State = stateValue,
                            Zip = zipValue
                        };

                        newData.Add(data);  
                        //Moves to the next property
                        count++;
                    }
                }
                return newData;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// Create a new excel file with formated data
        /// </summary>
        /// <param name="data"></param>
        /// <param name="fileName"></param>
        public static void WriteFile(List<NewExcelFile> data, string fileName)
        {
            try
            {
                //Creates a new folder
                string path = Path.Combine(location);
                //check if folder exist, case false, it creates a new folder
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                path = Path.Combine(path, fileName);
                FileInfo toFile = new FileInfo(path);

                using (var writer = new StreamWriter(toFile.FullName))
                using (var csvWriter = new CsvWriter(writer, System.Globalization.CultureInfo.CurrentCulture))
                {
                    //maps the data into NewExcel model to create excel file
                    csvWriter.Context.RegisterClassMap<NewExcelMap>();
                    //write data into excel file
                    csvWriter.WriteRecords(data);
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }


        }

    }

}
