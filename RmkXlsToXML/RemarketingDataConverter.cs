using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using ExcelDataReader;
using Serilog;


// ReSharper disable CommentTypo
// ReSharper disable StringLiteralTypo
// ReSharper disable IdentifierTypo


namespace RmkXlsToXml
{
    /// <summary>
    /// Converts the incoming Remarketing XLS file to the output xml format. 
    /// </summary>
    public class RemarketingDataConverter
    {
        private readonly ILogger _logger;

        public RemarketingDataConverter(ILogger logger)
        {
            _logger = logger;
        }
        public bool ConvertRemarketingFile(ConverterConfiguration config)
        {
            //read the data from the source xls file to a list of strongly typed data.
            var data = ReadRmkDataFromSourceFile(config);
            // make sure we've got some data before continuing
            if (data == null)
            {
                _logger.Error("Error reading data from file.");
                return false;
            }

            if (!data.Any())
            {
                _logger.Error("No Remarketing data found in file.");
                return false;
            }

            // write the data to an Xml file
            WriteDataToXmlFile(config, data);
            return true;
        }

        private static string ComposeOutputFileName(ConverterConfiguration config)
        {
            // the output file will be the same name as the source file, but with an .xml file extension
            var sourceFileName = Path.GetFileNameWithoutExtension(config.SourceFile);
            var fullOutputFileName = Path.Combine(config.OutputPath, $"{sourceFileName}.xml");
            return fullOutputFileName;
        }
        /// <summary>
        /// Writes the Remarketing data to an xml file.
        /// </summary>
        private void WriteDataToXmlFile(ConverterConfiguration config, List<RemarketingData> data)
        {
            XmlWriterSettings settings = new XmlWriterSettings {Indent = true};
            var outputFileName = ComposeOutputFileName(config);
            if (File.Exists(outputFileName))
            {
                _logger.Warning($"Overwriting existing output file: {outputFileName}");
            }

            using (XmlWriter writer = XmlWriter.Create(outputFileName, settings))
            {
                writer.WriteStartElement("Remarketing"); // root node
                writer.WriteStartElement("FileInfo"); // FileInfo node
                writer.WriteElementString("RSAClientID", config.RsaClientId);
                writer.WriteElementString("FileCreateDate", DateTime.Now.ToShortDateString());
                writer.WriteElementString("ItemCount", data.Count.ToString());
                writer.WriteEndElement(); // close FileInfoNode

                writer.WriteStartElement("RemarketingAssignmentList"); // starts the list of all the Remarketing assignments.
                foreach (var item in data)
                {
                    writer.WriteStartElement("RemarketingAssignment");
                    writer.WriteElementString("VIN",item.Vin);
                    writer.WriteElementString("AccountNumber", item.AccountNumber);
                    writer.WriteElementString("Year",item.Year);
                    writer.WriteElementString("Make",item.Make);
                    writer.WriteElementString("Model",item.Model);
                    writer.WriteElementString("Mileage",item.Mileage.ToString());
                    writer.WriteElementString("RepoDate",item.DateOfRepo.ToShortDateString());
                    writer.WriteElementString("ClearDate",item.DateOfClear.ToShortDateString());
                    writer.WriteElementString("LoanBalanceAmt",item.Balance.ToString(CultureInfo.InvariantCulture));
                    writer.WriteStartElement("VehicleLocationInfo");
                    writer.WriteElementString("IsVehicleAtCustomerSite","N");
                    writer.WriteElementString("LocationName",item.LocationOfUnit);
                    writer.WriteEndElement(); // closes VehicleLocationInfo
                    writer.WriteStartElement("CustomerInfo");
                    writer.WriteElementString("FullName", item.FullName);
                    writer.WriteEndElement(); // closes CustomerInfo
                    writer.WriteEndElement(); // closes RemarketingAssignment
                }
                writer.WriteEndElement(); // closes RemarketingAssignmentList
                writer.WriteEndElement(); // closes Remarketing
                writer.Flush();
            }
            _logger.Information($"Data has been written to: {outputFileName}");
        }
       
      

        public List<RemarketingData> ReadRmkDataFromSourceFile(ConverterConfiguration config)
        {
            var fileData = new List<RemarketingData>();

            // The incoming file is a BIFF 5.0 version of the excel file (Excel 2010)
            // so we have to use ExcelReader to read the contents.

            // running ExcelDataReading on .NET Core requires we specify the encoding provider
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // the easiest way to access the data is to read the entire workbook into a data set
            // and then parse the dataset data 
            DataSet ds;
            using (var stream = File.Open(config.SourceFile, FileMode.Open, FileAccess.Read))
            {
                using var reader = ExcelReaderFactory.CreateReader(stream);
                ds= reader.AsDataSet();
            }

            // each worksheet in the excel file is a data table in the data set
            var sheet = ds.Tables[0];
            int rowIndex = -1;
            foreach(DataRow row in sheet.Rows)
            {
                // track the row number so we can skip the first N header rows
                rowIndex++;
                if (rowIndex < config.NumberOfHeaderRows) continue;
                
                var items = row.ItemArray;
                // convert the data table data row to a strongly typed row of data.
                // ReSharper disable once UseObjectOrCollectionInitializer
                var data = new RemarketingData();

                data.AccountNumber = items[0].ToString();
                // if we encounter a missing account number, we'll use that to signify the end of the data rows.
                if (string.IsNullOrEmpty(data.AccountNumber)) break;

                data.LoanNumber = items[1].ToString();
                data.LastName = items[2].ToString();
                data.FirstName = items[3].ToString();
                data.Balance = Convert.ToDecimal(items[4]);
                data.Year = items[5].ToString();
                data.Make = items[6].ToString();
                data.Model = items[7].ToString();
                data.Vin = items[8].ToString();
                data.Mileage = Convert.ToInt32(items[9]);
                data.RepoAgentName = items[10].ToString();
                data.RepoAgentsLookup = items[11].ToString();
                data.LocationOfUnit = items[12].ToString();
                data.DateOfRepo =  Convert.ToDateTime(items[13]);
                data.DateOfClear = Convert.ToDateTime(items[14]);
                fileData.Add(data);
            }
            _logger.Information($"Read {fileData.Count} row(s) from: {config.SourceFile}");
            return fileData;
        }
    }
}
