using System;
using System.IO;
using CommandLine;

namespace RmkXlsToXml
{
    public class ConverterConfiguration
    {
        [Option('s', "sourceFile", Required = true, HelpText = "XLS file to convert.")]
        public string SourceFile { get; set; }

        [Option('o', "outputPath", Required = false, HelpText = "Output folder. (default is current folder)", Default = ".")]
        public string OutputPath { get; set; } = ".";
        [Option('r', "RSAClientId", Required = true, HelpText = "RSA Client Id")]
        public string RsaClientId { get; set; }

        [Option('h', "headerRowCount", Required = false, HelpText = "Number of header rows in the xls file.", Default = 3)]
        public int NumberOfHeaderRows { get; set; } = 3;


      
        public bool IsEmpty()
        {
            return string.IsNullOrEmpty(this.SourceFile) || 
                   string.IsNullOrEmpty(this.OutputPath) ||
                   string.IsNullOrEmpty(this.RsaClientId);
        }
        public bool Validate()
        {
            if (string.IsNullOrEmpty(this.SourceFile))
            {
                Console.WriteLine("Missing required source file argument.");
                return false;
            }
            if (string.IsNullOrEmpty(this.OutputPath))
            {
                Console.WriteLine("Missing required output path argument.");
                return false;
            }

            if (string.IsNullOrEmpty(this.RsaClientId))
            {
                Console.WriteLine("Missing required RSAClientId path argument.");
                return false;
            }
            if (!File.Exists(this.SourceFile))
            {
                Console.WriteLine($"Source File '{this.SourceFile}' not found.'");
                return false;
            }
            string absolute = Path.GetFullPath(OutputPath);
            if (string.IsNullOrEmpty(absolute))
            {
                Console.WriteLine($"Invalid output path '{absolute}'.'");
                return false;
            }

            if (!Directory.Exists(absolute))
            {
                try
                {
                    var _ = Directory.CreateDirectory(absolute);
                }
                catch (Exception )
                {
                    Console.WriteLine($"Invalid output path '{absolute}'.'");
                    return false;
                }
            }

            return true;
        }

        public static void ShowUsage()
        {
            string exeName = AppDomain.CurrentDomain.FriendlyName;
            Console.WriteLine($"{exeName} -s <SourceFile> -o <OutputPath> -r <RSAClientId> [-h NumberOfHeaderRows]");
        }
    }
}