using System;
using CommandLine;
using Serilog;
using Serilog.Events;

namespace RmkXlsToXml
{
    class Program
    {
        static void Main(string[] args)
        {
            ConfigureLogger();
            // parse the command line args into a strongly typed ConverterConfiguration instance
            Parser.Default.ParseArguments<ConverterConfiguration>(args)
                .WithParsed(config =>
                {
                    try
                    {
                        if (!config.Validate())
                        {
                            HandleError("Invalid configuration.");
                            return;
                        }

                        Log.Logger.Information($"Converting data from {config.SourceFile} to xml.");
                        var converter = new RemarketingDataConverter(Log.Logger);
                        var goodConvert = converter.ConvertRemarketingFile(config);
                        if (!goodConvert)
                        {
                            HandleError("Conversion failed.");
                        }
                        else
                        {
                            Log.Logger.Information("Conversion successful.");
                        }

                    }
                    catch (Exception e)
                    {
                        HandleError("Runtime Exception",e);
                    }
                })
                .WithNotParsed((errors) =>
                {

                    HandleError("Invalid Configuration");
                    ConverterConfiguration.ShowUsage();
                });
        }


     
        static void HandleError(string message, Exception e = null)
        {
            Log.Logger.Error(e,message);
            
            Environment.ExitCode = -1;
        }

        private static void ConfigureLogger()
        {
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .MinimumLevel.Override("Microsoft", LogEventLevel.Warning)
                .Enrich.FromLogContext()
                .WriteTo.Console()
                .CreateLogger();
        }
    }
}
