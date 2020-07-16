using Microsoft.Extensions.Configuration;
using MPSAssestManagementFileConverter.BusinessLogic;
using System;
using System.IO;
using System.Linq;

namespace AssetManagementUserFileConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
         .SetBasePath(Directory.GetCurrentDirectory())
         .AddJsonFile("appsettings.json");

            var configuration = builder.Build();

            const string inputArgument = "--input";
            const string outputArgument = "--output";
            const string companyArgument = "--company";
            const string headerArgument = "--header";

            string fileName = configuration["fileName"];
            string outputFileName = configuration["outputFileName"];
            string companyName = configuration["companyName"];
            string outputHeader = configuration["outputHeader"];
            
            // Test if input arguments were supplied.
            if (args.Length != 0)
            {
                //Override appsettings if args are supplied
                if (args.Contains(inputArgument))
                {
                    var inputValueIndex = Array.IndexOf(args, inputArgument) + 1;
                    fileName = args[inputValueIndex];
                }

                if (args.Contains(outputArgument))
                {
                    var outputValueIndex = Array.IndexOf(args, outputArgument) + 1;
                    outputFileName = args[outputValueIndex];
                }

                if (args.Contains(companyArgument))
                {
                    var companyValueIndex = Array.IndexOf(args, companyArgument) + 1;
                    companyName = args[companyValueIndex];
                }

                if (args.Contains(headerArgument))
                {
                    var headerValueIndex = Array.IndexOf(args, headerArgument) + 1;
                    outputHeader = args[headerValueIndex];
                }
            }


            try
            {
                var converter = new FileConverter();
                converter.ConvertFile(fileName, outputFileName, companyName, outputHeader);


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
