using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Pulse;

namespace CommandLineApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define required flags and allowed file types
            var requiredFlags = new[] { "-d", "-u", "-p", "-x", "-f", "-s" };
            var allowedFileTypes = new[] { "pxf", "pcf", "png" };

            // Validate presence of command and server name
            if (args.Length < 2)
            {
                ShowHelp();
                return;
            }

            var command = args[0];
            var serverName = args[1];

            if (!string.Equals(command, "GetDesign", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Unknown command '{command}'\n");
                ShowHelp();
                return;
            }

            if (string.IsNullOrWhiteSpace(serverName))
            {
                Console.WriteLine("ServerName is required.\n");
                ShowHelp();
                return;
            }

            // Process flag arguments
            var flagArgs = args.Skip(2).ToArray();
            if (flagArgs.Length != requiredFlags.Length * 2)
            {
                Console.WriteLine("Incorrect number of flags provided.\n");
                ShowHelp();
                return;
            }

            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < flagArgs.Length; i += 2)
            {
                var flag = flagArgs[i];
                var value = flagArgs[i + 1];

                if (!requiredFlags.Contains(flag, StringComparer.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"Unknown flag '{flag}'\n");
                    ShowHelp();
                    return;
                }
                values[flag] = value;
            }

            // Verify all flags are present
            foreach (var flag in requiredFlags)
            {
                if (!values.TryGetValue(flag, out var val) || string.IsNullOrWhiteSpace(val))
                {
                    Console.WriteLine($"Missing required flag {flag}\n");
                    ShowHelp();
                    return;
                }
            }

            // Extract and validate file type
            var databaseName = values["-d"];
            var userName = values["-u"];
            var password = values["-p"];
            var exportPath = values["-x"];
            var fileType = values["-f"].ToLowerInvariant();
            var design = values["-s"].ToUpperInvariant();

            if (!allowedFileTypes.Contains(fileType))
            {
                Console.WriteLine($"Invalid file type '{fileType}'. Allowed types: {string.Join(", ", allowedFileTypes)}\n");
                ShowHelp();
                return;
            }

            // Display summary
            Console.WriteLine("Running command with parameters:");
            Console.WriteLine($"  Command:     {command}");
            Console.WriteLine($"  Server:      {serverName}");
            Console.WriteLine($"  Database:    {databaseName}");
            Console.WriteLine($"  User:        {userName}");
            Console.WriteLine($"  Password:    {password}");
            Console.WriteLine($"  Design:      {design}");
            Console.WriteLine($"  Export Path: {exportPath}");
            Console.WriteLine($"  File Type:   {fileType}");

            // TODO: Implement GetDesign logic against the Librarian server
            IApplication PulseID = new Pulse.Application();
            try
            {
                ILibrarianConnection librarianConnection = PulseID.NewLibrarianConnection();
                try
                {
                    librarianConnection.Connect(serverName,9000,userName,password,LibrarianClientTypes.lctPdlClient);
                    librarianConnection.OpenDatabase(databaseName);
                 

                    string str = "[design_id] LIKE '*"+design+"*'";
                                           
                    ILibrarianDesigns designs = librarianConnection.Search(str,1000);
                    try
                    {
                        if (designs != null & designs.Count>0)
                        {
                            Console.WriteLine("Found " + designs.Count + " design(s) matching the criteria.");
                            ILibrarianDesign des = designs[0];
                            try
                            {
                                for (int i = 0; i < designs.Count; i++)
                                {
                                    Console.WriteLine("Exporting design: " + designs[i].Name);
                                    des = designs[i];
                                    IEmbDesign embDesign = librarianConnection.OpenDesign(des.Name);
                                    try
                                    {
                                        if ( fileType.ToUpper() == "PNG")
                                        {
                                            IBitmapImage image = PulseID.NewImage(300, 300);
                                            try
                                            {
                                                embDesign.Render(image, 0, 0, 290, 290);
                                                image.Save(exportPath + "\\" + des.Name + ".png", ImageTypes.itPNG);

                                            }
                                            finally
                                            {
                                                Marshal.ReleaseComObject(image); image = null;
                                            }
                                        }
                                        if (fileType.ToUpper() == "PCF")
                                        {
                                            embDesign.Save(exportPath + "\\" + des.Name + ".pcf", FileTypes.ftPCF);
                                        }
                                        if (fileType.ToUpper() == "PXF")
                                        {
                                            embDesign.Save(exportPath + "\\" + des.Name + ".pxf", FileTypes.ftPXF);
                                        }
                                    }
                                    finally
                                    {
                                        Marshal.ReleaseComObject(embDesign); embDesign = null;
                                    }
                                }

                            }




                            finally
                            {
                                Marshal.ReleaseComObject(des); des = null;
                            }
                            
                        }
                        
                       
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(designs); designs = null;

                    }
                }

                catch (Exception ex)
                {
                    Console.WriteLine("Unable to connect to server: " + serverName);
                    Console.WriteLine(ex.Message);
                    Console.WriteLine(ex.StackTrace);
                }
                finally
                {
                    
                    Marshal.ReleaseComObject( librarianConnection ); librarianConnection = null;
                }
              
            }
            finally
            {
                Marshal.ReleaseComObject(PulseID); PulseID = null;

            }
            


        }

        static void ShowHelp()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  LibrarianUtil.exe <Command> <ServerName> -d <DatabaseName> -u <UserName> -p <Password> -e <ExportPath> -f <FileType>");
            Console.WriteLine();
            Console.WriteLine("Commands:");
            Console.WriteLine("  GetDesign       Retrieves design from the Librarian server");
            Console.WriteLine();
            Console.WriteLine("Arguments:");
            Console.WriteLine("  ServerName      Name of the Librarian server");
            Console.WriteLine();
            Console.WriteLine("Flags (all required):");
            Console.WriteLine("  -d    Database name");
            Console.WriteLine("  -u    User name");
            Console.WriteLine("  -s    Embroidery design");
            Console.WriteLine("  -p    Password");
            Console.WriteLine("  -x    Export path");
            Console.WriteLine("  -f    File type (pxf, pcf, png)");
        }
    }
}
