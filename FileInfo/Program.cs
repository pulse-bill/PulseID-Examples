using Pulse;
using System.Data;
using System.Text;
using static System.Net.Mime.MediaTypeNames;
// See https://aka.ms/new-console-template for more information

IApplication PulseID = new Pulse.Application();

string[] _allowedExtensions = { "pcf", "pxf" };

string _inputFolder = @"C:\\LT\Files\";

StringBuilder sbOutput = new StringBuilder();
string csvRow = "";
string CSVFile = @"C:\LT\Output\Export.csv";

if (File.Exists(CSVFile))
    File.Delete(CSVFile);

sbOutput.AppendLine("Design, Width, Height, Stitches,NumColors,Palette,Colors");

File.AppendAllText(CSVFile, sbOutput.ToString());

sbOutput.Clear();


string[] _filesToProcess = Directory.EnumerateFiles(_inputFolder, "*.*", SearchOption.AllDirectories)
 .Where(processFiles => _allowedExtensions.Any(processFiles.ToLower().EndsWith))
 .ToArray();

foreach (string _file in _filesToProcess)
{
    csvRow = "";
    sbOutput.Clear();



    IEmbDesign myDesign = PulseID.OpenDesign(_file, FileTypes.ftAuto, OpenTypes.otDefault, "Tajima");

    INeedleSequence needles = myDesign.NeedleSequence;

    for (int i = 0; i < myDesign.AllThreadPalettes.Count; i++)
    {
        csvRow = "";
        IThreadPalette palette = myDesign.AllThreadPalettes[i];
        string strPalette = palette.Name;
        Console.WriteLine("Design: " + Path.GetFileName(_file) + " Palette: " + strPalette);






        IEmbDesignStatistics stats = myDesign.GetStatistics();
        csvRow += Path.GetFileName(_file) + ",";

        double width = myDesign.Width / 254.00;
        csvRow += width.ToString("#.##") + ",";


        double height = myDesign.Height / 254.00;
        csvRow += width.ToString("#.##") + ",";

        uint stitches = myDesign.NumStitches;
        csvRow += stitches.ToString() + ",";

        int numColors = stats.NumColourChanges + 1;
        csvRow += numColors.ToString() + ",";

        string paletteName = palette.Name;
        if (paletteName.Contains(","))
        {
            paletteName = paletteName.Replace(",", "-");

        }


        csvRow += paletteName + ",";

        foreach (UInt16 needle in needles)
        {



            csvRow += palette[(int)needle].Name + ",";

            csvRow += palette[(int)needle].Code + ",";





        }

        sbOutput.AppendLine(csvRow);






    }
    csvRow = "";
    File.AppendAllText(CSVFile, sbOutput.ToString());
    sbOutput.Clear();



}
