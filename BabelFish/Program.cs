using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Web;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using GoogleTranslateNET;
using Mono.Options;

namespace BabelFish
{
    class Program
    {
        static Program()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (s, e) =>
            {
                try
                {
                    var assembly = Assembly.GetExecutingAssembly();
                    var assemblyName = new AssemblyName(e.Name);
                    var fileName = assemblyName.Name + ".dll";
                    var resources = assembly.GetManifestResourceNames().Where(mrn => mrn.EndsWith(fileName));
                    if (resources.Any())
                    {
                        var resourceName = resources.First();
                        using (var stream = assembly.GetManifestResourceStream(resourceName))
                        {
                            if (stream == null)
                                return null;

                            var buffer = new byte[stream.Length];
                            try
                            {
                                stream.Read(buffer, 0, buffer.Length);
                                return Assembly.Load(buffer);
                            }
                            catch (IOException)
                            {
                                return null;
                            }
                            catch (BadImageFormatException)
                            {
                                return null;
                            }
                        }
                    }
                }
                catch
                {

                }
                return null;
            };
        }

        static void Main(string[] args)
        {
            if (Debugger.IsAttached)
            {
                args = new string[]
                {
                    "Hello World!",
                    //"-f",
                    //"en",
                    "-t",
                    "ru",
                    //"-s",
                    //"D:\\Temp\\en.TXT",
                    //@"D:\Source\Svn\GO.Desktop\GO.Library\GO.Res\Resources.resx",
                    //@"D:\Source\Svn\GO.Desktop\GO.Library\GO.Res\Resources.xlsx",
                    //"-d",
                    //"D:\\Temp\\fi.txt",
                    //@"D:\Source\Svn\GO.Desktop\GO.Library\GO.Res\Resources.fi.resx",
                    //@"D:\Source\Svn\GO.Desktop\GO.Library\GO.Res\Resources.fi.xlsx",
                    //"--take",
                    //"10",
                    //"вторник"
                    "-apikey",
                    File.ReadAllText("D:\\Frontmatec\\Bin\\babelfish.apikey"),
                };
            }

            try
            {
                if (args == null || !args.Any())
                    throw new Exception("");

                var fromCultureName = string.Empty;
                var toCultureName = string.Empty;
                var sourceFileName = string.Empty;
                var destinationFileName = string.Empty;
                var take = 0;
                var apiKey = string.Empty;
                var showHelp = false;
                var options = new OptionSet {
                    { "f|from=", "from language.", a => fromCultureName = a },
                    { "t|to=", "to language.", a => toCultureName = a },
                    { "s|source=", "specifies the file to be translated.", a => sourceFileName = a },
                    { "d|destination=", "specifies the file for the translated file.", a => destinationFileName = a },
                    { "take=", "specified the number or lines to translate.", a => take = int.Parse(a) },
                    { "apikey=", "specifies the google api key.", a => apiKey = a },
                    { "h|help", "shows this message.", a => showHelp = a != null },

                };
                var unprocessedArgs = options.Parse(args)?.ToArray();
                var sourceText = string.Join(" ", unprocessedArgs);

                if (showHelp)
                {
                    ShowHelp(options);
                    return;
                }
                
                if (!string.IsNullOrWhiteSpace(sourceFileName) && string.IsNullOrWhiteSpace(destinationFileName))
                {
                    var fileExtension = Path.GetExtension(sourceFileName);
                    if (".xlsx".Equals(fileExtension, StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (string.IsNullOrWhiteSpace(fromCultureName))
                            fromCultureName = "en";

                        if (string.IsNullOrWhiteSpace(toCultureName))
                        {
                            using (var workbook = new XLWorkbook(sourceFileName))
                                toCultureName = workbook.Worksheets.FirstOrDefault()?.Name ?? string.Empty;
                        }

                        destinationFileName = sourceFileName;
                    }
                }
                
                var sourceLanguage = !string.IsNullOrWhiteSpace(fromCultureName) ? ParseLanguage(fromCultureName) : Language.Unknown;
                if (sourceLanguage == Language.Unknown && !string.IsNullOrWhiteSpace(sourceFileName))
                    throw new Exception("from not specified.");

                if (string.IsNullOrWhiteSpace(toCultureName) && string.IsNullOrWhiteSpace(sourceFileName))
                    toCultureName = CultureInfo.CurrentCulture.TwoLetterISOLanguageName;
                
                var destinationLanguage = !string.IsNullOrWhiteSpace(toCultureName) ? ParseLanguage(toCultureName) : Language.Unknown;
                if (destinationLanguage == Language.Unknown)
                    throw new Exception("to not specified.");
                
                var resources = !string.IsNullOrWhiteSpace(sourceFileName) ? ReadFile(sourceFileName, fromCultureName) : new Resource[] { new Resource { SourceText = sourceText } };
                if (resources == null)
                    resources = new Resource[0];
                
                if (take > 0)
                    resources = resources.Take(take).ToArray();
                
                if (!string.IsNullOrWhiteSpace(destinationFileName) && File.Exists(destinationFileName) && !destinationFileName.Equals(sourceFileName))
                {
                    Console.Write($"{Path.GetFileName(destinationFileName)} already exists. Do you want to replace it? [Y/n]");
                    if (Console.ReadLine().Equals("n", StringComparison.InvariantCultureIgnoreCase))
                        return;
                    File.Delete(destinationFileName);
                }

                if (string.IsNullOrWhiteSpace(apiKey) && File.Exists(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "babelfish.apikey")))
                    apiKey = File.ReadAllText(Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "babelfish.apikey")));
                if (string.IsNullOrWhiteSpace(apiKey))
                    throw new Exception("apikey not specified.");

                var useProgressBar = !string.IsNullOrWhiteSpace(destinationFileName) && resources.Length > 3;
                var google = new GoogleTranslate(apiKey);
                for(var i = 0; i < resources.Length; i++)
                {
                    var resource = resources[i];

                    if (!string.IsNullOrWhiteSpace(resource.SourceText))
                    {
                        if (sourceLanguage == Language.Unknown)
                            sourceLanguage = ParseLanguage(google.DetectLanguage(resource.SourceText)[0].Language);
                        resource.DestinationText = HttpUtility.HtmlDecode(google.Translate(sourceLanguage, destinationLanguage, resource.SourceText)[0].TranslatedText);
                    }

                    if (useProgressBar)
                    {
                        Console.CursorLeft = 0;
                        Console.Write($"Translating: {(i + 1).ToString().PadLeft(resources.Length.ToString().Length)}/{resources.Length} [{new string('#', ((int)100.0 * (i + 1) / resources.Length) / 2).PadRight(50, '.')}]");
                    }
                }

                if (useProgressBar)
                    Console.WriteLine();

                WriteFile(destinationFileName, fromCultureName, toCultureName, resources);
            }
            catch (Exception ex)
            {
                if (!string.IsNullOrWhiteSpace(ex.Message))
                    Console.WriteLine(ex.Message);
                Console.WriteLine("Try 'babelfish --help' for more information.");
            }
            finally
            {
                if (Debugger.IsAttached)
                {
                    Console.ReadLine();
                }
            }
        }

        static void ShowHelp(OptionSet options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            Console.WriteLine("Usage: babelfish [options...] <text>");

            var stringBuilder = new StringBuilder();
            using (TextWriter writer = new StringWriter(stringBuilder))
                options.WriteOptionDescriptions(writer);
            Console.WriteLine(stringBuilder.ToString().ToLower());
        }

        static Language ParseLanguage(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                throw new ArgumentNullException(nameof(text));

            var culture = CultureInfo.GetCultureInfo(text);
            var language = (Language)Enum.Parse(typeof(Language), culture.EnglishName);
            return language;
        }

        static Resource[] ReadFile(string fileName, string fromCultureName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentNullException(nameof(fileName));
            
            // Get file type
            var fileExtension = Path.GetExtension(fileName);

            if (".txt".Equals(fileExtension, StringComparison.InvariantCultureIgnoreCase))
                return File.ReadAllLines(fileName).Select(l => new Resource { SourceFileName = fileName, ResourceName = string.Empty, SourceText = l }).ToArray();
            
            if (".resx".Equals(fileExtension, StringComparison.InvariantCultureIgnoreCase))
            {
                var resources = new List<Resource>();
                var reader = new ResXResourceReader(fileName) { UseResXDataNodes = true };
                foreach (DictionaryEntry entry in reader)
                {
                    var node = (ResXDataNode)entry.Value;
                    if (node.FileRef != null || node.Name.StartsWith("$this.") || node.Name.StartsWith(">>"))
                        continue;

                    var value = node.GetValue((ITypeResolutionService)null);
                    if (!(value is string))
                        continue;

                    resources.Add(new Resource { SourceFileName = fileName, ResourceName = node.Name, SourceText = (string)value });
                }
                return resources.ToArray();
            }

            if (".xlsx".Equals(fileExtension, StringComparison.InvariantCultureIgnoreCase))
            {
                var resources = new Dictionary<string, Resource>();
                using (var workbook = new XLWorkbook(fileName))
                {
                    var worksheets = workbook.Worksheets;

                    IXLWorksheet worksheet = null;
                    string columnLetter = null;
                    if ("en".Equals(fromCultureName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        worksheet = worksheets.FirstOrDefault();
                        columnLetter = "B";
                    }
                    else
                    {
                        worksheet = worksheets.FirstOrDefault(w => w.Name.Equals(fromCultureName, StringComparison.InvariantCultureIgnoreCase));
                        columnLetter = "C";
                    }

                    if (worksheet == null)
                        throw new Exception("Could not find worksheet.");

                    var rows = worksheet.Rows();
                    foreach (var row in rows)
                    {
                        if (row.RowNumber() == 1)
                            continue;

                        var resourceName = Convert.ToString(row.Cell("A").Value);
                        var sourceText = Convert.ToString(row.Cell(columnLetter).Value);
                        var resource = new Resource { SourceFileName = fileName, ResourceName = resourceName, SourceText = sourceText };

                        resources[resourceName] = resource;
                    }
                }
                return resources.Values.ToArray();
            }

            throw new NotSupportedException($"File type '{fileExtension}' is not supported.");
        }

        static void WriteFile(string fileName, string fromCultureName, string toCultureName, Resource[] resources)
        {
            if (resources == null)
                throw new ArgumentNullException(nameof(resources));

            // Console WriteLine
            if (string.IsNullOrWhiteSpace(fileName))
            {
                Console.OutputEncoding = Encoding.UTF8;
                foreach (var resource in resources)
                {
                    if (!resource.Equals(resources.First()))
                        Console.WriteLine();
                    Console.Write(resource.DestinationText);
                }
                return;
            }

            // Get file type
            var fileExtension = Path.GetExtension(fileName);
            
            if (".txt".Equals(fileExtension, StringComparison.InvariantCultureIgnoreCase))
            {
                File.WriteAllText(fileName, string.Join("\r\n", resources.Select(r => r.DestinationText)));
                return;
            }

            if (".resx".Equals(fileExtension, StringComparison.InvariantCultureIgnoreCase))
            {
                using (var writer = new ResXResourceWriter(fileName))
                {
                    foreach (var resource in resources)
                    {
                        if (string.IsNullOrWhiteSpace(resource.DestinationText))
                            continue;
                        writer.AddResource(new ResXDataNode(resource.ResourceName, resource.DestinationText));
                    }
                }
                return;
            }

            if (".xlsx".Equals(fileExtension, StringComparison.InvariantCultureIgnoreCase))
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(toCultureName);
                    worksheet.Column("A").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    worksheet.Column("B").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    worksheet.Column("C").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    worksheet.Column("A").Style.Alignment.WrapText = true;
                    worksheet.Column("B").Style.Alignment.WrapText = true;
                    worksheet.Column("C").Style.Alignment.WrapText = true;
                    worksheet.Column("A").Width = 50;
                    worksheet.Column("B").Width = 85;
                    worksheet.Column("C").Width = 85;

                    worksheet.Cell(1, "A").Style.Font.SetBold();
                    worksheet.Cell(1, "B").Style.Font.SetBold();
                    worksheet.Cell(1, "C").Style.Font.SetBold();
                    worksheet.Cell(1, "A").Value = "Resource Name";
                    worksheet.Cell(1, "B").Value = fromCultureName;
                    worksheet.Cell(1, "C").Value = toCultureName;

                    for (int i = 0; i < resources.Length; i++)
                    {
                        var resource = resources[i];
                        worksheet.Cell(i + 2, "A").Value = resource.ResourceName;
                        worksheet.Cell(i + 2, "B").Value = resource.SourceText;
                        worksheet.Cell(i + 2, "C").Value = resource.DestinationText;
                    }

                    worksheet.RangeUsed().SetAutoFilter();
                    workbook.SaveAs(fileName);
                }
                return;
            }
            
            throw new NotSupportedException($"File type '{fileExtension}' is not supported.");
        }
    }
}
