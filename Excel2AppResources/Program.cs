using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Excel2AppResources;

public class Program
{
    public static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        string currentFile = args[^1];

        Program program = new Program();
        ApplicationArguments applicationArguments = new ApplicationArguments();
        
        applicationArguments.Add("sheet", "Sheet index the translation are on. First sheet is 0", "1");
        applicationArguments.Add("rows", "Number of rows to import default 1", "1");
        applicationArguments.Add("resource-set", "Resource-Set name to add to the SQLs default its set to resource-name", "resource-name");
        applicationArguments.Add("resource-offset", "Resource offset is the address Column:Row the ResourceName is placed. default: H1 ", "H1");
        applicationArguments.Add("import-offset", "Import offset is the address Column:Row the translations are placed. default: I1 ", "I1");
        applicationArguments.Add("output", "Name of the output file generated. default: UpdateAppResources.sql ", "UpdateAppResources.sql");
        applicationArguments.Add("help", "shows this help context", "");
        
        
        for (int index = 0; index < args.Length; index++)
        {
            string argument = args[index].ToLower().Trim();
            if (argument.StartsWith("--"))
            {
                if (argument.Equals("--help"))
                {
                    applicationArguments.PrintHelp();
                    return;
                }
                int nextIndex = index + 1;
                if (nextIndex < args.Length)
                {
                    string value = args[nextIndex];
                    string key = argument[2..];
                    applicationArguments.Add(key, value.Trim());
                }
                index++;
            }

        }

        program.GenerateSqlFromFile(applicationArguments, currentFile);
    }

    private void GenerateSqlFromFile(ApplicationArguments applicationArguments, string currentFile)
    {
        Console.WriteLine($"Reading {currentFile}");

        int sheetIndex = applicationArguments.Get("sheet", 1);
        int rowsToImport = applicationArguments.Get("rows", 1);;
        string resourceSet = applicationArguments.Get("resource-set", "resource-name");
        
        string resourceName =  applicationArguments.Get("resource-offset", "H1");
        string startsFrom =    applicationArguments.Get("import-offset","I1");
        string outputFile = applicationArguments.Get("output","UpdateAppResources.sql");

        int updatesWritten = 0;
        if (File.Exists(outputFile)) File.Delete(outputFile);

        using (var fs = new FileStream(currentFile, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook = new XSSFWorkbook(fs);
            ISheet sheet = workbook.GetSheetAt(sheetIndex);

            SheetLocation resourceLocation = GetLocationFromString(resourceName);
            SheetLocation importLocation = GetLocationFromString(startsFrom);
            StreamWriter streamWriter = new StreamWriter(new FileStream(outputFile, FileMode.OpenOrCreate), Encoding.UTF8);

            streamWriter.AutoFlush = true;
            WriteStartOfScript(streamWriter);

            while (rowsToImport > 0)
            {
                Console.WriteLine("Exporting resource " + sheet.GetRow(resourceLocation.Row).Cells[resourceLocation.Column].StringCellValue);
                resourceName = sheet.GetRow(resourceLocation.Row).Cells[resourceLocation.Column].StringCellValue;

                int languageHeaderIndex = importLocation.Column;
                string language = sheet.GetRow(0).Cells[languageHeaderIndex].StringCellValue;

                while (language.Trim().Length > 0)
                {
                    languageHeaderIndex++;
                    if (sheet.GetRow(0).Cells.Count <= languageHeaderIndex) break;
                    language = sheet.GetRow(0).Cells[languageHeaderIndex].StringCellValue;
                    string translation = sheet.GetRow(importLocation.Row).Cells[languageHeaderIndex].StringCellValue.Replace("'","''");
                    WriteTranslation(streamWriter, resourceSet, resourceName, language, translation);
                    updatesWritten++;
                }

                resourceLocation.Row++;
                importLocation.Row++;
                rowsToImport--;
            }
    
            WriteEndOfScript(streamWriter);
            Console.WriteLine($"Total of {updatesWritten} updates are written");
        }
    }

    private void WriteTranslation(StreamWriter streamWriter, string set, string name, string language, string value)
    {
        string dateTime =  DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") ;
        string lines = $"\ndbo.UpsertAppResource  '{set}', '{name}', '{language}', N'{value}', '{dateTime}', 1\nGO\n";
        streamWriter.Write(lines);
    }

    private void WriteStartOfScript(StreamWriter streamWriter)
    {
        const string beginning = """
        -- Begin transaction
        BEGIN TRAN AppResourceUpdate

        --- Delete upsert procedure
        DROP PROCEDURE dbo.UpsertAppResource
        GO

        -- Create or update upsert procedure
        -- Unique index: [ResourceSet] + [ResourceName] + [ResourceCulture]

        CREATE PROCEDURE dbo.UpsertAppResource
               @ResourceSet NVARCHAR(255),
               @ResourceName NVARCHAR(128), 
               @ResourceCulture NVARCHAR(10), 
               @ResourceValue NVARCHAR(MAX), 
               @CreateDate DATETIME, 
               @isApproved BIT
        AS BEGIN
                UPDATE [dbo].[AppResources]
                SET ResourceSet = @ResourceSet, ResourceName = @ResourceName, ResourceValue = @ResourceValue, ResourceCulture = @ResourceCulture, CreateDate = @CreateDate, isApproved = @isApproved
                WHERE ResourceSet = @ResourceSet AND ResourceName = @ResourceName AND ResourceCulture = @ResourceCulture
                IF @@ROWCOUNT = 0
                    insert into [dbo].[AppResources] 
                    values (@ResourceSet, @ResourceName, @ResourceValue, @ResourceCulture, @CreateDate, @isApproved)
        END
        GO
        
        
        """;
        streamWriter.Write(beginning);
    }

    private void WriteEndOfScript(StreamWriter streamWriter)
    {
        const string ending = """
        
        -- Commit or rollback transaction 
        COMMIT TRANSACTION  AppResourceUpdate;
     
        -- ROLLBACK TRANSACTION  AppResourceUpdate;
        """;
        streamWriter.Write(ending);
    }


    private SheetLocation GetLocationFromString(string location)
    {
        string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string columnName = location[..1];
        string rowName = location[1..];

        SheetLocation sheetLocation = new SheetLocation(alphabet.IndexOf(columnName, StringComparison.Ordinal), int.Parse(rowName));
        return sheetLocation;
    }


    private class SheetLocation
    {
        public SheetLocation(int column, int row)
        {
            Column = column;
            Row = row;
        }

        public int Column { get; init; }
        public int Row { get; set; }

        public override string ToString()
        {
            return $"{nameof(Column)}: {Column}, {nameof(Row)}: {Row}";
        }
    }
}