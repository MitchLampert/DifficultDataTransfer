using System.Data.OleDb;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

/*
This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or distribute this software, either in source code form or as a compiled binary, for any purpose, commercial or non-commercial, and by any means.
In jurisdictions that recognize copyright laws, the author or authors of this software dedicate any and all copyright interest in the software to the public domain. We make this dedication for the benefit of the public at large and to the detriment of our heirs and successors. We intend this dedication to be an overt act of relinquishment in perpetuity of all present and future rights to this software under copyright law.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
//Original Author: Mitchell Scott Lampert


namespace DifficultDataTransfer
{
    internal enum BasicDataType
    {
        DTWholeNumber,
        DTDecimalNumber,
        DTString,
        DTDateTime,
        DTDateOnly,
        DTTimeOnly,
        DTBoolean,
        DTOther
    }

    internal enum TrimRule
    {
        TRTrimBothSides,
        TRTrimFrontOnly,
        TRTrimEndOnly,
        TRNoTrimming
    }

    internal enum TruncationRule
    {
        TRNoTruncating,
        TRTruncateString,
        TRTruncateWithEllipsis
    }

    internal struct DataTransCol
    {
        public string SourceField;
        public string DestinationField;
        
        public string OfficialDataType;
        public BasicDataType TransDataTypeRule;
        public bool IsNullable;

        public int StrFieldSize;
        public TrimRule StrTrimRule;
        public TruncationRule StrTruncRule;

        public bool DefaultIsNull;
        public string DefaultValue;
    }

    internal class Program
    {
        static Version? VersionNum = Assembly.GetEntryAssembly()?.GetName().Version;

        static string SourceConnStr = string.Empty;
        static string SourceTableName = string.Empty;
        static string WhereClause = string.Empty;

        static string DestConnStr = string.Empty;
        static string DestTableName = string.Empty;

        static bool IDInsert = false;
        static bool StopOn1stError = false;

        static int BatchSize = 100;

        static string LogFilePath = string.Empty;

        static List<DataTransCol> DataTransCols = new List<DataTransCol>();

        static string InsCmdStart = string.Empty;
        static string InsCmdEnd = string.Empty;

        static int RowsTransferred = 0;
        static int RowsWithErrors = 0;

        static bool AtLeastOneSuccess = false;
        static int CurrentRowNum = 0;

        static void Main(string[] args)
        {
            Console.WriteLine("Mitchell's Difficult Data Transfer Program");
            Console.WriteLine("------------------------------------------");
            Console.WriteLine("Version " + VersionNum);
            Console.WriteLine();

            //This version of the program will NOT truncate the destination table, first. You will have to do that in a separate process before running this, if you must.
            //Or alter the source, youself, to do so, if you insist.

            //-----------------------------------------------------------------------
            //Get configuration file location.
            if (args.Length == 0)
            {
                Console.WriteLine("No configuration file set. This program needs a path to a .DifficultDataConfig file in order to run. Use the DifficultDataAutoGenConfig program to create one. "
                                   + "Read instructions manual for more details.");
                return;
            }
            else
            {
                //Concant all args into a single string for the path of the file, and remove any quote marks from its start or end.
                string ConfigPath = string.Join(" ", args).Trim().TrimStart('"').TrimEnd('"');
                if (!ParseConfigurationFile(ConfigPath)) return;
                if (!BasicValidationTasks()) return;

                Console.WriteLine("Configuration File: " + ConfigPath);
                File.AppendAllText(LogFilePath, "Configuration File: " + ConfigPath + Environment.NewLine);
                Console.WriteLine();
            }

            using (OleDbConnection SourceCn = new OleDbConnection(SourceConnStr))
            using (OleDbConnection DestCn = new OleDbConnection(DestConnStr))
            {
                Console.WriteLine("Opening Source Connection...");
                try
                {
                    SourceCn.Open();
                }
                catch (Exception ex)
                {
                    string ErrStr = "Error opening Source connection: " + ex.Message;
                    Console.WriteLine(ErrStr);
                    File.AppendAllText(LogFilePath, ErrStr + Environment.NewLine);
                    return;
                }

                Console.WriteLine("Opening Destination Connection...");
                try
                {
                    DestCn.Open();
                }
                catch (Exception ex)
                {
                    string ErrStr = "Error opening Destination connection: " + ex.Message;
                    Console.WriteLine(ErrStr);
                    File.AppendAllText(LogFilePath, ErrStr + Environment.NewLine);
                    return;
                }

                InsCmdStart = BuildStartOfINSERTStatement();
                InsCmdEnd = BuildEndOfINSERTStatement();

                //Note 1: At this point in the program, you can add your own code to achieve certain things, before the data transfer starts.
                //For example, you can run a command on the destination connection (DestCn) to truncate the source table (SourceTableName), if you need to.

                //Note 2: We are not calling the Garbage Collector, here, in this version of the program. At least not anymore. You can add that, too, if you feel you should.
                
                //---------------------------------------------------------------------------------------------
                //Run command to get data from Source 
                Console.WriteLine("Running SQL SELECT Command...");
                OleDbCommand SourceCmd = new OleDbCommand(BuildSELECTStatementOfSource(), SourceCn);

                try
                {
                    using (OleDbDataReader TheReader = SourceCmd.ExecuteReader())
                    {
                        if (!TheReader.HasRows)
                        {
                            string ErrStr = "Error: No rows found in source.";
                            Console.WriteLine(ErrStr);
                            File.AppendAllText(LogFilePath, ErrStr + Environment.NewLine);
                            return;
                        }

                        //--------------------------------------------------------------------------------
                        //Data transfer time!
                        Console.WriteLine();
                        Console.WriteLine("Starting Data Transfer...");

                        List<string> InsertValues = new List<string>();
                    
                        bool HasARow = true;

                        while (HasARow) //For each row in the Source:
                        {
                            HasARow = TheReader.Read();
                            if (HasARow)
                            {
                                CurrentRowNum++;
                                InsertValues.Add(BuildINSERTValue(TheReader));
                            } 

                            if (InsertValues.Count == BatchSize || (!HasARow & InsertValues.Count > 0))
                            {
                                if (!ExecuteINSERTStatement(InsertValues, DestCn)) return;

                                InsertValues.Clear();

                            } //Executing stuff. BatchSize or !HasARow 

                        }  //While Reader has a Row

                    } //Using DataReader
                }
                catch (Exception ex)
                {
                    string ErrStr = "Error running Reader from Source: " + ex.Message;
                    Console.WriteLine(ErrStr);
                    File.AppendAllText(LogFilePath, ErrStr + Environment.NewLine);
                    return;
                }

                Console.WriteLine("Closing Connections...");
            } //Using Source and Dest connections

            //Show num rows added and those errored:
            Console.WriteLine();
            Console.WriteLine("Rows Successfully Transferred: " + RowsTransferred);
            Console.WriteLine("Unrecoverable Errors: " + RowsWithErrors);

            File.AppendAllText(LogFilePath, "End of Run: " + DateTime.Now + Environment.NewLine 
                + "Unrecoverable Error Rows: " + RowsWithErrors + Environment.NewLine 
                + "Rows Transferred: " + RowsTransferred 
                + Environment.NewLine + "Transfer Process Complete" + Environment.NewLine);

            Console.WriteLine("Log File Path: " + LogFilePath);
            Console.WriteLine("DONE!");

        } //Main Function

        static bool ParseConfigurationFile(string ConfigPath)
        {
            //-----------------------------------------------------------------------
            //Parse the configuration file. Loads its settings into static variables.

            try
            {
                foreach (string L in File.ReadLines(ConfigPath))
                {
                    //All fields start with a [ character:
                    if (L.StartsWith('['))
                    {
                        if (L.StartsWith("[["))
                        {
                            Console.WriteLine("Error: Field names that start or end with brackets are not supported. Line: " + L);
                            continue;
                        }

                        try
                        {
                            DataTransCol Col = new DataTransCol();

                            //Find end of source col. Using Watch out for double ]] characters!
                            int EndOfPart = Regex.Match(L, @"(?<!\])\](?!\])").Index + 1;
                            Col.SourceField = L.Substring(0, EndOfPart);

                            //Find range of destination col.
                            string RestOfLine = L.Substring(EndOfPart + 1);
                            int StartOfPart = RestOfLine.IndexOf('[');
                            EndOfPart = Regex.Match(RestOfLine.Substring(StartOfPart + 1), @"(?<!\])\](?!\])").Index + 2;
                            Col.DestinationField = RestOfLine.Substring(StartOfPart, EndOfPart);

                            int EndOfCols = Regex.Matches(L, @"(?<!\])\](?!\])")[1].Index + 1;

                            //Get the Data Type portion of the line:
                            if (L.Substring(EndOfCols).Contains("Default"))
                                EndOfPart = L.Substring(EndOfCols).IndexOf("Default") - 1;
                            else
                                EndOfPart = L.Substring(EndOfCols).Length;

                            string DTypePart = L.Substring(EndOfCols, EndOfPart).Trim().ToLower();

                            if (DTypePart.Contains('('))
                                EndOfPart = DTypePart.IndexOf('(');
                            else if (DTypePart.Contains(' '))
                                EndOfPart = DTypePart.IndexOf(' ');
                            else
                                EndOfPart = DTypePart.Length;

                            Col.OfficialDataType = DTypePart.Substring(0, EndOfPart);

                            switch (Col.OfficialDataType)
                            {
                                case "char":
                                case "nchar":
                                case "ntext":
                                case "nvarchar":
                                case "text":
                                case "varchar":
                                    Col.TransDataTypeRule = BasicDataType.DTString;
                                    break;
                                case "bigint":
                                case "int":
                                case "smallint":
                                case "tinyint":
                                    Col.TransDataTypeRule = BasicDataType.DTWholeNumber;
                                    break;
                                case "decimal":
                                case "float":
                                case "money":
                                case "numeric":
                                case "real":
                                case "smallmoney":
                                    Col.TransDataTypeRule = BasicDataType.DTDecimalNumber;
                                    break;
                                case "datetime":
                                case "datetimeoffset":
                                case "datetime2":
                                case "smalldatetime":
                                    Col.TransDataTypeRule = BasicDataType.DTDateTime;
                                    break;
                                case "date":
                                    Col.TransDataTypeRule = BasicDataType.DTDateOnly;
                                    break;
                                case "time":
                                    Col.TransDataTypeRule = BasicDataType.DTTimeOnly;
                                    break;
                                case "bit":
                                    Col.TransDataTypeRule = BasicDataType.DTBoolean;
                                    break;
                                default:
                                    Col.TransDataTypeRule = BasicDataType.DTOther;
                                    break;
                            }

                            Col.IsNullable = !DTypePart.Contains("notnull");

                            if (Col.TransDataTypeRule == BasicDataType.DTString)
                            {
                                StartOfPart = EndOfPart + 1;
                                EndOfPart = DTypePart.IndexOf(')');
                                string SizePart = DTypePart.Substring(StartOfPart, EndOfPart - StartOfPart);
                                if (SizePart.ToUpper() == "MAX")
                                    Col.StrFieldSize = -1;
                                else
                                    int.TryParse(SizePart, out Col.StrFieldSize);

                                if (DTypePart.Contains("trim:y") || DTypePart.Contains("trim:b"))
                                    Col.StrTrimRule = TrimRule.TRTrimBothSides;
                                else if (DTypePart.Contains("trim:f"))
                                    Col.StrTrimRule = TrimRule.TRTrimFrontOnly;
                                else if (DTypePart.Contains("trim:e"))
                                    Col.StrTrimRule = TrimRule.TRTrimEndOnly;
                                else
                                    Col.StrTrimRule = TrimRule.TRNoTrimming;

                                if (DTypePart.Contains("trunc:y..."))
                                    Col.StrTruncRule = TruncationRule.TRTruncateWithEllipsis;
                                else if (DTypePart.Contains("trunc:n"))
                                    Col.StrTruncRule = TruncationRule.TRNoTruncating;
                                else
                                    Col.StrTruncRule = TruncationRule.TRTruncateString;
                            }
                            else
                            {
                                Col.StrFieldSize = 0;
                                Col.StrTrimRule = TrimRule.TRNoTrimming;
                                Col.StrTruncRule = TruncationRule.TRNoTruncating;
                            }

                            //Get Default Value portion of the line:
                            StartOfPart = L.Substring(EndOfCols).IndexOf("Default");
                            string DefValuePart = "";
                            if (StartOfPart != -1) DefValuePart = L.Substring(EndOfCols + StartOfPart);

                            bool HasDefPart = !string.IsNullOrWhiteSpace(DefValuePart);

                            if (HasDefPart)
                            {
                                if (DefValuePart.Trim().ToLower() == "defaultisnull")
                                {
                                    if (Col.IsNullable)
                                    {
                                        Col.DefaultIsNull = true;
                                        Col.DefaultValue = "";
                                    }
                                    else
                                    {
                                        Console.WriteLine("NOTE: Can not set Null as default value to a non-nullable field: " + Col.DestinationField);
                                        HasDefPart = false;
                                    }
                                }
                                else if (DefValuePart.Contains(":"))
                                {
                                    string[] DParts = DefValuePart.Split(':', 2);
                                    Col.DefaultIsNull = false;
                                    Col.DefaultValue = DParts[1];
                                }
                                else
                                {
                                    HasDefPart = false;
                                }
                            }

                            if (!HasDefPart)
                            {
                                if (Col.IsNullable)
                                {
                                    Col.DefaultIsNull = true;
                                    Col.DefaultValue = "";
                                }
                                else
                                {
                                    Col.DefaultIsNull = false;
                                    switch (Col.TransDataTypeRule)
                                    {
                                        case BasicDataType.DTWholeNumber:
                                        case BasicDataType.DTDecimalNumber:
                                        case BasicDataType.DTBoolean:
                                            Col.DefaultValue = "0";
                                            break;
                                        case BasicDataType.DTDateTime:
                                        case BasicDataType.DTDateOnly:
                                        case BasicDataType.DTTimeOnly:
                                            Col.DefaultValue = "MIN";
                                            break;
                                        default:
                                            Col.DefaultValue = "";
                                            break;
                                    }
                                }
                            } //Doesn't have proper default part.

                            DataTransCols.Add(Col);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error Parsing Field: " + ex.Message);
                            Console.WriteLine(L);
                            return false;
                        }
                    } //Is a field to parse

                    //Other settings should have a colon in them:
                    else if (L.Contains(':'))
                    {
                        string[] LP = L.Split(':', 2);

                        if (LP.Length < 2) continue;

                        string CfgValue = LP[1].Trim();

                        switch (LP[0].ToLower())
                        {
                            case "source connection string":
                                SourceConnStr = CfgValue;
                                Console.WriteLine("Source connection string set.");
                                break;
                            case "source table name":
                                SourceTableName = CfgValue;
                                Console.WriteLine("Source Table Name: " + SourceTableName);
                                break;
                            case "source where clause":
                                WhereClause = CfgValue;
                                Console.WriteLine("Source WHERE clause set.");
                                break;
                            case "destination connection string":
                                DestConnStr = CfgValue;
                                Console.WriteLine("Destination connection string set.");
                                break;
                            case "destination table name":
                                DestTableName = CfgValue;
                                Console.WriteLine("Destination Table Name: " + DestTableName);
                                break;
                            case "identity insert":
                                IDInsert = (CfgValue.ToUpper() == "Y");
                                if (IDInsert) Console.WriteLine("NOTE: Identity Insert is ON!");
                                break;
                            case "stop on error":
                                StopOn1stError = (CfgValue.ToUpper() == "Y");
                                break;
                            case "batch size":
                            case "rows at a time":
                                int.TryParse(CfgValue, out BatchSize);
                                break;
                            case "log file path":
                                LogFilePath = CfgValue;
                                Console.WriteLine("Log File Path: " + LogFilePath);
                                break;
                        } //switch LP[0]

                    } //Is a setting, with a ":" character in it.

                } //for each Line in Config File
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error attempting to parse configuration file: " + ex.Message);
                return false;
            }

            return true;
        } //Parse configuration file

        static bool BasicValidationTasks()
        {
            //-------------------------------------
            //Basic validation tasks:
            if (string.IsNullOrWhiteSpace(SourceConnStr))
            {
                Console.WriteLine("Error: Source connection string is missing.");
                return false;
            }
            if (string.IsNullOrWhiteSpace(SourceTableName))
            {
                Console.WriteLine("Error: Source table name is missing.");
                return false;
            }
            if (string.IsNullOrWhiteSpace(DestConnStr))
            {
                Console.WriteLine("Error: Destination connection string is missing.");
                return false;
            }
            if (string.IsNullOrWhiteSpace(DestTableName))
            {
                Console.WriteLine("Error: Destination table name is missing.");
                return false;
            }
            if (UnquotedSemiColonCheck(WhereClause))
            {
                Console.WriteLine("Error: WHERE clause may contain injection of multiple commands. Remove semi-colon characters (;), that are not wrapped within a pair of single quote marks ('), and try again.");
                return false;
            }
            if (string.IsNullOrWhiteSpace(LogFilePath))
            {
                Console.WriteLine("Error: Log File path is missing.");
                return false;
            }
            if (DataTransCols.Count == 0)
            {
                Console.WriteLine("Error: No fields set up.");
                return false;
            }

            {
                bool HasEmptyColNames = false;
                foreach (DataTransCol C in DataTransCols)
                {

                    if (string.IsNullOrWhiteSpace(C.DestinationField) |
                        string.IsNullOrWhiteSpace(C.SourceField))
                    {
                        Console.WriteLine("Error: Field Name not set for a source or destination column: " + C.DestinationField + C.SourceField);
                        HasEmptyColNames = true;
                    }
                }
                if (HasEmptyColNames) return false;
            }

            if (BatchSize < 1) BatchSize = 1;
            if (BatchSize > 1000) BatchSize = 1000;
            Console.WriteLine("Batch Size (rows inserted at time): " + BatchSize);

            //----------------------------------------------------------
            //Append initial line to log file.
            try
            {
                File.AppendAllText(LogFilePath, "Difficult Data Transfer - Version " + VersionNum + Environment.NewLine
                    + "Run Start: " + DateTime.Now + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error appending to Log File: " + ex.Message);
                return false;
            }

            return true;
        } //Basic validation Tasks

        static bool UnquotedSemiColonCheck(string StrToCheck)
        {
            bool Result = false;

            bool BetweenQuote = false;
            foreach (char C in StrToCheck)
            {
                if (C == ';')
                    if (!BetweenQuote) Result = true;
                if (C == '\'')
                    BetweenQuote = !BetweenQuote;
            }

            return Result;
        }

        static string BuildStartOfINSERTStatement()
        {
            //-------------------------------------------------------
            //Form the start of every INSERT statement:
            StringBuilder InsCmdStart = new StringBuilder();

            StringBuilder LogTableFields = new StringBuilder();
            LogTableFields.AppendLine("Destination Table Fields:");

            string DstTable = DestTableName;
            //Put brackets around destination table name, if it doesn't already have them, and if there either: the name requires it or if there is no owner set to the table.
            if (!DstTable.StartsWith("[") && !DstTable.EndsWith("]") && (!DstTable.Contains(".") || DstTable.Contains(" ") || DstTable.Contains("'")))
                DstTable = "[" + DstTable.Replace("]", "]]") + "]";

            if (IDInsert) InsCmdStart.AppendLine("SET IDENTITY_INSERT " + DstTable + " ON;");
            InsCmdStart.Append("INSERT INTO " + DstTable + " (");
            bool IsFirstCol = true;
            foreach (DataTransCol Col in DataTransCols)
            {
                if (IsFirstCol)
                    IsFirstCol = false;
                else
                    InsCmdStart.Append(", ");

                InsCmdStart.Append(Col.DestinationField.Replace("'", "''"));

                LogTableFields.Append(Col.DestinationField + " " + Col.OfficialDataType);
                if (Col.TransDataTypeRule == BasicDataType.DTString) LogTableFields.Append("(" + Col.StrFieldSize + ")");
                if (Col.IsNullable == false) LogTableFields.Append(" NOT NULL");
                
                if (Col.TransDataTypeRule == BasicDataType.DTOther)
                {
                    Console.WriteLine("Unsupported Data Type: " + Col.OfficialDataType);
                    LogTableFields.Append("  Unsupported Data Type");
                }

                LogTableFields.AppendLine();
            }
            InsCmdStart.AppendLine(") VALUES ");

            File.AppendAllText(LogFilePath, LogTableFields.ToString());

            return InsCmdStart.ToString();
        } //Build start of INSERT statement

        static string BuildEndOfINSERTStatement()
        {

            //-------------------------------------------------------
            //Form the ending of every INSERT statement:
            if (IDInsert)
                return "; SET IDENTITY_INSERT [" + DestTableName.Replace("'", "''") + "] OFF;";
            else
                return ";";

        } //Set ending of INSERT statement.

        static string BuildSELECTStatementOfSource()
        {
            StringBuilder SBSrcCmd = new StringBuilder("SELECT ");
            bool IsFirstCol = true;
            foreach (DataTransCol C in DataTransCols)
            {
                if (IsFirstCol)
                    IsFirstCol = false;
                else
                    SBSrcCmd.Append(", ");

                SBSrcCmd.Append(C.SourceField.Replace("'", "''"));
            } //for each column

            string SrcTable = SourceTableName;
            //Put brackets around source table name, if it doesn't already have them, and if there either: the name requires it or if there is no owner set to the table.
            if (!SrcTable.StartsWith("[") && !SrcTable.EndsWith("]") && (!SrcTable.Contains(".") || SrcTable.Contains(" ") || SrcTable.Contains("'")))
                SrcTable = "[" + SrcTable.Replace("]", "]]") + "]";

            SBSrcCmd.Append(" FROM " + SrcTable);

            if (!string.IsNullOrWhiteSpace(WhereClause))
            {
                SBSrcCmd.Append(" WHERE " + WhereClause);
            }

            Console.WriteLine("Source SQL: " + SBSrcCmd.ToString());
            File.AppendAllText(LogFilePath, "Source SQL: " + SBSrcCmd.ToString() + Environment.NewLine);

            return SBSrcCmd.ToString();
        } //Build SELECT statement

        static string BuildINSERTValue(OleDbDataReader TheReader)
        {
            StringBuilder SBRowValues = new StringBuilder("(");
            bool IsFirstCol = true;


            foreach (DataTransCol Col in DataTransCols)
            {
                if (IsFirstCol)
                    IsFirstCol = false;
                else
                    SBRowValues.Append(", ");

                string SrcField = Col.SourceField.Replace("]]","]");
                if (SrcField.StartsWith("[")) SrcField = SrcField.Substring(1);
                if (SrcField.EndsWith("]")) SrcField = SrcField.Substring(0, SrcField.Length - 1);

                var OrigSourceValue = TheReader[SrcField];
                string DestValue = string.Empty;

                //  For each column in the row:
                //      Test if data will fit. If so, add that value to VALUES chain. If not, use default value. Log any truncations or significant alterations of values.

                //If you need to add code that will operate on a value, such as conversion mathematics, you can apply that, here.

                //If you are merging the value of two or more fields, you can achieve that by storing the OrigSourceValue in a separate variable when you reach a specific column name, \
                // and then a "continue" command to skip the rest of this and move onto the next column.

                if ((OrigSourceValue == DBNull.Value || OrigSourceValue is null) && Col.IsNullable)
                {
                    DestValue = "NULL";
                }
                else //Original value is not null:
                {
                    string OrigSourceValStr = (OrigSourceValue + "").Trim();
                    switch (Col.TransDataTypeRule)
                    {
                        case BasicDataType.DTBoolean:
                            switch (OrigSourceValStr.ToLower())
                            {
                                case "1":
                                case "true":
                                case "t":
                                    DestValue = "1";
                                    break;
                                case "0":
                                case "false":
                                case "f":
                                    DestValue = "0";
                                    break;
                                default:
                                    if (Col.DefaultIsNull & Col.IsNullable)
                                        DestValue = "NULL";
                                    else if (Col.DefaultValue == "1" || Col.DefaultValue == "0")
                                        DestValue = Col.DefaultValue;
                                    else
                                        DestValue = "0";
                                    break;
                            }
                            break;
                        case BasicDataType.DTDateTime:
                            bool B1 = DateTime.TryParse(OrigSourceValStr, out DateTime DT);

                            if (B1)
                                DestValue = "'" + OrigSourceValStr + "'";
                            else if ((Col.DefaultIsNull | Col.DefaultValue.Trim().ToUpper() == "NULL") && Col.IsNullable)
                                DestValue = "NULL";
                            else if (Col.DefaultValue.Trim().ToUpper() == "NOW")
                                DestValue = "'" + DateTime.Now + "'";
                            else if (Col.DefaultValue.Trim().ToUpper() == "MIN")
                            {
                                switch (Col.OfficialDataType)
                                {
                                    case "datetime":
                                        DestValue = "'1/1/1753 12:00:00'";
                                        break;
                                    case "datetime2":
                                        DestValue = "'1/1/0001 12:00:00'";
                                        break;
                                    case "smalldatetime":
                                        DestValue = "'1/1/1900 12:00:00'";
                                        break;
                                    default:
                                        DestValue = "'1/1/1900 12:00:00'";
                                        break;
                                }
                            }
                            else if (DateTime.TryParse(Col.DefaultValue, out DateTime DTemp))
                                DestValue = "'" + Col.DefaultValue + "'";
                            else
                                DestValue = "'1/1/1900 12:00:00'";
                            break;
                        case BasicDataType.DTDateOnly:
                            bool B2 = DateOnly.TryParse(OrigSourceValStr, out DateOnly DOn);

                            if (!B2)
                            {
                                B2 = DateTime.TryParse(OrigSourceValStr, out DateTime DTemp2);
                                OrigSourceValStr = DateOnly.FromDateTime(DTemp2).ToString();
                            }

                            if (B2)
                                DestValue = "'" + OrigSourceValStr + "'";
                            else if ((Col.DefaultIsNull || Col.DefaultValue.Trim().ToUpper() == "NULL") && Col.IsNullable)
                                DestValue = "NULL";
                            else if (Col.DefaultValue.Trim().ToUpper() == "NOW")
                                DestValue = "'" + DateOnly.FromDateTime(DateTime.Now) + "'";
                            else if (Col.DefaultValue.Trim().ToUpper() == "MIN")
                                DestValue = "'1/1/0001'";
                            else
                            {
                                if (DateTime.TryParse(Col.DefaultValue, out DateTime DTemp))
                                    DestValue = "'" + DateOnly.FromDateTime(DTemp) + "'";
                                else
                                    DestValue = "'1/1/0001'";
                            }
                            break;
                        case BasicDataType.DTTimeOnly:
                            bool B3 = TimeOnly.TryParse(OrigSourceValStr, out TimeOnly TOn);

                            if (!B3)
                            {
                                B3 = DateTime.TryParse(OrigSourceValStr, out DateTime DTemp2);
                                OrigSourceValStr = TimeOnly.FromDateTime(DTemp2).ToString();
                            }

                            if (B3)
                                DestValue = "'" + OrigSourceValStr + "'";
                            else if ((Col.DefaultIsNull || Col.DefaultValue.Trim().ToUpper() == "NULL") && Col.IsNullable)
                                DestValue = "NULL";
                            else if (Col.DefaultValue.Trim().ToUpper() == "NOW")
                                DestValue = "'" + TimeOnly.FromDateTime(DateTime.Now) + "'";
                            else if (Col.DefaultValue.Trim().ToUpper() == "MIN")
                                DestValue = "'12:00:00'";
                            else
                            {
                                if (DateTime.TryParse(Col.DefaultValue, out DateTime DTemp))
                                    DestValue = "'" + TimeOnly.FromDateTime(DTemp) + "'";
                                else
                                    DestValue = "'12:00:00'";
                            }
                            break;
                        case BasicDataType.DTDecimalNumber:
                            bool DTry = decimal.TryParse(OrigSourceValStr.Replace(",", ""), out decimal DVal);

                            if (DTry)
                                DestValue = DVal.ToString();
                            else
                            {
                                Match M1 = Regex.Match(OrigSourceValStr.Replace(",", "").Replace("-", ""), @"^\d*\.?\d+");

                                if (M1.Success && decimal.TryParse(M1.Value, out DVal))
                                {
                                    DestValue = DVal.ToString();
                                    if (OrigSourceValStr.StartsWith("-")) DestValue = "-" + DestValue;

                                    File.AppendAllText(LogFilePath, "Decimal value had invalid characters removed in Row #" + CurrentRowNum
                                        + " Column:" + Col.DestinationField + " Value: " + OrigSourceValStr + Environment.NewLine);
                                }
                                else if (Col.DefaultIsNull & Col.IsNullable)
                                    DestValue = "NULL";
                                else if (decimal.TryParse(Col.DefaultValue, out DVal))
                                    DestValue = Col.DefaultValue;
                                else
                                    DestValue = "0";
                            }

                            if (OrigSourceValStr.Contains(","))
                                File.AppendAllText(LogFilePath, "NOTE: Original decimal value contained comma characters in Row #" + CurrentRowNum 
                                    + " Column:" + Col.DestinationField + Environment.NewLine);

                            break;
                        case BasicDataType.DTWholeNumber:
                            bool ITry = long.TryParse(OrigSourceValStr.Replace(",", ""), out long IVal);
                            if (ITry)
                                DestValue = IVal.ToString();
                            else
                            {
                                Match M2 = Regex.Match(OrigSourceValStr.Replace(",", "").Replace("-", ""), @"^\d+");

                                if (M2.Success && long.TryParse(M2.Value, out IVal))
                                {
                                    DestValue = IVal.ToString();
                                    if (OrigSourceValStr.StartsWith("-")) DestValue = "-" + DestValue;

                                    File.AppendAllText(LogFilePath, "Integer value had invalid characters removed in Row #" + CurrentRowNum
                                        + " Column:" + Col.DestinationField + " Value: " + OrigSourceValStr + Environment.NewLine);
                                }
                                else if (Col.DefaultIsNull & Col.IsNullable)
                                    DestValue = "NULL";
                                else if (long.TryParse(Col.DefaultValue, out IVal))
                                    DestValue = Col.DefaultValue;
                                else
                                    DestValue = "0";
                            }

                            if (OrigSourceValStr.Contains("."))
                                File.AppendAllText(LogFilePath, "NOTE: Decimal places possibly lost in Row #" + CurrentRowNum 
                                    + " Column:" + Col.DestinationField + " Value: " + OrigSourceValStr + Environment.NewLine);
                            if (OrigSourceValStr.Contains(","))
                                File.AppendAllText(LogFilePath, "NOTE: Original integer value contained comma characters in Row #" + CurrentRowNum 
                                    + " Column:" + Col.DestinationField + " Value: " + OrigSourceValStr + Environment.NewLine);

                            break;
                        case BasicDataType.DTString:
                            bool DestIsNULL = false;

                            if (OrigSourceValue == DBNull.Value || OrigSourceValue is null)
                            {
                                DestValue = Col.DefaultValue;
                            }
                            else
                            {
                                switch (Col.StrTrimRule)
                                {
                                    case TrimRule.TRNoTrimming:
                                        DestValue = OrigSourceValue + ""; //Original source, before it was trimmed.
                                        break;
                                    case TrimRule.TRTrimFrontOnly:
                                        DestValue = (OrigSourceValue + "").TrimStart();
                                        break;
                                    case TrimRule.TRTrimEndOnly:
                                        DestValue = (OrigSourceValue + "").TrimEnd();
                                        break;
                                    case TrimRule.TRTrimBothSides:
                                        DestValue = OrigSourceValStr; //already trimmed at both ends.
                                        break;
                                    default:
                                        DestValue = OrigSourceValue + "";
                                        break;
                                } //trim rule
                            }
                            
                            if (Col.StrFieldSize > 0 & DestValue.Length > Col.StrFieldSize)
                            {
                                switch (Col.StrTruncRule)
                                {
                                    case TruncationRule.TRTruncateString:
                                        DestValue = DestValue.Substring(0, Col.StrFieldSize);
                                        break;
                                    case TruncationRule.TRTruncateWithEllipsis:
                                        if (Col.StrFieldSize > 3) //Only do the ellipsis if there is more than 3 characters in the field size.
                                            DestValue = DestValue.Substring(0, Col.StrFieldSize - 3) + "...";
                                        else
                                            DestValue = DestValue.Substring(0, Col.StrFieldSize);
                                        break;
                                    default: //No truncating
                                        if (Col.DefaultIsNull & Col.IsNullable)
                                        {
                                            DestValue = "NULL";
                                            DestIsNULL = true;
                                        }
                                        else
                                        {
                                            DestValue = Col.DefaultValue;
                                            if (DestValue.Length > Col.StrFieldSize)
                                                DestValue = DestValue.Substring(0, Col.StrFieldSize);
                                        }
                                        break;
                                } //truncation rule

                                File.AppendAllText(LogFilePath, "String Value Truncated in Row #" + CurrentRowNum + " Column:" + Col.DestinationField + Environment.NewLine);
                                File.AppendAllText(LogFilePath, "String Value: " + OrigSourceValue + Environment.NewLine);
                            } //If too long to fit.

                            if (!DestIsNULL) DestValue = "'" + DestValue + "'";

                            break;
                        default: //Other
                            DestValue = "'" + OrigSourceValStr + "'";

                            break;
                    } //Data Type Rule
                } //Not Null

                //Embed or Remove Single Quote Marks:
                if (DestValue.StartsWith("'") && DestValue.EndsWith("'"))
                {
                    string MidVal = DestValue.Substring(1, DestValue.Length - 2);
                    DestValue = "'" + MidVal.Replace("'", "''") + "'";
                }
                else
                {
                    DestValue.Replace("'", ""); //remove from fields that shant use them.
                }

                //If you are splitting the value of a field into two or more columns, you can add logic to do that here. Append to SBRowValues multiple times, separated by commas, if you need to.

                SBRowValues.Append(DestValue);
            } //For each Column

            SBRowValues.AppendLine(")"); //Not putting commas here, because we'll add those later.

            return SBRowValues.ToString();
        } //Build VALUE string for INSERT command

        static bool ExecuteINSERTStatement(List<string> InsertValues, OleDbConnection DestCn)
        {
            //Return false to END the program, if we shouldn't do any further insertions.
            //Do NOT return false if an error occured, but we SHOULD try to continue further with other insertions, anyway.

            bool DoSingleRows = (BatchSize == 1);

            if (!DoSingleRows)
            {
                StringBuilder FullInsertCommand = new StringBuilder(InsCmdStart.ToString());
                bool IsFirstCol = true;
                foreach (string V in InsertValues)
                {
                    if (IsFirstCol)
                        IsFirstCol = false;
                    else
                        FullInsertCommand.Append(", ");

                    FullInsertCommand.Append(V);
                }
                FullInsertCommand.Append(InsCmdEnd);

                //Upon Errors:
                //If 'stop on error' is set, stop the process. And log it. Show collection of values trying to go in.
                //If not set, then:
                //  Break apart collection of VALUES and try each one at a time.
                //      If error on one row, continue with the next and log it.
                //      If the first 25 rows error out, then stop the process, anyway.

                try
                {
                    OleDbCommand DestInsCmd = new OleDbCommand(FullInsertCommand.ToString(), DestCn);
                    DestInsCmd.ExecuteNonQuery();
                    AtLeastOneSuccess = true;
                    RowsTransferred += InsertValues.Count;
                    Console.WriteLine("Rows Transferred So Far: " + RowsTransferred);
                }
                catch (Exception ex)
                {
                    if (StopOn1stError)
                    {
                        string ErrStr = "ERROR Inserting Rows: " + ex.Message + Environment.NewLine + FullInsertCommand.ToString();
                        Console.WriteLine(ErrStr);
                        File.AppendAllText(LogFilePath, ErrStr + Environment.NewLine);
                        Console.WriteLine("Stop On First Error Requested.");
                        return false;
                    }
                    else
                    {
                        DoSingleRows = true;
                    }
                }
            } //NOT Doing Single Rows, yet.

            if (DoSingleRows)
            {
                foreach (string V in InsertValues)
                {
                    string SingleInsertCmd = InsCmdStart.ToString() + V + InsCmdEnd;

                    try
                    {
                        OleDbCommand DestInsCmd = new OleDbCommand(SingleInsertCmd, DestCn);
                        DestInsCmd.ExecuteNonQuery();
                        AtLeastOneSuccess = true;
                        RowsTransferred++;

                        if (RowsTransferred % 1000 == 0) 
                            Console.WriteLine("Rows Transferred So Far: " + RowsTransferred);
                    }
                    catch (Exception ex)
                    {
                        string ErrStr = "ERROR Inserting Row: " + ex.Message + Environment.NewLine + SingleInsertCmd;
                        Console.WriteLine(ErrStr);
                        File.AppendAllText(LogFilePath, ErrStr + Environment.NewLine);
                        RowsWithErrors++;

                        if (StopOn1stError)
                        {
                            Console.WriteLine("Stop On First Error Requested.");
                            return false;
                        }

                        if (!AtLeastOneSuccess & RowsWithErrors >= 100)
                        {
                            Console.WriteLine("Transfer stopped because nothing was going in, at all, for a while!");
                            File.AppendAllText(LogFilePath, "Transfer Stopped." + Environment.NewLine);
                            return false;
                        }
                    }
                }
            } //Try Doing Single Rows instead of bundle
            Thread.Sleep(250);
            return true;
        } //Execute INSERT statement!

        //Perhaps a future version will encrypt connection strings?

    } //Program class
} //namespace
