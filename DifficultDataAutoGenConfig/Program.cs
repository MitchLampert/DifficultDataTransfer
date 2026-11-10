using System.Data;
using System.Data.OleDb;
using System.Reflection;
using System.Text;

/*
This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or distribute this software, either in source code form or as a compiled binary, for any purpose, commercial or non-commercial, and by any means.
In jurisdictions that recognize copyright laws, the author or authors of this software dedicate any and all copyright interest in the software to the public domain. We make this dedication for the benefit of the public at large and to the detriment of our heirs and successors. We intend this dedication to be an overt act of relinquishment in perpetuity of all present and future rights to this software under copyright law.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
//Original Author: Mitchell Scott Lampert

namespace DifficultDataAutoConfigGen
{
    internal struct DDTableField
    {
        public string FieldName;
        public string DataType;
        public int StringSize;
        public bool IsNullable;
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            //This source code is not intended to look too impressive to code reviewers. It just gets its simple job done in a very linear fashion, that's all.
            //The source for the main program, DifficultDataTransfer.exe, might be more interesting to look at.

            Version? VersionNum = Assembly.GetEntryAssembly()?.GetName().Version;

            Console.WriteLine("Mitchell's Difficult Data Automated Configuration Generator");
            Console.WriteLine("-----------------------------------------------------------");
            Console.WriteLine("Version " + VersionNum);
            Console.WriteLine();
            Console.WriteLine("This program writes a configuration file for use with the Difficult Data Transfer program. This app does NOT do any actual data transferring, itself. But, it will need to examine a table on a destination database connection, at least. Details are available in the instruction manual.");

            Console.WriteLine();
            Console.WriteLine("Configuration files usually end with a .DifficultDataConfig extension, but it doesn't really matter that much. It's just going to be a plain text file, anyway.");
            Console.Write("File Path to Write Configuration File: ");
            string ConfigFinalPath = (Console.ReadLine() + "").Trim();

            if (string.IsNullOrEmpty(ConfigFinalPath))
            {
                Console.WriteLine("No path entered. We need somewhere to write the configuration file to.");
                return;
            }

            try
            {
                File.WriteAllText(ConfigFinalPath, "TEST");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error trying to test file location: " + ex.Message);
                return;
            }

            Console.WriteLine();
            Console.Write("Destination Connection String (OLEDB): ");
            string CString = Console.ReadLine() + "";

            if (string.IsNullOrEmpty(CString))
            {
                Console.WriteLine("No connection string entered. Sorry, this won't work unless we connect to a database.");
                return;
            }

            string DestTableName = string.Empty;
            List<DDTableField> DDTableFields = new List<DDTableField>();

            using (OleDbConnection Cn = new OleDbConnection(CString))
            {
                try
                {
                    Cn.Open();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Could not open connection to database: " + ex.Message);
                    return;
                }

                try
                {
                    OleDbCommand CmdTables = new OleDbCommand("SELECT table_name FROM information_schema.tables ORDER BY table_name", Cn);
                    DataTable TblTables = new DataTable();
                    TblTables.Load(CmdTables.ExecuteReader());
                    Console.WriteLine();
                    Console.WriteLine("Tables Found:");
                    foreach (DataRow R in TblTables.Rows) Console.WriteLine(R["table_name"] + "");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error Getting Table Names: " + ex.Message);
                    return;
                }

                Console.WriteLine();
                Console.WriteLine("For table names, brackets [] are not usually required. However, if the table name contains spaces or special characters, we recommend using them." +
                    " You can also include owner or schema names, such as dbo., if needed.");


                Console.WriteLine();
                Console.Write("Destination Table Name: ");
                DestTableName = Console.ReadLine() + "";

                if (string.IsNullOrEmpty(DestTableName))
                {
                    Console.WriteLine("No table entered, for some reason. Oh well, I guess we can't continue, then. Sorry.");
                    return;
                }

                try
                {
                    OleDbCommand CmdFields = new OleDbCommand($"SELECT column_name, data_type, character_maximum_length, is_nullable FROM information_schema.columns " +
                        $"WHERE table_name = '{DestTableName.Replace("'", "''")}' ORDER BY ordinal_position;", Cn);
                    DataTable TblFields = new DataTable();
                    TblFields.Load(CmdFields.ExecuteReader());

                    Console.WriteLine();
                    Console.WriteLine($"Fields Found in Table {DestTableName}:");

                    foreach (DataRow R in TblFields.Rows)
                    {
                        DDTableField F = new DDTableField();

                        if (int.TryParse(R["character_maximum_length"] + "", out int StrSize))
                            F.StringSize = StrSize;
                        else
                            F.StringSize = 0;

                        F.FieldName = R["column_name"] + "";
                        F.DataType = (R["data_type"] + "").ToLower();

                        string IsNullStr = R["is_nullable"] + "";
                        F.IsNullable = (IsNullStr.StartsWith("Y") | IsNullStr.StartsWith("T") | IsNullStr == "1");

                        Console.Write(F.FieldName + " (" + F.DataType);
                        if (F.StringSize > 0)
                            Console.WriteLine("(" + F.StringSize + "))");
                        else if (F.StringSize == -1)
                            Console.WriteLine("(MAX))");
                        else
                            Console.WriteLine(")");

                        DDTableFields.Add(F);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error Getting Field Names: " + ex.Message);
                    return;
                }
            } //Using database connection

            Console.WriteLine();
            Console.WriteLine("For this prompt, enter a numeric value, or enter the phrase NULL. Default value will be 0 if you don't enter anything, at all. It can be a negative number or have decimals, too.");
            Console.Write("Default Numeric Value Upon Error: ");
            string DefNumStr = Console.ReadLine() + "";
            bool DefNumIsNull = (DefNumStr.Trim().ToUpper() == "NULL");
            float DefNum = 0;
            if (!DefNumIsNull)
                float.TryParse(DefNumStr, out DefNum);

            Console.WriteLine();
            Console.WriteLine("For this value, enter any string. The special phrase NULL will imply you want the default value to be null. (If you want the actual word to be NULL, put it in single quotes: 'NULL'.)");
            Console.WriteLine("Default value will be an empty string, if you don't enter anything at all.");
            Console.Write("Default String Value Upon Error: ");
            string DefStrStr = Console.ReadLine() + "";
            bool DefStrIsNull = (DefStrStr.Trim().ToUpper() == "NULL");
            string DefStr = "";
            if (!DefStrIsNull)
                DefStr = DefStrStr;
            if (DefStr.ToUpper() == "'NULL'") DefStr = DefStrStr.Replace("'", "");

            Console.WriteLine();
            Console.WriteLine("For this value, either enter a valid date and/or time, or one of these words: NOW, MIN, NULL. Default value will be 1/1/1900 12:00, if you don't enter anything at all.");
            Console.Write("Default Date/Time Value Upon Error: ");
            string DefDateStr = Console.ReadLine() + "";
            if (!DateTime.TryParse(DefDateStr, out DateTime DT))
            {
                if (new string[] { "NOW", "MIN", "NULL" }.Contains(DefDateStr.Trim().ToUpper()))
                    DefDateStr = DefDateStr.Trim().ToUpper();
                else
                    DefDateStr = "1/1/1900 12:00:00";
            }

            Console.WriteLine();
            Console.WriteLine("For this value, enter either: 1 or 0 or NULL. Default is 0.");
            Console.Write("Default Boolean/Bit Value Upon Error: ");
            string DefBolStr = (Console.ReadLine() + "").Trim().ToUpper();
            bool DefBolIsNull = (DefBolStr == "NULL");
            bool DefBol = (DefBolStr == "1");

            Console.WriteLine();
            Console.WriteLine("Trim String Rule Values: Y or B = Trim Both Ends, F = Front Only, E = End Only, N = None. Default value is to trim both ends.");
            Console.Write("Trim Strings? ");
            string TrimStr = (Console.ReadLine() + "").Trim().ToUpper();
            if (!(new string[] { "Y", "F", "E", "N" }.Contains(TrimStr))) TrimStr = "Y";

            Console.WriteLine();
            Console.WriteLine("Answer Y or N. If you say no, the Default String Value is used in the field, instead. Either way, the full value will be entered into the log when they happen.  Default is Y.");
            Console.Write("Allow Truncation of Strings Too Long to Fit? ");
            string TruncStr = (Console.ReadLine() + "").Trim().ToUpper();
            bool DoTrunc = !(TruncStr == "N");

            bool EllipsisUponTrunc = false;
            if (DoTrunc)
            {
                Console.Write("Add Ellipsis (...) at end of truncated string? Y or N (default is N)? ");
                string ElipStr = (Console.ReadLine() + "").Trim().ToUpper();
                EllipsisUponTrunc = (ElipStr.StartsWith("Y"));
            }

            Console.WriteLine();
            Console.Write("Allow Identity Insert? Not recommended, unless you know you need it! Y or N (default is N): ");
            string IdStr = (Console.ReadLine() + "").Trim().ToUpper();
            bool AllowIDIns = (IdStr.StartsWith("Y"));

            if (AllowIDIns)
            {
                Console.Write("Are you sure? Y or N (default is N): ");
                IdStr = (Console.ReadLine() + "").Trim().ToUpper();
                AllowIDIns = (IdStr.StartsWith("Y"));
            }

            Console.WriteLine();
            Console.WriteLine("An unrecoverable error is one where the transfer program can not resolve how to fit data into a field, on its own, for a specific row. This rarely happens, given how flexible the system is. But, when it does: You can choose to stop the entire process upon the first such error, or continue trying the rest of the rows. (The program might stop automatically, anyway, if it sees too many such errors happening at the beginning of the process.) All errors are entered to the log, regardless.");
            Console.Write("Stop Upon First Unrecoverable Error? Y or N (default is N): ");
            string StopErStr = (Console.ReadLine() + "").Trim().ToUpper();
            bool StopUponError = (StopErStr.StartsWith("Y"));

            Console.WriteLine();
            Console.WriteLine("For this next prompt: Enter a number from 1 to 1000 for the Batch Size. For most servers, 100 is usually optimal, which is the default if you don't enter anything. If you know the data is very problematic, it might be best to enter 1.");
            Console.Write("Batch Size: ");
            string BatchSizeStr = Console.ReadLine() + "";
            int BatchSize = 100;
            int.TryParse(BatchSizeStr, out BatchSize);
            if (BatchSize < 1) BatchSize = 100;
            if (BatchSize > 1000) BatchSize = 1000;

            Console.WriteLine();
            Console.WriteLine("The Source Connection String will NOT be tested in this application. You can even leave it blank, for now. But, it will need to be filled in before you run the main program.");
            Console.Write("Source Connection String (OLEDB): ");
            string SourceCnStr = Console.ReadLine() + "";

            Console.Write($"Source Table Name (Default will be: {DestTableName}): ");
            string SourceTableName = Console.ReadLine() + "";
            if (string.IsNullOrWhiteSpace(SourceTableName)) SourceTableName = DestTableName;

            Console.Write("Source WHERE Clause (Optional. Do NOT include the word WHERE): ");
            string WhereClause = Console.ReadLine() + "";

            if (UnquotedSemiColonCheck(WhereClause))
                Console.WriteLine("WARNING: Where clause contains semicolon characters (outside of quoted values), which could be used to inject multiple commands into the SQL command. This clause might be rejected by the main program!");

            Console.WriteLine();
            Console.Write("Log File Location Path (optional): ");
            string LogFilePath = Console.ReadLine() + "";

            Console.WriteLine();
            Console.WriteLine("Writing Configuration File....");

            //===================================================================================================================================================================

            StringBuilder SB = new StringBuilder("Mitchell's Difficult Data Transfer - Version " + VersionNum + Environment.NewLine);

            SB.AppendLine("Source Connection String: " + SourceCnStr);
            SB.AppendLine("Source Table Name: " + SourceTableName);

            if (!string.IsNullOrWhiteSpace(WhereClause))
            {
                SB.AppendLine("Source Where Clause: " + WhereClause);
            }

            SB.AppendLine("Destination Connection String: " + CString);
            SB.AppendLine("Destination Table Name: " + DestTableName);

            SB.Append("Identity Insert: ");
            if (AllowIDIns) SB.AppendLine("Y"); else SB.AppendLine("N");

            SB.Append("Stop On Error: ");
            if (StopUponError) SB.AppendLine("Y"); else SB.AppendLine("N");

            SB.AppendLine("Batch Size:" + BatchSize);

            SB.AppendLine("Log File Path: " + LogFilePath);

            SB.AppendLine("Field Maps Format: [SourceField] To [DestinationField] DataType StringRules DefaultValue");
            foreach (DDTableField F in DDTableFields)
            {
                SB.Append("[" + F.FieldName.Replace("]", "]]") + "] To [" + F.FieldName.Replace("]", "]]") + "] " + F.DataType);

                bool IsAString = (new string[] { "char", "nchar", "ntext", "nvarchar", "text", "varchar" }.Contains(F.DataType));
                if (IsAString)
                {
                    if (F.StringSize == -1)
                        SB.Append("(MAX)");
                    else
                        SB.Append("(" + F.StringSize + ")");
                }

                if (!F.IsNullable) SB.Append(" NotNULL");

                if (new string[] { "bigint", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinyint"}.Contains(F.DataType))
                {
                    if (F.IsNullable & DefNumIsNull)
                        SB.AppendLine(" DefaultIsNULL");
                    else
                        SB.AppendLine(" Default:" + DefNum);
                }
                else if (IsAString)
                {
                    SB.Append(" Trim:" + TrimStr);

                    SB.Append(" Trunc:");
                    if (DoTrunc)
                    {
                        if (EllipsisUponTrunc)
                            SB.Append("Y...");
                        else
                            SB.Append("Y");
                    }
                    else SB.Append("N");

                    if (F.IsNullable & DefStrIsNull)
                        SB.AppendLine(" DefaultIsNULL");
                    else
                        SB.AppendLine(" Default:" + DefStr);
                }
                else if (new string[] { "date", "datetime", "datetimeoffset", "datetime2", "smalldatetime", "time" }.Contains(F.DataType))
                {
                    if (DefDateStr.ToUpper() == "NULL")
                    {
                        if (F.IsNullable)
                            SB.AppendLine(" DefaultIsNULL");
                        else
                            SB.AppendLine(" Default:MIN");
                    }
                    else SB.AppendLine(" Default:" + DefDateStr);
                }
                else if (F.DataType == "bit")
                {
                    if (F.IsNullable & DefBolIsNull)
                        SB.AppendLine(" DefaultIsNULL");
                    else if (DefBol)
                        SB.AppendLine(" Default:1");
                    else
                        SB.AppendLine(" Default:0");
                }
                else //image, binary, varbinary, json, uniqueidentifier, xml, etc.
                {
                    Console.WriteLine("Warning: The main program might not support this data type: " + F.DataType);
                    SB.AppendLine();
                }
            } //For each field

            try
            {
                File.WriteAllText(ConfigFinalPath, SB.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error trying to write to file location: " + ex.Message);
                Console.WriteLine("Press Enter to output contents to screen, instead.");
                Console.ReadLine();
                Console.Write(SB.ToString());
                return;
            }

            Console.WriteLine();
            Console.WriteLine("DONE!");
            Console.WriteLine("Tasks To Do Next:");
            Console.WriteLine("* Please check over the names of the source fields and change any of them, if you need to, to match the actual source!");
            Console.WriteLine("* You can also change the default values and text rules of individual fields if you want to.");
            Console.WriteLine("* Don't forget to fill in any settings that you left blank!");

        } //Main function

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
        } //unquoted semi-colons check

        //For a future version, we could connect to the source database and attempt to match up fields to the destination automatically. Look for fields with similar names, and such.

    } //Program class
} //namespace
