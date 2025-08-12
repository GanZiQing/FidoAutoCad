using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using System.Windows;
using System.Windows.Forms;
using Action = System.Action;
using Application = Microsoft.Office.Interop.Excel.Application;
using Button = System.Windows.Forms.Button;
//using Microsoft.Office.Tools.Excel;



namespace FidoAutoCad
{
    static class CommonUtilities
    {
        #region Read Data from Excel
        #region Doubles
        public static double ReadDoubleFromCell(Range cell, bool emptyIsZero = false)
        {
            if (cell.Value2 is double)
            {
                return cell.Value2;
            }
            else if (cell.Value2 == null || cell.Text == "")
            {
                if (emptyIsZero)
                {
                    return 0;
                }
                else
                {
                    return double.NaN;
                }
            }
            else
            {
                try
                {
                    return double.Parse(cell.Value2);
                }
                catch
                {
                    ThrowExceptionBox($"Unable to parse value {cell.Value2} at {cell.Worksheet.Name}!{cell.Address}");
                    throw new Exception();
                }
            }
        }

        public static double ReadDoubleFromCell2(Range cell, double emptyValue = 0, double errorValue = double.NaN)
        {
            if (cell.Value2 is double)
            {
                return cell.Value2;
            }
            else if (cell.Value2 == null || cell.Text == "")
            {
                return emptyValue;
            }
            else
            {
                try
                {
                    return double.Parse(cell.Value2);
                }
                catch
                {
                    return errorValue;
                }
            }
        }

        public static double[] GetContentsAsDoubleArray(Range range, double emptyValue = 0, double errorValue = double.NaN)
        {
            double[] output = new double[range.Cells.Count];

            for (int i = 0; i < range.Cells.Count; i++)
            {
                Range cell = range.Cells[i + 1];
                var cellValue = cell.Value2;
                output[i] = ReadDoubleFromCell2(cell, emptyValue, errorValue);
            }

            return output;
        }

        //public static double GetDoubleFromObject(object item)
        //{
        //    if (item is double)
        //    {
        //        return (double)item;
        //    }
        //    else
        //    {
        //        try
        //        {
        //            return double.Parse(item.ToString());
        //        }
        //        catch
        //        {
        //            ThrowExceptionBox($"Unable to parse value {item} to double");
        //            throw new Exception();
        //        }
        //    }
        //}
        #endregion

        #region String
        public static HashSet<string> GetContentsAsStringHash(Range range)
        {
            List<string> rangeList = GetContentsAsStringList(range, true);
            return new HashSet<string>(rangeList);
        }

        public static string[] GetContentsAsStringArray(Range range, bool ignoreEmpty)
        {
            List<string> rangeList = GetContentsAsStringList(range, ignoreEmpty);
            return rangeList.ToArray();
        }

        public static List<string> GetContentsAsStringList(Range range, bool ignoreEmpty)
        {
            List<string> rangeList = new List<string>();
            foreach (Range cell in range.Cells)
            {
                if (cell.Value2 != null && cell.Value2.ToString() != "")
                {
                    rangeList.Add(cell.Value2.ToString());
                }
                else if (!ignoreEmpty)
                {
                    rangeList.Add("");
                }
            }
            return rangeList;
        }

        public static string GetContentsAsString(Range range, string emptyValue = "")
        {
            if (range.Value2 == null)
            {
                return emptyValue;
            }
            else if (range.Cells.Count > 1)
            {
                throw new Exception($"Unable to convert range {range.Worksheet.Name}!{range.AddressLocal} to single string.\nNumber of cells in range expected is 1, number of cells found is {range.Cells.Count}");
            }
            else
            {
                return range.Value2.ToString();
            }
        }
        #endregion

        #region Object
        public static object[,] GetContentsAsObject2DArray(Range range)
        {
            object[,] rangeArray = new object[range.Rows.Count, range.Columns.Count];
            for (int rowNum = 0; rowNum < range.Rows.Count; rowNum++)
            {
                for (int colNum = 0; colNum < range.Columns.Count; colNum++)
                {
                    Range cell = range.Cells[rowNum+1, colNum+1];
                    rangeArray[rowNum, colNum] = cell.Value2;
                }
            }
            return rangeArray;
        }
        public static object[] GetContentsAsObject1DArray(Range range)
        {
            object[] rangeArray = new object[range.Cells.Count];
            int counter = 0;
            foreach (Range cell in range.Cells)
            {
                rangeArray[counter] = cell.Value2;
                counter++;
            }
            return rangeArray;
        }
        #endregion

        public static (int, int, int, int) GetRangeDetails(Range selectedRange)
        {
            //(int startRow, int endRow, int startCol, int endCol) = GetRangeDetails(selectedRange);
            int startRow = selectedRange.Row;
            int endRow = selectedRange.Row + selectedRange.Rows.Count - 1;
            int startCol = selectedRange.Column;
            int endCol = selectedRange.Column + selectedRange.Columns.Count - 1;
            return (startRow, endRow, startCol, endCol);
        }

        #endregion

        #region Check Excel Sheets
        public static (List<string>, List<string>) CheckIfSheetsExist(Workbook workbook, IEnumerable<string> sheetNames, bool throwError = false)
        {
            #region Check
            List<string> existingSheets = new List<string>();
            List<string> missingSheets = new List<string>();

            foreach (string sheetName in sheetNames)
            {
                try
                {
                    Worksheet worksheet = workbook.Sheets[sheetName];
                    existingSheets.Add(sheetName);
                }
                catch //(Exception ex)
                {
                    // Do nothing
                    missingSheets.Add(sheetName);
                }
            }
            #endregion

            #region Throw Error
            if (throwError)
            {
                if (missingSheets.Count > 0)
                {
                    string sheetString = "";
                    foreach (string sheet in missingSheets)
                    {
                        sheetString += sheet + "\n";
                    }
                    throw new Exception($"The following sheets do not exist:\n{sheetString}");
                }
            }
            #endregion

            return (existingSheets, missingSheets);
        }

        public static void AskToDeleteSheets(Workbook workbook, IEnumerable<string> sheetNames, string msg = "Delete the following sheets?")
        {
            string stringOfSheets = "";
            foreach (string sheetName in sheetNames)
            {
                stringOfSheets += sheetName + "\n";
            }

            DialogResult result = MessageBox.Show($"{msg}\n{stringOfSheets}", "Confirmation", MessageBoxButtons.YesNoCancel);
            if (result == DialogResult.Yes)
            {
                foreach (string sheetName in sheetNames)
                {
                    Worksheet worksheet = workbook.Sheets[sheetName];
                    worksheet.Delete();
                }
            }
            else if (result == DialogResult.Cancel)
            {
                throw new Exception("Terminated by user");
            }
        }
        #endregion

        #region Others
        public static bool Confirmation(string msg, bool throwException = true)
        {
            DialogResult res = MessageBox.Show(msg, "Confirmation", MessageBoxButtons.OKCancel);
            if (res != DialogResult.OK)
            {
                if (throwException)
                {
                    throw new Exception("Cancelled by user");
                }
                return false;
            }
            return true;
        }

        public static void ThrowExceptionBox(string msg)
        {
            Console.WriteLine($"{msg}");
            throw new Exception(msg);
        }

        public static double[] OffsetPoint(double[] startPoint, double offX, double offY, double offZ = 0)
        {
            double[] endPoint = (double[])startPoint.Clone();
            endPoint[0] += offX;
            endPoint[1] += offY;
            endPoint[2] += offZ;
            return endPoint;
        }

        public static (int, int, int) DecimalToRGB(double decimalColor)
        {
            //decimalColor = B * 65536 + G * 256 + R

            double B = Math.Floor(decimalColor / 65536);
            double G = Math.Floor((decimalColor - B * 65536) / 256);
            double R = Math.Floor(decimalColor - B * 65536 - G * 256);

            return ((int)R, (int)G, (int)B);
        }

        public static double RGBToDecimal(int R, int G, int B)
        {
            double decimalColor = B * 65536 + G * 256 + R;
            return decimalColor;
        }

        public static int ConvertToProgress(int currentProgress, int maxProgress)
        {
            double progressDouble = Convert.ToDouble(currentProgress) / Convert.ToDouble(maxProgress) * 100;
            int progress = Convert.ToInt32(progressDouble);
            return progress;
        }

        public static string ConvertToString(IEnumerable<string> items, string delimitor = "\n")
        {
            string finalString = "";
            foreach (string item in items)
            {
                finalString += item + delimitor;
            }
            return finalString;
        }

        public static string[] SplitAndTrim(string inputString, char deliminator = ',')
        {
            string[] parts = inputString.Split(deliminator);
            for (int i = 0; i < parts.Length; i++)
            {
                parts[i] = parts[i].Trim();
            }
            return parts;
        }
        #endregion

        #region File Operations
        public static bool DeleteFile(string path)
        {
            // Returns true if file is successfully deleted
            // Returns false if file is not successfully deleted
            try
            {
                File.Delete(path);
                return true;
            }
            catch
            {
                DialogResult result = MessageBox.Show("Unable to delete file.\nFile may be opened by user.\nTry again?", "Error", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    return DeleteFile(path);
                }
                else
                {
                    return false;
                }
            }
        }
        public static bool CheckAndDeleteFile(string filePath)
        {
            // If file doesn't exist return true
            // If file exist and is deleted return true
            // If file exist and is not deleted return false

            if (!File.Exists(filePath))
            {
                return true;
            }

            if (DialogResult.Yes == MessageBox.Show($"File already exist at following path, delete file and proceed?\n\n{filePath}", "Error", MessageBoxButtons.YesNo))
            {
                return DeleteFile(filePath);
            }
            else
            {
                return false;
            }
        }

        public static void OpenFiles(string[] inputPaths)
        {
            if (inputPaths.Length > 5)
            {
                if (DialogResult.No == MessageBox.Show($"{inputPaths.Length} files detected, continue to open all files?", "Warning", MessageBoxButtons.YesNo))
                {
                    return;
                }
            }

            foreach (string inputPath in inputPaths)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(inputPath) { UseShellExecute = true });
            }
        }

        public static string MergeFileNameAndDir(string dir, string fileName, string extension = "")
        {
            if (dir == "")
            {
                throw new Exception($"Directory cannot be empty");
            }
            else if (fileName == "")
            {
                throw new Exception($"File Name cannot be empty");
            }

            string finalFileName = fileName;
            string check = Path.GetExtension(finalFileName);
            if (extension != "")
            {
                if (Path.GetExtension(fileName) != extension)
                {
                    finalFileName += extension;
                }
            }

            return Path.Combine(dir, finalFileName);
        }
        public static string SanitiseFileName(string inputFileName)
        {
            HashSet<char> invalidChars = new HashSet<char>(Path.GetInvalidFileNameChars());

            StringBuilder sanitizedFileName = new StringBuilder();
            foreach (char c in inputFileName)
            {
                if (!invalidChars.Contains(c))
                {
                    sanitizedFileName.Append(c);
                }
            }

            return sanitizedFileName.ToString();
        }
        #endregion

        #region Check Excel Data
        public static bool CheckRangeFileExist(Range checkRange, bool showError = false, bool ignoreEmpty = false)
        {
            List<Range> errorRange = new List<Range>();

            foreach (Range cell in checkRange.Cells)
            {
                if (cell.Value2 != null)
                {
                    string cellValue = cell.Value2.ToString();
                    if (!File.Exists(cellValue))
                    {
                        if (!showError)
                        {
                            // Terminate upon first occurrence
                            return false;
                        }
                        errorRange.Add(cell);
                    }
                }
                else if (ignoreEmpty)
                {
                    continue;
                }
                else
                {
                    if (!showError)
                    {
                        // Terminate upon first occurrence
                        return false;
                    }
                    errorRange.Add(cell);
                }
            }

            if (errorRange.Count > 0)
            {
                if (showError)
                {
                    string msg = $"The following inputs are not valid filepaths:\n\n";
                    foreach (Range cell in errorRange)
                    {
                        string cellDetails;
                        if (cell.Value2 != null)
                        {
                            cellDetails = $"Cell Address: {cell.Address[false, false]}\n{cell.Value2.ToString()}\n";
                        }
                        else
                        {
                            cellDetails = $"Cell Address: {cell.Address[false, false]}\nEmpty Cell\n";
                        }
                        msg += cellDetails;
                    }
                    MessageBox.Show(msg, "Error");
                }
                return false;
            }
            return true;
        }

        public static bool CheckRangeIsFilled(Range checkRange, bool showError = false)
        {
            List<Range> errorRange = new List<Range>();

            foreach (Range cell in checkRange.Cells)
            {
                if (cell.Value2 == null)
                {
                    errorRange.Add(cell);
                }
                else if (cell.Value2 == "")
                {
                    errorRange.Add(cell);
                }
            }

            if (errorRange.Count > 0)
            {
                if (showError)
                {
                    string msg = $"The following cells are empty:\n\n";
                    foreach (Range cell in errorRange)
                    {
                        msg += $"{cell.Address[false, false]}\n";
                    }
                    MessageBox.Show(msg, "Error");
                }
                return false;
            }
            return true;
        }

        public static void CheckRangeSize(Range selectedRange, int numRows, int numCols, string attName = "", bool ignoreLargerSize = false)
        {
            // Checks that range meets the desired size. Set numRows/numCols = 0 to skip check
            #region Check Rows
            if (numRows == 0) { } // skip check
            else if (ignoreLargerSize && numRows > selectedRange.Rows.Count)
            {
                string msg = $"Number of rows expected is at least {numRows}\nNumber of rows selected is {selectedRange.Rows.Count}";
                if (attName != "")
                {
                    msg = $"Attribute Name: {attName}\n" + msg;
                }
                throw new Exception(msg);
            }
            else if (!ignoreLargerSize && numRows != selectedRange.Rows.Count)
            {
                string msg = $"Number of rows expected is {numRows}\nNumber of rows selected is {selectedRange.Rows.Count}";
                if (attName != "")
                {
                    msg = $"Attribute Name: {attName}\n" + msg;
                }
                throw new Exception(msg);
            }
            #endregion

            #region Check Columns
            if (numCols == 0) { } // skip check
            else if (ignoreLargerSize && numCols > selectedRange.Columns.Count)
            {
                string msg = $"Number of columns expected is at least {numCols}\nNumber of columns selected is {selectedRange.Columns.Count}";
                if (attName != "")
                {
                    msg = $"Attribute Name: {attName}\n" + msg;
                }
                throw new Exception(msg);
            }
            else if (!ignoreLargerSize && numCols != selectedRange.Columns.Count)
            {
                string msg = $"Number of columns expected is {numCols}\nNumber of columns selected is {selectedRange.Columns.Count}";
                if (attName != "")
                {
                    msg = $"Attribute Name: {attName}\n" + msg;
                }
                throw new Exception(msg);
            }
            #endregion
        }
        public static (bool passCheck, List<Range> failedRanges) AssertStandardRangeSize(Range[] ranges, string type = null, bool throwError = true)
        {
            //Checks that all ranges have the same size
            #region Generate check type
            bool checkRow;
            bool checkCol;

            if (type == null)
            {
                checkRow = true;
                checkCol = true;
            }
            else if (type == "column")
            {
                checkRow = false;
                checkCol = true;
            }
            else if (type == "row")
            {
                checkRow = true;
                checkCol = false;
            }
            else
            {
                throw new ArgumentException($"AssertRangeSize, input type {type} not found.");
            }
            #endregion
            int numRows = ranges[0].Rows.Count;
            int numCols = ranges[0].Columns.Count;

            List<int> failedRangeNums = new List<int>();

            for (int i = 1; i < ranges.Length; i++)
            {
                Range range = ranges[i];
                bool failedCheck = false;

                if (checkRow && range.Rows.Count != numRows)
                {
                    failedCheck = true;
                }

                if (checkCol && range.Columns.Count != numCols)
                {
                    failedCheck = true;
                }

                if (failedCheck)
                {
                    failedRangeNums.Add(i);
                }
            }

            if (failedRangeNums.Count == 0) { return (true, new List<Range>()); }

            #region Create range list
            List<Range> failedRanges = new List<Range>();
            foreach (int rangeNum in failedRangeNums)
            {
                failedRanges.Add(ranges[rangeNum]);
            }
            #endregion

            if (!throwError) { return (false, failedRanges); }

            string msg = $"Not all ranges have the same size. The following ranges do not match range 1 {ranges[0].Address[false, false]} with size [{ranges[0].Rows.Count},{ranges[0].Columns.Count}]:\n";
            foreach (Range failedRange in failedRanges)
            {
                msg += $"Range at {failedRange.Address[false, false]}: " +
                    $"size is [{failedRange.Rows.Count}, {failedRange.Columns.Count}]\n";
            }
            throw new ArgumentException(msg);
        }
        public static void IntersectRanges(ref Range range1, ref Range range2, string type = null)
        {
            if (type == null)
            {
                int minRows = Math.Min(range1.Rows.Count, range2.Rows.Count);
                int minCols = Math.Min(range1.Columns.Count, range2.Columns.Count);
                range1 = range1.Resize[minRows, minCols];
                range2 = range2.Resize[minRows, minCols];
            }
            else if (type == "column")
            {
                int minCols = Math.Min(range1.Columns.Count, range2.Columns.Count);
                range1 = range1.Resize[range1.Rows.Count, minCols];
                range2 = range2.Resize[range2.Rows.Count, minCols];
            }
            else if (type == "row")
            {
                int minRows = Math.Min(range1.Rows.Count, range2.Rows.Count);
                range1 = range1.Resize[minRows, range1.Columns.Count];
                range2 = range2.Resize[minRows, range2.Columns.Count];
            }
            else
            {
                throw new ArgumentException($"IntersectRanges, input type {type} not found.");
            }
        }
        #endregion

        #region Check Input
        //public static Range CheckStringIsRange(string address, bool withSheet)
        //{
        //    if (address == "")
        //    {
        //        throw new ArgumentNullException($"Empty input");
        //    }

        //    if (!withSheet)
        //    {
        //        string sheetName = Globals.ThisAddIn.Application.ActiveSheet.Name;
        //        address = sheetName + "!" + address;
        //    }
        //    var parts = address.Split('!');
        //    if (parts.Length != 2)
        //    {
        //        throw new ArgumentException($"Invalid address format for address: {address}. Expected format: SheetName!CellAddress");
        //    }

        //    try
        //    {
        //        Worksheet ThisWorksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[parts[0]];
        //        Range returnRange = ThisWorksheet.Range[parts[1]];
        //        return returnRange;
        //    }
        //    catch
        //    {
        //        throw new ArgumentException($"Error Returning Range at {address}");
        //    }
        //}

        #endregion

        #region Format Excel
        //public static void FormatCellEqualsTo(Range range, string matchString, Color fontColor)
        //{

        //}

        public static void FormatCellEqualsTo(Range range, string[] rangeContents, string matchString, Color fontColor)
        {
            Range formatRange = null;

            foreach (Range cell in range)
            {
                string cellValue = cell.Value2?.ToString().Trim();
                if (cellValue == matchString)
                {
                    if (formatRange == null)
                    {
                        formatRange = cell;
                    }
                    else
                    {
                        formatRange = range.Application.Union(formatRange, cell);
                    }
                }
            }

            if (formatRange != null)
            {
                formatRange.Font.Color = fontColor;
            }
        }
        #endregion

        #region Manipulate Excel Ranges
        public static void TerminateRangeAtNullFirstCell(ref Range range, int checkCol = 0)
        {
            // checkCol determinates which column will be checked to be null
            // checkCol == 0 means all columns will be checked

            (int startRow, int endRow, int startCol, int endCol) = GetRangeDetails(range);
            Worksheet ws = range.Worksheet;
            int finalRowNum = -1;
            for (int rowNum = startRow; rowNum <= endRow; rowNum++)
            {
                Range thisStartCell = ws.Cells[rowNum, startCol];
                Range thisEndCell = ws.Cells[rowNum, endCol];
                Range row = ws.Range[thisStartCell, thisEndCell];

                // Check if the cell value is null or empty
                if (checkForNull(row))
                {
                    finalRowNum = rowNum - 1;
                    break;
                }
            }

            if (finalRowNum == -1)
            {
                return;
            }
            else
            {
                Range startCell = ws.Cells[startRow, startCol];
                Range endCell = ws.Cells[finalRowNum, endCol];
                Range returnRange = ws.Range[startCell, endCell];
                range = returnRange;
            }

            bool checkForNull(Range row)
            {
                if (checkCol == 0)
                {
                    foreach (Range cell in row)
                    {
                        if (cell.Value2 == null)
                        {
                            return true;
                        }
                    }
                    return false;
                }
                else
                {
                    Range cell = row.Columns[checkCol];
                    if (cell.Value2 == null)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
        }
        public static void TerminateRangeAtFirstNullRow(ref Range range)
        {
            for (int rowNum = 1; rowNum <= range.Rows.Count; rowNum++)
            {
                bool isAllNull = true;
                foreach (Range cell in range.Rows[rowNum].Cells)
                {
                    if (cell.Value2 != null)
                    {
                        isAllNull = false;
                        break;
                    }
                }

                if (isAllNull)
                {
                    range = range.Resize[rowNum - 1, range.Columns.Count];
                    break;
                }
            }
        }

        //public static Range GetColRangeFromRanges(Range columnCell, Range rowRange, int offsetRow = 0, int offsetCol = 0)
        public static Range GetColRangeFromRanges(Range rowRange, Range columnCell, int offsetRow = 0, int offsetCol = 0)
        {
            // Returns a range with column number equals to column range, but row(s) equal to rowRange
            (int startRowNum, int endRowNum, _, _) = GetRangeDetails(rowRange);
            int colNum = columnCell.Column;
            Worksheet worksheet = columnCell.Worksheet;
            Range startCellL = worksheet.Cells[startRowNum, colNum];
            Range endCellL = worksheet.Cells[endRowNum, colNum];
            Range returnRange = worksheet.Range[startCellL, endCellL];

            if (offsetRow != 0 || offsetCol != 0)
            {
                returnRange = returnRange.Offset[offsetRow, offsetCol];
            }

            return returnRange;
        }
        public static Range GetLastCellFromEnd(Worksheet worksheet, int colOrRowNum = 1, XlDirection xlDirection = XlDirection.xlUp)
        {
            if (colOrRowNum < 1) { throw new ArgumentException($"Column or row number cannot be < 1. Col num = 1 refers to column A. Number provided = {colOrRowNum}."); }
            Range lastCell = null;
            if (xlDirection == XlDirection.xlUp)
            {
                lastCell = worksheet.Cells[1048576, colOrRowNum];
            }
            else if (xlDirection == XlDirection.xlToLeft)
            {
                lastCell = worksheet.Cells[colOrRowNum, 16384];
            }
            else { throw new Exception($"Invalid xlDirection"); }

            Range lastUsedCell = lastCell.End[xlDirection];
            if (lastUsedCell.MergeCells)
            {
                Range mergedArea = lastUsedCell.MergeArea;
                if (xlDirection == XlDirection.xlUp)
                {
                    lastUsedCell = mergedArea.Rows[mergedArea.Rows.Count];
                }
                else if (xlDirection == XlDirection.xlToLeft)
                {
                    lastUsedCell = mergedArea.Columns[mergedArea.Columns.Count];
                }
                else { throw new Exception($"Invalid xlDirection"); }
            }
            return lastUsedCell;
        }

        public static Range GetLastCellFromStartCell(Worksheet worksheet, int rowNum = 1, int colNum = 1)
        {
            if (colNum < 1) { throw new ArgumentException($"Column or row number cannot be < 1. Col num = 1 refers to column A. Number provided = {colNum}."); }
            if (rowNum < 1) { throw new ArgumentException($"Column or row number cannot be < 1. Row num = 1 refers to row 1. Number provided = {rowNum}."); }

            Range startCell = worksheet.Cells[rowNum, colNum];
            Range lastVerticalCell = startCell.End[XlDirection.xlDown];
            lastVerticalCell.Select();
            Range lastHorizontalCell = startCell.End[XlDirection.xlToRight];
            Range lastUsedCell = worksheet.Cells[lastVerticalCell.Row, lastHorizontalCell.Column];

            if (lastUsedCell.MergeCells)
            {
                Range mergedArea = lastUsedCell.MergeArea;
                lastUsedCell = mergedArea.Cells[mergedArea.Rows.Count - 1, mergedArea.Columns.Count - 1];
            }
            return lastUsedCell;
        }
        #endregion

        #region Write to Excel
        public static void WriteToExcelSelectionAsRow(Application excelApp, int rowOff, int colOff, bool warning, params Array[] arrays)
        {
            // This code takes any number of arrays (of various types) and outputs them into excel 
            // Output order depends on order of the input array
            // Output location is the first cell of the current selection, offset by rowOff and colOff

            // Find number of rows and columns
            int numRow = 0;
            int numCol = arrays.Length;
            for (int col = 0; col < arrays.Length; col++)
            {
                if (arrays[col].Length > numRow)
                {
                    numRow = arrays[col].Length; // Finds max number of rows out of all the various arrays
                }
            }
            #region Get Confirmation
            if (warning)
            {
                DialogResult result = MessageBox.Show("Confirm to export values to current selection? This will override cell values at current selection and cannot be undone.\n" +
                "Output table size:\n" +
                $"Number of rows: {numRow}\n" +
                $"Number of columns: {numCol}", "Confirmation", MessageBoxButtons.YesNo);
                if (result != DialogResult.Yes)
                {
                    throw new Exception("Terminated by user");
                }
            }
            #endregion

            // Initiate object
            object[,] dataArray = new object[numRow, numCol];
            for (int col = 0; col < arrays.Length; col++)
            {
                for (int row = 0; row < arrays[col].Length; row++)
                {
                    dataArray[row, col] = arrays[col].GetValue(row);
                }
            }

            // Add section to read input data from Excel
            Workbook objBook = excelApp.ActiveWorkbook;
            Worksheet objSheet = objBook.ActiveSheet;
            Range selectedRange = excelApp.ActiveWindow.RangeSelection;

            try
            {
                // Write to Excel
                objBook.Application.ScreenUpdating = false;
                // Write the entire array to the worksheet in one go using Value2
                Range startCell = objSheet.Cells[selectedRange.Row + rowOff, selectedRange.Column + colOff];
                Range endCell = startCell.Offset[numRow - 1, numCol - 1];
                Range writeRange = objSheet.Range[startCell, endCell];
                writeRange.Value2 = dataArray;
            }
            finally
            {
                objBook.Application.ScreenUpdating = true;
            }
        }

        public static Range WriteToExcelRangeAsRow(Application excelApp, Range startRange, int rowOff, int colOff, bool warning, params Array[] arrays)
        {
            // This code takes any number of arrays (of various types) and outputs them into excel 
            // Output order depends on order of the input array
            // Output location is the first cell of the provided range, offset by rowOff and colOff

            // Find number of rows and columns
            if (startRange == null) { startRange = excelApp.ActiveWindow.RangeSelection; }
            int numRow = arrays.Length;
            int numCol = 0;
            foreach (Array array in arrays)
            {
                if (array.Length > numCol) { numCol = array.Length; }
            }

            #region Get Confirmation
            if (warning)
            {
                DialogResult result = MessageBox.Show("Confirm to export values to current selection? This will override cell values at current selection and cannot be undone.\n" +
                "Output table size:\n" +
                $"Number of rows: {numRow}\n" +
                $"Number of columns: {numCol}", "Confirmation", MessageBoxButtons.YesNo);
                if (result != DialogResult.Yes)
                {
                    throw new Exception("Terminated by user");
                }
            }
            #endregion

            #region Create Data Array
            object[,] dataArray = new object[numRow, numCol];
            for (int row = 0; row < numRow; row++)
            {
                for (int col = 0; col < numCol; col++)
                {
                    dataArray[row, col] = arrays[row].GetValue(col);
                }
            }
            #endregion

            #region Write to Excel
            // Add section to read input data from Excel
            Workbook workBook = excelApp.ActiveWorkbook;
            Worksheet workSheet = startRange.Worksheet;

            Range writeRange = null;
            try
            {
                workBook.Application.ScreenUpdating = false;
                Range startCell = startRange.Cells[1, 1];
                startCell = startCell.Offset[rowOff, colOff];
                Range endCell = startCell.Offset[numRow - 1, numCol - 1];
                writeRange = workSheet.Range[startCell, endCell];
                writeRange.Value2 = dataArray;
            }
            finally
            {
                workBook.Application.ScreenUpdating = true;
            }
            return writeRange;
            #endregion
        }

        public static Range WriteToExcelRangeAsCol(Application excelApp, Range startRange, int rowOff, int colOff, bool warning, params Array[] arrays)
        {
            // This code takes any number of arrays (of various types) and outputs them into excel 
            // Output order depends on order of the input array
            // Output location is the first cell of the provided range, offset by rowOff and colOff

            // Find number of rows and columns
            if (startRange == null) { startRange = excelApp.ActiveWindow.RangeSelection; }
            int numCol = arrays.Length;
            int numRow = 0;

            foreach (Array array in arrays)
            {
                if (array.Length > numRow) { numRow = array.Length; }
            }

            #region Get Confirmation
            if (warning)
            {
                DialogResult result = MessageBox.Show("Confirm to export values to current selection? This will override cell values at current selection and cannot be undone.\n" +
                "Output table size:\n" +
                $"Number of rows: {numRow}\n" +
                $"Number of columns: {numCol}", "Confirmation", MessageBoxButtons.YesNo);
                if (result != DialogResult.Yes)
                {
                    throw new Exception("Terminated by user");
                }
            }
            #endregion

            #region Create Data Array
            object[,] dataArray = new object[numRow, numCol];
            for (int row = 0; row < numRow; row++)
            {
                for (int col = 0; col < numCol; col++)
                {
                    dataArray[row, col] = arrays[col].GetValue(row);
                }
            }
            #endregion

            #region Write to Excel
            // Add section to read input data from Excel
            Workbook workBook = excelApp.ActiveWorkbook;
            Worksheet workSheet = startRange.Worksheet;
            Range writeRange = null;
            try
            {
                workBook.Application.ScreenUpdating = false;
                Range startCell = startRange.Cells[1, 1];
                startCell = startCell.Offset[rowOff, colOff];
                Range endCell = startCell.Offset[numRow - 1, numCol - 1];
                writeRange = workSheet.Range[startCell, endCell];
                writeRange.Value2 = dataArray;
            }
            finally
            {
                workBook.Application.ScreenUpdating = true;
            }
            return writeRange;
            #endregion
        }

        public static Range WriteObjectToExcelRange(Application excelApp, Range startRange, int rowOff, int colOff, bool warning, object[,] writeObject)
        {
            int numRow = writeObject.GetLength(0);
            int numCol = writeObject.GetLength(1);

            if (startRange == null) { startRange = excelApp.ActiveCell; }
            Worksheet workSheet = startRange.Worksheet;
            Range writeRange = null;
            try
            {
                excelApp.ScreenUpdating = false;
                Range startCell = startRange.Cells[1, 1];
                startCell = startCell.Offset[rowOff, colOff];
                Range endCell = startCell.Offset[numRow - 1, numCol - 1];
                writeRange = workSheet.Range[startCell, endCell];
                writeRange.Value2 = writeObject;
            }
            finally
            {
                excelApp.ScreenUpdating = true;
            }
            return writeRange;
        }

        public static Range WriteObjectToExcelRange(Application excelApp, Range startRange, int rowOff, int colOff, bool warning, double[,] writeObject)
        {
            // Quick copy and paste, come back to refractor with method above^
            int numRow = writeObject.GetLength(0);
            int numCol = writeObject.GetLength(1);

            if (startRange == null) { startRange = excelApp.ActiveCell; }
            Worksheet workSheet = startRange.Worksheet;
            Range writeRange = null;
            try
            {
                excelApp.ScreenUpdating = false;
                Range startCell = startRange.Cells[1, 1];
                startCell = startCell.Offset[rowOff, colOff];
                Range endCell = startCell.Offset[numRow - 1, numCol - 1];
                writeRange = workSheet.Range[startCell, endCell];
                writeRange.Value2 = writeObject;
            }
            finally
            {
                excelApp.ScreenUpdating = true;
            }
            return writeRange;
        }

        public static void ClearRangeForPrintingObject(Range startRange, int rowOff, int colOff, object[,] writeObject)
        {

            Worksheet workSheet = startRange.Worksheet;
            Range startCell = startRange.Cells[1, 1];
            startCell = startCell.Offset[rowOff, colOff];

            int numRow = writeObject.GetLength(0);
            int numCol = writeObject.GetLength(1);
            Range endCell = startCell.Offset[numRow - 1, numCol - 1];
            Range writeRange = workSheet.Range[startCell, endCell];
            MessageBox.Show($"startRange = {startRange.Address}");
            MessageBox.Show($"writeRange = {writeRange.Address}");
            writeRange.Select();

            writeRange.UnMerge();
            if (writeRange.MergeCells) { writeRange.UnMerge(); }
            writeRange.ClearContents();
        }
        #endregion

        #region Worksheet
        public static Worksheet CopyNewSheetAtBack(Worksheet refSheet, string newName = "", bool deleteExisting = false)
        {
            Workbook thisWorkbook = refSheet.Parent;
            Worksheet newSheet;
            // Check if sheet exist
            if (newName != "")
            {
                foreach (Worksheet sheet in thisWorkbook.Sheets)
                {
                    if (sheet.Name == newName)
                    {
                        if (deleteExisting)
                        {
                            try
                            {
                                refSheet.Application.DisplayAlerts = false;
                                sheet.Delete();
                            }
                            finally
                            {
                                refSheet.Application.DisplayAlerts = true;
                            }

                        }
                        else
                        {
                            throw new Exception($"Worksheet already exist.\nWorksheet Name:{newName}");
                        }

                    }
                }
            }

            // Copy sheet
            try
            {
                refSheet.Copy(Before: refSheet);
                newSheet = thisWorkbook.Sheets[refSheet.Index - 1];
                newSheet.Move(After: thisWorkbook.Sheets[thisWorkbook.Sheets.Count]);
            }
            catch (Exception ex)
            {
                throw new Exception($"Unable to copy worksheet {refSheet.Name}\n\n" + ex.Message);
            }

            // Rename sheet (should be no error)
            try
            {
                if (newName != "")
                {
                    newSheet.Name = newName;
                }
                return newSheet;
            }
            catch (Exception ex)
            {
                throw new Exception($"Unable to rename worksheet {newSheet} to worksheet with new name {newName}\n\n" + ex.Message);
            }
        }
        #endregion

        #region Array Functions
        public static string[] ConcatArrays(List<Array> ArraysToWrite)
        {
            // Turns list of array into a single array (append)
            List<string> listToWrite = new List<string>();
            foreach (Array array in ArraysToWrite)
            {
                foreach (object obj in array)
                {
                    if (obj == null) { listToWrite.Add(null); }
                    else { listToWrite.Add(obj.ToString()); }

                }
            }
            return listToWrite.ToArray();
        }



        #endregion


        #region Get Excel Workbook
        public static Workbook OpenAndGetWorkbook(Application app, string path)
        {
            string extension = Path.GetExtension(path);
            if (!(extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsb", StringComparison.OrdinalIgnoreCase)))
            {
                throw new Exception($"File type is not valid excel file.\n{path}");
            }
            string workbookName = Path.GetFileName(path);

            Workbook workbook;
            try
            {
                workbook = app.Workbooks[workbookName];
            }
            catch
            {
                workbook = app.Workbooks.Open(path, ReadOnly: true);
                return workbook;
            }

            string workbookPath = workbook.FullName;
            if (workbookPath == path) { return workbook; }
            else { throw new Exception($"File with similar workbook name {workbook.Name} is already open. Please close this workbook before proceeding."); }
        }
        #endregion

        #region Headers 

        #region Insert Headers At Selection
        //public static void InsertListHeadersAtSelection(List<string> headers, string type = "cols", bool format = true)
        //{
        //    Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
        //    Range writeRange = null;
        //    if (type == "cols")
        //    {
        //        writeRange = WriteToExcelRangeAsRow(selectedRange, 0, 0, true, headers.ToArray());
        //    }
        //    else if (type == "rows")
        //    {
        //        writeRange = WriteToExcelRangeAsCol(selectedRange, 0, 0, true, headers.ToArray());
        //    }
        //    else
        //    {
        //        throw new Exception($"Input type \"{type}\" is invalid");
        //    }

        //    if (format && writeRange != null)
        //    {
        //        writeRange.Font.Bold = true;
        //        writeRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        //        writeRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
        //    }
        //}

        //public static void InsertArrayHeadersAtSelection(string[,] headers, bool format = true)
        //{
        //    Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
        //    Range writeRange = WriteObjectToExcelRange(selectedRange, 0, 0, true, headers);

        //    if (format && writeRange != null)
        //    {
        //        writeRange.Font.Bold = true;
        //        writeRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        //        writeRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
        //    }
        //}

        //public static void InsertHeadersAtSelection(object headers, bool format = true)
        //{
        //    // this cannot be combined because only list has the option to add header as column :(
        //    if (headers is List<string> listHeaders)
        //    {
        //        InsertListHeadersAtSelection(listHeaders);
        //    }
        //    else if (headers is string[,] arrayHeaders)
        //    {
        //        InsertArrayHeadersAtSelection(arrayHeaders);
        //    }
        //    else
        //    {
        //        throw new Exception("Unsupported header object type");
        //    }
        //}

        //public static void InsertDynamicHeader(Func<object> headerGenerator, bool format = true)
        //{
        //    try
        //    {
        //        var header = headerGenerator();
        //        InsertHeadersAtSelection(header, format);
        //    }
        //    catch (Exception ex) { MessageBox.Show($"Unable to generate header\n{ex.Message}","Error"); }
        //}
        #endregion

        #region Subscribe Button
        //public static void AddHeaderMenuToButton(Button button, List<string> headerText, string toolStripName = "Add Header", bool format = true, string headerOrientation = "cols")
        //{
        //    if (button.ContextMenuStrip == null) { button.ContextMenuStrip = new ContextMenuStrip(); }
        //    ToolStripMenuItem headerItem = new ToolStripMenuItem(toolStripName);
        //    button.ContextMenuStrip.Items.Add(headerItem);
        //    headerItem.Click += (sender, e) => InsertListHeadersAtSelection(headerText, headerOrientation, format);
        //}

        //public static void AddHeaderMenuToButton(Button button, string[,] headerText, string toolStripName = "Add Header", bool format = true)
        //{
        //    if (button.ContextMenuStrip == null) { button.ContextMenuStrip = new ContextMenuStrip(); }
        //    ToolStripMenuItem headerItem = new ToolStripMenuItem(toolStripName);
        //    button.ContextMenuStrip.Items.Add(headerItem);
        //    headerItem.Click += (sender, e) => InsertArrayHeadersAtSelection(headerText, format);
        //}

        //public static void AddDynamicActionToButton(Button button, Action action, string toolStripName = "Add Header")
        //{
        //    if (button.ContextMenuStrip == null) { button.ContextMenuStrip = new ContextMenuStrip(); }
        //    ToolStripMenuItem headerItem = new ToolStripMenuItem(toolStripName);
        //    button.ContextMenuStrip.Items.Add(headerItem);
        //    headerItem.Click += (sender, e) => action();
        //}
        #endregion

        #region Combo Headers
        //public static void AddComboHeaderMenuToButton(Button button, Dictionary<string, object> headerDic, ComboBox comboBox, string toolStripName = "Add Header", bool format = true, string headerOrientation = "cols")
        //{
        //    #region Usage Example
        //    //#region Dynamic Headers for Get Coordinates from AutoCAD Pane
        //    //Dictionary<string, object> headerDic = new Dictionary<string, object>();
        //    //string[,] arrayHeaders = new string[,] { 
        //    //    { "Start", "", "", "Mid", "", "", "End", "", "" }, 
        //    //    { "X", "Y", "Z", "X", "Y", "Z", "X", "Y", "Z" } };
        //    //headerDic.Add("1 Start, mid, end", arrayHeaders);

        //    //arrayHeaders = new string[,] {
        //    //    { "Start", "", "", "End", "", "" },
        //    //    { "X", "Y", "Z", "X", "Y", "Z" } };
        //    //headerDic.Add("2 Start, end", arrayHeaders);

        //    //arrayHeaders = new string[,] {
        //    //    { "Mid", "", ""},
        //    //    { "X", "Y", "Z"} };
        //    //headerDic.Add("3 Mid", arrayHeaders);

        //    //AddComboHeaderMenuToButton(getLineCoords, headerDic, dispLineOptions);
        //    //#endregion
        //    #endregion
        //    if (button.ContextMenuStrip == null) { button.ContextMenuStrip = new ContextMenuStrip(); }
        //    ToolStripMenuItem headerItem = new ToolStripMenuItem(toolStripName);
        //    button.ContextMenuStrip.Items.Add(headerItem);
        //    headerItem.Click += (sender, e) => InsertComboHeaderAtSelection(headerDic, comboBox, headerOrientation, format);
        //}

        //public static void InsertComboHeaderAtSelection(Dictionary<string, object> headerDic, object attObject, string headerOrientation, bool format)
        //{
        //    try
        //    {
        //        // Get parameter value
        //        string value;
        //        if (attObject is ComboBox comboBox)
        //        {
        //            value = comboBox.Text;
        //        }
        //        else if (attObject is string stringObj)
        //        {
        //            value = stringObj;
        //        }
        //        else { throw new Exception($"Incorrect type for header dynamic attObject"); }

        //        if (!headerDic.ContainsKey(value)) { throw new Exception($"No matching header for value \"{value}\""); }

        //        // Print Value
        //        if (headerDic[value] is List<string> listHeaders)
        //        {
        //            InsertListHeadersAtSelection(listHeaders);
        //        }
        //        else if (headerDic[value] is string[,] arrayHeaders)
        //        {
        //            InsertArrayHeadersAtSelection(arrayHeaders);
        //        }
        //        else
        //        {
        //            throw new Exception("Unsupported header type");
        //        }
        //    }
        //    catch (Exception ex) { MessageBox.Show($"Unable to add header, {ex.Message}", "Error"); }
        //}

        #endregion
        #endregion

        #region Get Directories
        public static void getFiles(string directory, ref List<string> globalFileList, bool checkNest = true, string extensionType = "")
        {
            // Get all directories and files within the current directory
            string[] subDirectoryList = Directory.GetDirectories(directory);
            string[] fileList = Directory.GetFiles(directory);

            //Add current files to global list
            if (extensionType == "") { globalFileList.AddRange(fileList); }
            else
            {
                HashSet<string> validExtensions = CreateExtensionHashset(extensionType);
                foreach (string file in fileList)
                {
                    if (validExtensions.Contains(Path.GetExtension(file)))
                    {
                        globalFileList.Add(file);
                    }
                }
            }

            // Recursively call this method to get files for each subdirectory
            if (!checkNest) { return; }
            
            foreach (string subDir in subDirectoryList)
            {
                getFiles(subDir, ref globalFileList, checkNest, extensionType);
            }
        }

        public static HashSet<string> CreateExtensionHashset(string allExtensions)
        {
            string[] parts = allExtensions.Split(',');
            HashSet<string> extensionHash = new HashSet<string>();
            for (int i = 0; i < parts.Length; i++)
            {
                string extensionCleaned = parts[i].Trim();
                if (extensionCleaned[0] != '.') { throw new Exception($"Invalid extension {extensionCleaned} provided"); }
                if (extensionCleaned == ".excel")
                {
                    HashSet<string> excelExtension = GetExcelExtensions();
                    foreach (string extension in excelExtension)
                    {
                        if (!extensionHash.Contains(extension)) { extensionHash.Add(extension); }
                    }
                }
                else
                {
                    if (!extensionHash.Contains(extensionCleaned)) { extensionHash.Add(extensionCleaned); }
                }
            }
            return extensionHash;
        }
        public static void getFolders(string directory, ref List<string> globalDirectoryList, bool checkNest = true)
        {
            // Get all directories and files within the specified directory
            string[] subDirectoryList = Directory.GetDirectories(directory);

            // Recursively call this method for each subdirectory
            if (checkNest)
            {
                foreach (string subDir in subDirectoryList)
                {
                    globalDirectoryList.Add(subDir);
                    getFolders(subDir, ref globalDirectoryList);
                }
            }
            else
            {
                globalDirectoryList.AddRange(subDirectoryList);
            }
        }

        //public static void getSpecificFiles(string directory, string extensionType, ref List<string> globalFileList, bool checkNest = true)
        //{
        //    // Get all directories and files within the specified directory
        //    string[] subDirectoryList = Directory.GetDirectories(directory);
        //    string[] fileList = Directory.GetFiles(directory);

        //    // Add directories and files to the global lists
        //    foreach (string file in fileList)
        //    {
        //        Path.GetExtension(file);
        //        if (Path.GetExtension(file) == extensionType)
        //        {
        //            globalFileList.Add(file);
        //        }
        //    }

        //    // Recursively call this method for each subdirectory
        //    if (checkNest)
        //    {
        //        foreach (string subDir in subDirectoryList)
        //        {
        //            getSpecificFiles(subDir, extensionType, ref globalFileList);
        //        }
        //    }
        //}
        
        public static HashSet<string> GetExcelExtensions()
        {
            HashSet<string> excelExtensions = new HashSet<string> 
            { 
                ".xlsx",
                ".xlsm",
                ".xlsb",
                ".xls",
                ".csv",
            };
            return excelExtensions;
        }

        public static void CreateDestinationFolder(string folderPath)
        {
            //Check if path exist
            if (!Directory.Exists(folderPath))
            {
                DialogResult result = MessageBox.Show("Folder does not currently exist. Create new folder?", "Error Opening Folder", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    catch
                    {
                        MessageBox.Show($"Unable to create folder at {folderPath}");
                        return;
                    }
                }
                else
                {
                    throw new Exception("Terminated by user");
                }
            }
        }
        #endregion

        #region Right Click Menu
        public static void AddContextStripEvent(System.Windows.Forms.Button button, string contextText, EventHandler eventHandler)
        {
            if (button.ContextMenuStrip == null) { button.ContextMenuStrip = new ContextMenuStrip(); }
            ToolStripMenuItem newItem = new ToolStripMenuItem(contextText);
            button.ContextMenuStrip.Items.Add(newItem);
            newItem.Click += eventHandler;
        }
        #endregion

        #region Convert Double to String
        public static string[] ConvertDoubleArrayToStringArray(double[] input, double n = 0)
        {
            string[] result = new string[input.Length];

            for (int i = 0; i < input.Length; i++)
            {
                double val = input[i];

                if (double.IsNaN(val))
                {
                    result[i] = "NA";
                }
                else
                {
                    double rounded = RoundToNearest(val, n);
                    result[i] = rounded.ToString();
                }
            }

            return result;
        }

        public static double RoundToNearest(double value, double n)
        {
            if (n == 0) return value; // avoid divide by zero

            double rounded = Math.Round(value / n) * n;
            // if n < 1, round to appropriate decimal places to avoid floating point artifact
            if (n < 1)
            {
                int decimals = (int)Math.Ceiling(-Math.Log10(n)) + 1;
                rounded = Math.Round(rounded, decimals);
            }
            return rounded;
        }
        public enum RoundingMode
        {
            Nearest,
            Up,
            Down
        }

        public static void ReplaceNaN(ref object[,] obj, string replaceValue = "NaN")
        {
            for (int rowNum = 0; rowNum < obj.GetLength(0); rowNum++)
            {
                for (int colNum = 0; colNum < obj.GetLength(1); colNum++)
                {
                    if (obj[rowNum, colNum] is double value && double.IsNaN(value))
                    {
                        obj[rowNum, colNum] = replaceValue;
                    }
                }
            }
        }
        #endregion

        #region Rounding
        public static double RoundDouble(double value, double n, RoundingMode mode = RoundingMode.Nearest)
        {
            if (n == 0) return value; // avoid divide by zero

            double rounded;

            switch (mode)
            {
                case RoundingMode.Nearest:
                    rounded = Math.Round(value / n) * n;
                    break;

                case RoundingMode.Up:
                    rounded = Math.Ceiling(value / n) * n;
                    break;

                case RoundingMode.Down:
                    rounded = Math.Floor(value / n) * n;
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(mode), "Unsupported rounding mode.");
            }

            // If n < 1, round to appropriate decimal places to avoid floating point artifact
            if (n < 1)
            {
                int decimals = (int)Math.Ceiling(-Math.Log10(n)) + 1;
                rounded = Math.Round(rounded, decimals);
            }

            return rounded;
        }

        #endregion
    }

    static class TwoDArrayFunctions
    {
        public enum OutputArrayType
        {
            Object,
            Double,
            String
        }
        public static (int totalRows, int totalCols) GetSizeOfConcat(List<object[,]> arrays)
        {
            // Get size of array if each array is just appended to the end
            // Total No. rows = sum (numRow)
            // Total No. cols = max (numCol)

            int totalRows = 0;
            int totalCols = 0;
            foreach (object[,] array in arrays)
            {
                totalRows += array.GetLength(0);
                if (totalCols < array.GetLength(1)) { totalCols = array.GetLength(1); }
            }

            return (totalRows, totalCols);
        }

        public static (int totalRows, int totalCols) GetSizeOfConcat(List<string[,]> arrays)
        {
            // Get size of array if each array is just appended to the end
            // Total No. rows = sum (numRow)
            // Total No. cols = max (numCol)

            int totalRows = 0;
            int totalCols = 0;
            foreach (string[,] array in arrays)
            {
                totalRows += array.GetLength(0);
                if (totalCols < array.GetLength(1)) { totalCols = array.GetLength(1); }
            }

            return (totalRows, totalCols);
        }

        public static (int totalRows, int totalCols) GetSizeOfConcat(List<double[,]> arrays)
        {
            // Get size of array if each array is just appended to the end
            // Total No. rows = sum (numRow)
            // Total No. cols = max (numCol)

            int totalRows = 0;
            int totalCols = 0;
            foreach (double[,] array in arrays)
            {
                totalRows += array.GetLength(0);
                if (totalCols < array.GetLength(1)) { totalCols = array.GetLength(1); }
            }

            return (totalRows, totalCols);
        }

        public static Array CreateEmptyArrayOfConcatSize(List<object[,]> arrays, OutputArrayType type, int addRows = 0, int addCols = 0)
        {
            // Returns empty array with size of concated array, plus additional rows and columns as required
            (int totalRows, int totalCols) = GetSizeOfConcat(arrays);

            switch (type)
            {
                case OutputArrayType.Double:
                    return new double[totalRows + addRows, totalCols + addCols];
                case OutputArrayType.String:
                    return new string[totalRows + addRows, totalCols + addCols];
                default:
                    return new object[totalRows + addRows, totalCols + addCols];
            }
        }

        public static Array CreateEmptyArrayOfConcatSize(List<double[,]> arrays, OutputArrayType type, int addRows = 0, int addCols = 0)
        {
            // Returns empty array with size of concated array, plus additional rows and columns as required
            (int totalRows, int totalCols) = GetSizeOfConcat(arrays);

            switch (type)
            {
                case OutputArrayType.Double:
                    return new double[totalRows + addRows, totalCols + addCols];
                case OutputArrayType.String:
                    return new string[totalRows + addRows, totalCols + addCols];
                default:
                    return new object[totalRows + addRows, totalCols + addCols];
            }
        }

        public static Array CreateEmptyArrayOfConcatSize(List<string[,]> arrays, OutputArrayType type, int addRows = 0, int addCols = 0)
        {
            // Returns empty array with size of concated array, plus additional rows and columns as required
            (int totalRows, int totalCols) = GetSizeOfConcat(arrays);

            switch (type)
            {
                case OutputArrayType.Double:
                    return new double[totalRows + addRows, totalCols + addCols];
                case OutputArrayType.String:
                    return new string[totalRows + addRows, totalCols + addCols];
                default:
                    return new object[totalRows + addRows, totalCols + addCols];
            }
        }

        public static double[] GetSingleRow(double[,] array, int rowNum)
        {
            int numCol = array.GetLength(1);
            double[] row = new double[numCol];
            for (int col = 0; col < numCol; col++)
            {
                row[col] = array[rowNum, col];
            }
            return row;
        }

        public static void WriteSingleRow(ref double[,] array, double[] row, int rowNum, bool checkNumCol = true)
        {
            int numCol = array.GetLength(1);

            if (checkNumCol)
            {
                if (row.Length != numCol)
                {
                    throw new ArgumentException("Row length does not match array column count.");
                }    
            }
            else
            {
                if (row.Length > numCol)
                {
                    throw new ArgumentException("Row length exceeds array column count.");
                }
            }

            for (int col = 0; col < numCol; col++)
            {
                array[rowNum, col] = row[col];
            }
        }
        public static void WriteArrayIntoArray(ref object[,] destArray, object[,] sourceArray, int insertRowNum, int insertColNum)
        {
            int sourceNumRows = sourceArray.GetLength(0);
            int sourceNumCols = sourceArray.GetLength(1);

            // Bounds check
            if (insertRowNum + sourceNumRows > destArray.GetLength(0) ||
                insertColNum + sourceNumCols > destArray.GetLength(1))
            {
                throw new ArgumentException("Source array does not fit into destination array at the given position.");
            }

            // Copy data
            for (int rowNum = 0; rowNum < sourceNumRows; rowNum++)
            {
                for (int colNum = 0; colNum < sourceNumCols; colNum++)
                {
                    destArray[insertRowNum + rowNum, insertColNum + colNum] = sourceArray[rowNum, colNum];
                }
            }
        }

        public static void WriteArrayIntoArray(ref double[,] destArray, double[,] sourceArray, int insertRowNum, int insertColNum)
        {
            int sourceNumRows = sourceArray.GetLength(0);
            int sourceNumCols = sourceArray.GetLength(1);

            // Bounds check
            if (insertRowNum + sourceNumRows > destArray.GetLength(0) ||
                insertColNum + sourceNumCols > destArray.GetLength(1))
            {
                throw new ArgumentException("Source array does not fit into destination array at the given position.");
            }

            // Copy data
            for (int rowNum = 0; rowNum < sourceNumRows; rowNum++)
            {
                for (int colNum = 0; colNum < sourceNumCols; colNum++)
                {
                    destArray[insertRowNum + rowNum, insertColNum + colNum] = sourceArray[rowNum, colNum];
                }
            }
        }

        public static void WriteArrayIntoArray(ref string[,] destArray, string[,] sourceArray, int insertRowNum, int insertColNum)
        {
            int sourceNumRows = sourceArray.GetLength(0);
            int sourceNumCols = sourceArray.GetLength(1);

            // Bounds check
            if (insertRowNum + sourceNumRows > destArray.GetLength(0) ||
                insertColNum + sourceNumCols > destArray.GetLength(1))
            {
                throw new ArgumentException("Source array does not fit into destination array at the given position.");
            }

            // Copy data
            for (int rowNum = 0; rowNum < sourceNumRows; rowNum++)
            {
                for (int colNum = 0; colNum < sourceNumCols; colNum++)
                {
                    destArray[insertRowNum + rowNum, insertColNum + colNum] = sourceArray[rowNum, colNum];
                }
            }
        }
    }

    
    class CustomFolderBrowser
    {
        OpenFileDialog dialog = new OpenFileDialog();
        public CustomFolderBrowser()
        {
            dialog.ValidateNames = false;  // Allows selecting folders
            dialog.Filter = "Folders|*. ";
            dialog.CheckFileExists = false;
            dialog.CheckPathExists = true;
            dialog.FileName = "Select Folder";  // Fake name to allow folder selection
        }

        public string folderPath = null;
        public DialogResult ShowDialog()
        {
            DialogResult dialogResult = dialog.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                string test = dialog.FileName;
                folderPath = Path.GetDirectoryName(dialog.FileName);
            }
            return dialogResult;
        }

        public string GetFolderPath()
        {
            if (folderPath == null)
            {
                throw new Exception("Folder path is not set");
            }
            return folderPath;
        }
        
        public void SetDirectory(string folderPatth)
        {
            dialog.InitialDirectory = folderPatth;
        }
    }
}

