using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;

using Microsoft.Office.Core;
using Microsoft.Win32;

using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;

using ExcelInterop = Microsoft.Office.Interop.Excel;

// Excel.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2014-2021 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
// associated documentation files (the "Software"), to deal in the Software without restriction,
// including without limitation the rights to use, copy, modify, merge, publish, distribute,
// sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
// NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NON
// INFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES
// OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR
// IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

namespace OfficeConverter
{
    #region Struct

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    // ReSharper disable once InconsistentNaming
    internal struct INTERFACEINFO
    {
        [MarshalAs(UnmanagedType.IUnknown)]
        public object punk;

        public Guid iid;
        public ushort wMethod;
    }

    #endregion Struct

    #region Interfaces

    [ComImport, ComConversionLoss, InterfaceType(1),
     Guid("00000016-0000-0000-C000-000000000046")]
    internal interface IMessageFilter
    {
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall,
             MethodCodeType = MethodCodeType.Runtime)]
        int HandleInComingCall([In] uint dwCallType, [In] IntPtr htaskCaller,
            [In] uint dwTickCount,
            [In, MarshalAs(UnmanagedType.LPArray)] INTERFACEINFO[]
                lpInterfaceInfo);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall,
             MethodCodeType = MethodCodeType.Runtime)]
        int RetryRejectedCall([In] IntPtr htaskCallee, [In] uint dwTickCount,
            [In] uint dwRejectType);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall,
             MethodCodeType = MethodCodeType.Runtime)]
        int MessagePending([In] IntPtr htaskCallee, [In] uint dwTickCount,
            [In] uint dwPendingType);
    }

    #endregion Interfaces

    #region MessageFilter

    internal class MessageFilter : IMessageFilter
    {
        public int HandleInComingCall(uint dwCallType, IntPtr htaskCaller, uint dwTickCount, INTERFACEINFO[] lpInterfaceInfo)
        {
            return 1;
        }

        public int MessagePending(IntPtr htaskCallee, uint dwTickCount, uint dwPendingType)
        {
            return 1;
        }

        public void RegisterFilter()
        {
            _ = CoRegisterMessageFilter(this, out _);
            Thread.Sleep(100);
        }

        // ReSharper disable once NotAccessedField.Local
        public int RetryRejectedCall(IntPtr htaskCallee, uint dwTickCount, uint dwRejectType)
        {
            return 1;
        }

        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(IMessageFilter lpMessageFilter, out IMessageFilter lplpMessageFilter);
    }

    #endregion MessageFilter

    /// <summary>
    /// This class is used as a placeholder for all Excel related methods
    /// </summary>
    internal class Excel : IDisposable
    {
        #region Private class ShapePosition

        /// <summary>
        /// Placeholder for shape information
        /// </summary>
        private class ShapePosition
        {
            /// <summary>
            /// Returns the bottom right column
            /// </summary>
            public int BottomRightColumn { get; }

            /// <summary>
            /// Returns the bottom right row
            /// </summary>
            public int BottomRightRow { get; }

            /// <summary>
            /// Returns the top left column
            /// </summary>
            public int TopLeftColumn { get; }

            /// <summary>
            /// Returns the top left row
            /// </summary>
            public int TopLeftRow { get; }

            /// <summary>
            /// Creates this object and sets it's needed properties
            /// </summary>
            /// <param name="shape"> The shape object </param>
            public ShapePosition(ExcelInterop.Shape shape)
            {
                ExcelInterop.Range topLeftCell = shape.TopLeftCell;
                ExcelInterop.Range bottomRightCell = shape.BottomRightCell;
                TopLeftRow = topLeftCell.Row;
                TopLeftColumn = topLeftCell.Column;
                BottomRightRow = bottomRightCell.Row;
                BottomRightColumn = bottomRightCell.Column;
                _ = Marshal.ReleaseComObject(topLeftCell);
                _ = Marshal.ReleaseComObject(bottomRightCell);
            }
        }

        #endregion Private class ShapePosition

        #region Private class ExcelPaperSize

        /// <summary>
        /// Placeholder for papersize and orientation information
        /// </summary>
        private class ExcelPaperSize
        {
            /// <summary>
            /// Returns the orientation
            /// </summary>
            public ExcelInterop.XlPageOrientation Orientation { get; }

            /// <summary>
            /// Returns the papersize
            /// </summary>
            public ExcelInterop.XlPaperSize PaperSize { get; }

            /// <summary>
            /// Creates this object and sets it's needed properties
            /// </summary>
            /// <param name="paperSize"> The papersize </param>
            /// <param name="orientation"> The orientation </param>
            public ExcelPaperSize(ExcelInterop.XlPaperSize paperSize, ExcelInterop.XlPageOrientation orientation)
            {
                PaperSize = paperSize;
                Orientation = orientation;
            }
        }

        #endregion Private class ExcelPaperSize

        #region Private enum MergedCellSearchOrder

        /// <summary>
        /// Direction to search in merged cells
        /// </summary>
        private enum MergedCellSearchOrder
        {
            /// <summary>
            /// Search for first row in the merge area
            /// </summary>
            FirstRow,

            /// <summary>
            /// Search for first column in the merge area
            /// </summary>
            FirstColumn,

            /// <summary>
            /// Search for last row in the merge area
            /// </summary>
            LastRow,

            /// <summary>
            /// Search for last column in the merge area
            /// </summary>
            LastColumn
        }

        #endregion Private enum MergedCellSearchOrder

        #region Fields

        /// <summary>
        /// Excel maximum rows
        /// </summary>
        private readonly int _maxRows;

        /// <summary>
        /// Paper sizes to use when detecting optimal page size with the <see
        /// cref="SetWorkSheetPaperSize"/> method
        /// </summary>
        private readonly List<ExcelPaperSize> _paperSizes = new()
        {
            new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA4, ExcelInterop.XlPageOrientation.xlPortrait),
            new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA4, ExcelInterop.XlPageOrientation.xlLandscape),
            new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA3, ExcelInterop.XlPageOrientation.xlLandscape),
            new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA3, ExcelInterop.XlPageOrientation.xlPortrait)
        };

        /// <summary>
        /// Excel version number
        /// </summary>
        private readonly int _versionNumber;

        /// <summary>
        /// Zoom ration to use when detecting optimal page size with the <see
        /// cref="SetWorkSheetPaperSize"/> method
        /// </summary>
        private readonly List<int> _zoomRatios = new() { 100, 95, 90, 85, 80, 75, 70 };

        /// <summary>
        /// Keeps track is we already disposed our resources
        /// </summary>
        private bool _disposed;

        /// <summary>
        /// <see cref="ExcelInterop.ApplicationClass"/>
        /// </summary>
        private ExcelInterop.ApplicationClass _excel;

        /// <summary>
        /// A <see cref="Process"/> object to Excel
        /// </summary>
        private Process _excelProcess;

        /// <summary>
        /// When set then this folder is used for temporary files
        /// </summary>
        private DirectoryInfo _tempDirectory;

        #endregion Fields

        #region Properties

        /// <summary>
        /// When set to <c> true </c> then the <see cref="TempDirectory"/> will not be deleted when
        /// the extraction is done
        /// </summary>
        /// <remarks> For debugging perpeses </remarks>
        public bool DoNotDeleteTempDirectory { get; set; }

        /// <summary>
        /// When set then this directory is used to store temporary files
        /// </summary>
        /// <exception cref="DirectoryNotFoundException">
        /// Raised when the given directory does not exists
        /// </exception>
        public string TempDirectory
        {
            get => _tempDirectory.FullName;
            set
            {
                if (!Directory.Exists(value))
                {
                    throw new DirectoryNotFoundException($"The directory '{value}' does not exists");
                }

                _tempDirectory = new DirectoryInfo(Path.Combine(value, Guid.NewGuid().ToString()));
            }
        }

        /// <summary>
        /// Returns a reference to the temp directory
        /// </summary>
        private DirectoryInfo GetTempDirectory
        {
            get
            {
                if (_tempDirectory == null)
                {
                    _tempDirectory = new DirectoryInfo(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString()));
                }

                if (!_tempDirectory.Exists)
                {
                    _tempDirectory.Create();
                }

                return _tempDirectory;
            }
        }

        /// <summary>
        /// Returns <c> true </c> when Excel is running
        /// </summary>
        /// <returns> </returns>
        private bool IsExcelRunning
        {
            get
            {
                if (_excelProcess == null)
                {
                    return false;
                }

                _excelProcess.Refresh();
                return !_excelProcess.HasExited;
            }
        }

        #endregion Properties

        #region Constructor

        /// <summary>
        /// This constructor checks to see if all requirements for a successful conversion are here.
        /// </summary>
        /// <exception cref="OCConfiguration">
        /// Raised when the registry could not be read to determine Excel version
        /// </exception>
        internal Excel()
        {
            MessageFilter messageFilter = new();
            messageFilter.RegisterFilter();

            Logger.WriteToLog("Checking what version of Excel is installed");

            try
            {
                RegistryKey baseKey = Registry.ClassesRoot;
                RegistryKey subKey = baseKey.OpenSubKey(@"Excel.Application\CurVer");
                if (subKey != null)
                {
                    switch (subKey.GetValue(string.Empty).ToString().ToUpperInvariant())
                    {
                        // Excel 2003
                        case "EXCEL.APPLICATION.11":
                            _versionNumber = 11;
                            Logger.WriteToLog("Excel 2003 is installed");
                            break;

                        // Excel 2007
                        case "EXCEL.APPLICATION.12":
                            _versionNumber = 12;
                            Logger.WriteToLog("Excel 2007 is installed");
                            break;

                        // Excel 2010
                        case "EXCEL.APPLICATION.14":
                            _versionNumber = 14;
                            Logger.WriteToLog("Excel 2010 is installed");
                            break;

                        // Excel 2013
                        case "EXCEL.APPLICATION.15":
                            _versionNumber = 15;
                            Logger.WriteToLog("Excel 2013 is installed");
                            break;

                        // Excel 2016
                        case "EXCEL.APPLICATION.16":
                            _versionNumber = 16;
                            Logger.WriteToLog("Excel 2016 is installed");
                            break;

                        // Excel 2019
                        case "EXCEL.APPLICATION.17":
                            _versionNumber = 17;
                            Logger.WriteToLog("Excel 2019 is installed");
                            break;

                        default:
                            throw new OCConfiguration("Could not determine Excel version");
                    }
                }
                else
                {
                    throw new OCConfiguration("Could not find registry key Excel.Application\\CurVer");
                }
            }
            catch (Exception exception)
            {
                throw new OCConfiguration("Could not read registry to check Excel version", exception);
            }

            const int excelMaxRowsFrom2003AndBelow = 65535;
            const int excelMaxRowsFrom2007AndUp = 1048576;

            _maxRows = _versionNumber switch
            {
                // Excel 2007
                12 or 14 or 15 or 16 or 17 => excelMaxRowsFrom2007AndUp,
                // Excel 2003 and older
                _ => excelMaxRowsFrom2003AndBelow,
            };
            Logger.WriteToLog($"Setting maximum Excel rows to {_maxRows}");

            // We only need to perform this check if we are running on a server
            if (NativeMethods.IsWindowsServer())
            {
                CheckIfSystemProfileDesktopDirectoryExists();
            }

            CheckIfPrinterIsInstalled();
        }

        #endregion Constructor

        #region StartExcel

        /// <summary>
        /// Starts Excel
        /// </summary>
        private void StartExcel()
        {
            if (IsExcelRunning)
            {
                Logger.WriteToLog($"Excel is already running on PID {_excelProcess.Id}... skipped");
                return;
            }

            Logger.WriteToLog("Starting Excel");

            _excel = new ExcelInterop.ApplicationClass
            {
                Interactive = false,
                ScreenUpdating = false,
                DisplayAlerts = false,
                DisplayDocumentInformationPanel = false,
                DisplayRecentFiles = false,
                DisplayScrollBars = false,
                AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable,
                PrintCommunication = true // DO NOT REMOVE THIS LINE, NO NEVER EVER ... DON'T EVEN TRY IT
            };

            _ = ProcessHelpers.GetWindowThreadProcessId(_excel.Hwnd, out int processId);
            _excelProcess = Process.GetProcessById(processId);

            Logger.WriteToLog($"Excel started with process id {_excelProcess.Id}");
        }

        #endregion StartExcel

        #region StopExcel

        /// <summary>
        /// Stops Excel
        /// </summary>
        private void StopExcel()
        {
            if (IsExcelRunning)
            {
                Logger.WriteToLog("Stopping Excel");

                try
                {
                    _excel.Quit();
                }
                catch (Exception exception)
                {
                    Logger.WriteToLog($"Excel did not shutdown gracefully, exception: {ExceptionHelpers.GetInnerException(exception)}");
                }

                int counter = 0;

                // Give Excel 2 seconds to close
                while (counter < 200)
                {
                    if (!IsExcelRunning)
                    {
                        break;
                    }

                    counter++;
                    Thread.Sleep(10);
                }

                if (IsExcelRunning)
                {
                    Logger.WriteToLog($"Excel did not shutdown gracefully... killing it on process id {_excelProcess.Id}");
                    _excelProcess.Kill();
                    _excelProcess = null;
                    Logger.WriteToLog("Excel process killed");
                }
                else
                {
                    Logger.WriteToLog("Excel stopped");
                }
            }

            if (_excel != null)
            {
                _ = Marshal.ReleaseComObject(_excel);
                _excel = null;
            }

            _excelProcess = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        #endregion StopExcel

        #region CheckIfSystemProfileDesktopDirectoryExists

        /// <summary>
        /// If you want to run this code on a server then the following folders must exist, if they
        /// don't then you can't use Excel to convert files to PDF
        /// </summary>
        /// <exception cref="OCConfiguration">
        /// Raised when the needed directory could not be created
        /// </exception>
        private static void CheckIfSystemProfileDesktopDirectoryExists()
        {
            if (Environment.Is64BitOperatingSystem)
            {
                string x64DesktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    @"SysWOW64\config\systemprofile\desktop");

                Logger.WriteToLog($"Checking if system profile desktop directory exists in '{x64DesktopPath}'");

                if (!Directory.Exists(x64DesktopPath))
                {
                    try
                    {
                        _ = Directory.CreateDirectory(x64DesktopPath);
                        Logger.WriteToLog("Directory did not exist ... created it");
                    }
                    catch (Exception exception)
                    {
                        throw new OCConfiguration("Can't create folder '" + x64DesktopPath +
                                                  "' Excel needs this folder to work on a server, error: " +
                                                  ExceptionHelpers.GetInnerException(exception));
                    }
                }
            }
            else
            {
                string x86DesktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    @"System32\config\systemprofile\desktop");

                Logger.WriteToLog($"Checking if system profile desktop directory exists in '{x86DesktopPath}'");

                if (!Directory.Exists(x86DesktopPath))
                {
                    try
                    {
                        _ = Directory.CreateDirectory(x86DesktopPath);
                        Logger.WriteToLog("Directory did not exist ... created it");
                    }
                    catch (Exception exception)
                    {
                        throw new OCConfiguration("Can't create folder '" + x86DesktopPath +
                                                  "' Excel needs this folder to work on a server, error: " +
                                                  ExceptionHelpers.GetInnerException(exception));
                    }
                }
            }
        }

        #endregion CheckIfSystemProfileDesktopDirectoryExists

        #region CheckIfPrinterIsInstalled

        /// <summary>
        /// Excel needs a default printer to export to PDF, this method will check if there is one
        /// </summary>
        /// <exception cref="OCConfiguration"> Raised when an default printer does not exists </exception>
        private static void CheckIfPrinterIsInstalled()
        {
            Logger.WriteToLog("Excel needs a printer to convert sheets to pdf ... checking if a printer exists");

            bool result = false;

            PrinterSettings.StringCollection installedPrinters;

            try
            {
                installedPrinters = PrinterSettings.InstalledPrinters;
            }
            catch (Win32Exception win32Exception)
            {
                throw new OCConfiguration($"Printer spooler service not enabled, error: {ExceptionHelpers.GetInnerException(win32Exception)}");
            }

            foreach (string printerName in installedPrinters)
            {
                // Retrieve the printer settings.
                PrinterSettings printer = new() { PrinterName = printerName };

                // Check that this is a valid printer. (This step might be required if you read the
                // printer name from a user-supplied value or a registry or configuration file setting.)
                if (printer.IsValid)
                {
                    Logger.WriteToLog($"A valid printer '{printer.PrinterName}' is found");
                    result = true;
                    break;
                }
            }

            if (!result)
            {
                throw new OCConfiguration("There is no default printer installed, Excel needs one to export to PDF");
            }
        }

        #endregion CheckIfPrinterIsInstalled

        #region GetColumnAddress

        /// <summary>
        /// Returns the column address for the given <paramref name="column"/>
        /// </summary>
        /// <param name="column"> </param>
        /// <returns> </returns>
        private string GetColumnAddress(int column)
        {
            if (column <= 26)
            {
                return System.Convert.ToChar(column + 64).ToString(CultureInfo.InvariantCulture);
            }

            int div = column / 26;
            int mod = column % 26;
            if (mod != 0)
            {
                return GetColumnAddress(div) + GetColumnAddress(mod);
            }

            mod = 26;
            div--;

            return GetColumnAddress(div) + GetColumnAddress(mod);
        }

        #endregion GetColumnAddress

        #region GetColumnNumber

        #endregion GetColumnNumber

        #region CheckForMergedCell

        /// <summary>
        /// Checks if the given cell is merged and if so returns the last column or row from this
        /// merge. When the cell is not merged it just returns the cell
        /// </summary>
        /// <param name="range"> The cell </param>
        /// <param name="searchOrder"> <see cref="MergedCellSearchOrder"/> </param>
        /// <returns> </returns>
        private static int CheckForMergedCell(ExcelInterop.Range range, MergedCellSearchOrder searchOrder)
        {
            if (range == null)
            {
                return 0;
            }

            int result = 0;
            ExcelInterop.Range mergeArea = range.MergeArea;

            switch (searchOrder)
            {
                case MergedCellSearchOrder.FirstRow:
                    result = mergeArea.Row;
                    break;

                case MergedCellSearchOrder.FirstColumn:
                    result = mergeArea.Column;
                    break;

                case MergedCellSearchOrder.LastRow:
                    {
                        result = range.Row;
                        ExcelInterop.Range entireRow = range.EntireRow;

                        for (int i = 1; i < range.Column; i++)
                        {
                            ExcelInterop.Range cell = (ExcelInterop.Range)entireRow.Cells[i];
                            ExcelInterop.Range cellMergeArea = cell.MergeArea;
                            ExcelInterop.Range cellMergeAreaRows = cellMergeArea.Rows;

                            _ = Marshal.ReleaseComObject(cellMergeAreaRows);
                            _ = Marshal.ReleaseComObject(cellMergeArea);
                            _ = Marshal.ReleaseComObject(cell);

                            int tempResult = result;

                            if (cellMergeAreaRows.Count > 1 && range.Row + cellMergeAreaRows.Count > tempResult)
                            {
                                tempResult = result + cellMergeAreaRows.Count;
                            }

                            result = tempResult;
                        }

                        _ = Marshal.ReleaseComObject(entireRow);

                        break;
                    }

                case MergedCellSearchOrder.LastColumn:
                    {
                        result = range.Column;
                        ExcelInterop.Range columns = mergeArea.Columns;

                        if (columns.Count > 1)
                        {
                            result += columns.Count;
                        }

                        _ = Marshal.ReleaseComObject(columns);

                        break;
                    }
            }

            if (mergeArea != null)
            {
                _ = Marshal.ReleaseComObject(mergeArea);
            }

            return result;
        }

        #endregion CheckForMergedCell

        #region GetWorksheetPrintArea

        /// <summary>
        /// Figures out the used cell range. This are the cell's that contain any form of text and
        /// returns this range. An empty range will be returned when there are shapes used on a worksheet
        /// </summary>
        /// <param name="worksheet"> </param>
        /// <returns> </returns>
        private string GetWorksheetPrintArea(ExcelInterop._Worksheet worksheet)
        {
            int firstColumn = 1;
            int firstRow = 1;

            List<ShapePosition> shapesPosition = new();

            // We can't use this method when there are shapes on a sheet so we return an empty string
            ExcelInterop.Shapes shapes = worksheet.Shapes;
            if (shapes.Count > 0)
            {
                if (_versionNumber < 14)
                {
                    return "shapes";
                }

                // The shape TopLeftCell and BottomRightCell is only supported from Excel 2010 and up
                foreach (ExcelInterop.Shape shape in worksheet.Shapes)
                {
                    if (shape.AutoShapeType != MsoAutoShapeType.msoShapeMixed)
                    {
                        shapesPosition.Add(new ShapePosition(shape));
                    }

                    _ = Marshal.ReleaseComObject(shape);
                }

                _ = Marshal.ReleaseComObject(shapes);
            }

            ExcelInterop.Range range = worksheet.Cells[1, 1] as ExcelInterop.Range;
            if (range?.Value == null)
            {
                if (range != null)
                {
                    _ = Marshal.ReleaseComObject(range);
                }

                ExcelInterop.Range firstCellByColumn = worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByColumns);
                bool foundByFirstColumn = false;
                if (firstCellByColumn != null)
                {
                    foundByFirstColumn = true;
                    firstColumn = CheckForMergedCell(firstCellByColumn, MergedCellSearchOrder.FirstColumn);
                    firstRow = CheckForMergedCell(firstCellByColumn, MergedCellSearchOrder.FirstRow);
                    _ = Marshal.ReleaseComObject(firstCellByColumn);
                }

                // Search the first used cell row wise
                ExcelInterop.Range firstCellByRow = worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByRows);
                if (firstCellByRow == null)
                {
                    return string.Empty;
                }

                if (foundByFirstColumn)
                {
                    if (firstCellByRow.Column < firstColumn)
                    {
                        firstColumn = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstColumn);
                    }

                    if (firstCellByRow.Row < firstRow)
                    {
                        firstRow = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstRow);
                    }
                }
                else
                {
                    firstColumn = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstColumn);
                    firstRow = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstRow);
                }

                _ = Marshal.ReleaseComObject(firstCellByRow);
            }

            foreach (ShapePosition shapePosition in shapesPosition)
            {
                if (shapePosition.TopLeftColumn < firstColumn)
                {
                    firstColumn = shapePosition.TopLeftColumn;
                }

                if (shapePosition.TopLeftRow < firstRow)
                {
                    firstRow = shapePosition.TopLeftRow;
                }
            }

            int lastColumn = firstColumn;
            int lastRow = firstRow;

            ExcelInterop.Range lastCellByColumn =
                worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByColumns,
                    SearchDirection: ExcelInterop.XlSearchDirection.xlPrevious);

            if (lastCellByColumn != null)
            {
                lastColumn = lastCellByColumn.Column;
                lastRow = lastCellByColumn.Row;
                _ = Marshal.ReleaseComObject(lastCellByColumn);
            }

            ExcelInterop.Range lastCellByRow =
                worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByRows,
                    SearchDirection: ExcelInterop.XlSearchDirection.xlPrevious);

            if (lastCellByRow != null)
            {
                if (lastCellByRow.Column > lastColumn)
                {
                    lastColumn = CheckForMergedCell(lastCellByRow, MergedCellSearchOrder.LastColumn);
                }

                if (lastCellByRow.Row > lastRow)
                {
                    lastRow = CheckForMergedCell(lastCellByRow, MergedCellSearchOrder.LastRow);
                }

                ExcelInterop.Protection protection = worksheet.Protection;
                if (!worksheet.ProtectContents || protection.AllowDeletingRows)
                {
                    ExcelInterop.Range previousLastCellByRow =
                        worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByRows,
                            SearchDirection: ExcelInterop.XlSearchDirection.xlPrevious,
                            After: lastCellByRow);

                    _ = Marshal.ReleaseComObject(lastCellByRow);

                    if (previousLastCellByRow != null)
                    {
                        int previousRow = CheckForMergedCell(previousLastCellByRow, MergedCellSearchOrder.LastRow);
                        _ = Marshal.ReleaseComObject(previousLastCellByRow);

                        if (previousRow < lastRow - 2)
                        {
                            ExcelInterop.Range rangeToDelete =
                                worksheet.Range[GetColumnAddress(firstColumn) + (previousRow + 1) + ":" +
                                                GetColumnAddress(lastColumn) + (lastRow - 2)];

                            _ = rangeToDelete.Delete(ExcelInterop.XlDeleteShiftDirection.xlShiftUp);
                            _ = Marshal.ReleaseComObject(rangeToDelete);
                            lastRow = previousRow + 2;
                        }
                    }

                    _ = Marshal.ReleaseComObject(protection);
                }
            }

            foreach (ShapePosition shapePosition in shapesPosition)
            {
                if (shapePosition.BottomRightColumn > lastColumn)
                {
                    lastColumn = shapePosition.BottomRightColumn;
                }

                if (shapePosition.BottomRightRow > lastRow)
                {
                    lastRow = shapePosition.BottomRightRow;
                }
            }

            return GetColumnAddress(firstColumn) + firstRow + ":" +
                   GetColumnAddress(lastColumn) + lastRow;
        }

        #endregion GetWorksheetPrintArea

        #region CountVerticalPageBreaks

        /// <summary>
        /// Returns the total number of vertical pagebreaks in the print area
        /// </summary>
        /// <param name="pageBreaks"> </param>
        /// <returns> </returns>
        private static int CountVerticalPageBreaks(ExcelInterop.VPageBreaks pageBreaks)
        {
            int result = 0;

            try
            {
                foreach (ExcelInterop.VPageBreak pageBreak in pageBreaks)
                {
                    if (pageBreak.Extent == ExcelInterop.XlPageBreakExtent.xlPageBreakPartial)
                    {
                        result += 1;
                    }

                    _ = Marshal.ReleaseComObject(pageBreak);
                }
            }
            catch (COMException)
            {
                result = pageBreaks.Count;
            }

            return result;
        }

        #endregion CountVerticalPageBreaks

        #region SetWorkSheetPaperSize

        /// <summary>
        /// This method wil figure out the optimal paper size to use and sets it
        /// </summary>
        /// <param name="worksheet"> </param>
        /// <param name="printArea"> </param>
        private void SetWorkSheetPaperSize(ExcelInterop._Worksheet worksheet, string printArea)
        {
            Logger.WriteToLog($"Detecting optimal paper size for sheet {worksheet.Name} with print area '{printArea}'");

            ExcelInterop.PageSetup pageSetup = worksheet.PageSetup;
            ExcelInterop.Pages pages = pageSetup.Pages;

            pageSetup.PrintArea = printArea;
            pageSetup.LeftHeader = worksheet.Name;

            int pageCount = pages.Count;

            if (pageCount == 1)
            {
                return;
            }

            try
            {
                pageSetup.Order = ExcelInterop.XlOrder.xlOverThenDown;

                foreach (ExcelPaperSize paperSize in _paperSizes)
                {
                    bool exitfor = false;
                    pageSetup.PaperSize = paperSize.PaperSize;
                    pageSetup.Orientation = paperSize.Orientation;
                    worksheet.ResetAllPageBreaks();

                    foreach (int zoomRatio in _zoomRatios)
                    {
                        // Yes these page counts look lame, but so is Excel 2010 in not updating the
                        // pages collection otherwise. We need to call the count methods to make
                        // this code work
                        pageSetup.Zoom = zoomRatio;
                        // ReSharper disable once RedundantAssignment
                        pageCount = pages.Count;

                        if (CountVerticalPageBreaks(worksheet.VPageBreaks) == 0)
                        {
                            exitfor = true;
                            break;
                        }
                    }

                    if (exitfor)
                    {
                        break;
                    }
                }

                Logger.WriteToLog($"Paper size set to '{pageSetup.PaperSize}', orientation to '{pageSetup.Orientation}' and zoom ratio to '{pageSetup.Zoom}'");
            }
            finally
            {
                _ = Marshal.ReleaseComObject(pages);
                _ = Marshal.ReleaseComObject(pageSetup);
            }
        }

        #endregion SetWorkSheetPaperSize

        #region SetChartPaperSize

        /// <summary>
        /// This method wil set the papersize for a chart
        /// </summary>
        /// <param name="chart"> </param>
        private static void SetChartPaperSize(ExcelInterop._Chart chart)
        {
            Logger.WriteToLog($"Setting paper site for chart '{chart.Name}' to A4 landscape");

            ExcelInterop.PageSetup pageSetup = chart.PageSetup;
            ExcelInterop.Pages pages = pageSetup.Pages;

            try
            {
                pageSetup.LeftHeader = chart.Name;
                pageSetup.PaperSize = ExcelInterop.XlPaperSize.xlPaperA4;
                pageSetup.Orientation = ExcelInterop.XlPageOrientation.xlLandscape;
            }
            finally
            {
                _ = Marshal.ReleaseComObject(pages);
                _ = Marshal.ReleaseComObject(pageSetup);
            }
        }

        #endregion SetChartPaperSize

        #region Convert

        /// <summary>
        /// Converts an Excel sheet to PDF
        /// </summary>
        /// <param name="inputFile"> The Excel input file </param>
        /// <param name="outputFile"> The PDF output file </param>
        /// <returns> </returns>
        /// <exception cref="OCCsvFileLimitExceeded">
        /// Raised when a CSV <paramref name="inputFile"/> has to many rows
        /// </exception>
        internal void Convert(string inputFile, string outputFile)
        {
            DeleteResiliencyKeys();

            ExcelInterop.Workbook workbook = null;

            try
            {
                StartExcel();

                string extension = Path.GetExtension(inputFile);
                if (string.IsNullOrWhiteSpace(extension))
                {
                    extension = string.Empty;
                }

                if (extension.ToUpperInvariant() == ".CSV")
                {
                    string tempFileName = Path.Combine(GetTempDirectory.FullName, Guid.NewGuid() + ".txt");

                    // Yes this look somewhat weird but we have to change the extension if we want
                    // to handle CSV files with different kind of separators. Otherwhise Excel will
                    // always overrule whatever setting we make to open a file
                    Logger.WriteToLog($"Copying CSV file '{inputFile}' to temporary file '{tempFileName}' and setting that one as the input file");
                    File.Copy(inputFile, tempFileName);
                    inputFile = tempFileName;
                }

                workbook = OpenWorkbook(inputFile, extension, false);

                // We cannot determine a print area when the document is marked as final so we
                // remove this
                workbook.Final = false;

                // Fix for "This command is not available in a shared workbook."
                if (workbook.MultiUserEditing)
                {
                    string tempFileName = Path.Combine(GetTempDirectory.FullName, Guid.NewGuid() + Path.GetExtension(inputFile));
                    Logger.WriteToLog($"Excel file '{inputFile}' is in 'multi user editing' mode saving it to temporary file '{tempFileName}' to set it to exclusive mode");
                    workbook.SaveAs(tempFileName, AccessMode: ExcelInterop.XlSaveAsAccessMode.xlExclusive);
                }

                int usedSheets = 0;

                ExcelInterop.Window activeWindow = _excel.ActiveWindow;

                if (activeWindow == null)
                {
                    const string message = "There is no window active in Excel";
                    Logger.WriteToLog(message);
                    throw new OCFileContainsNoData(message);
                }

                foreach (object sheetObject in workbook.Sheets)
                {
                    switch (sheetObject)
                    {
                        // Invisible sheets will not be converted... they are not visible
                        case ExcelInterop.Worksheet sheet
                            when sheet.Visible != ExcelInterop.XlSheetVisibility.xlSheetVisible:
                            continue;

                        case ExcelInterop.Worksheet sheet:
                            ExcelInterop.Protection protection = sheet.Protection;

                            try
                            {
                                // ReSharper disable once RedundantCast
                                (sheet as ExcelInterop._Worksheet).Activate();
                                if (!sheet.ProtectContents || protection.AllowFormattingColumns)
                                {
                                    if (activeWindow.View != ExcelInterop.XlWindowView.xlPageLayoutView)
                                    {
                                        Logger.WriteToLog($"Auto fitting colums on sheet '{sheet.Name}'");
                                        _ = sheet.Columns.AutoFit();
                                    }
                                }
                            }
                            catch (COMException)
                            {
                                // Do nothing, this sometimes failes and there is nothing we can do
                                // about it
                            }
                            finally
                            {
                                _ = Marshal.ReleaseComObject(protection);
                            }

                            string printArea = GetWorksheetPrintArea(sheet);
                            Logger.WriteToLog($"Print area for sheet {sheet.Name} set to '{printArea}'");

                            switch (printArea)
                            {
                                case "shapes":
                                    SetWorkSheetPaperSize(sheet, string.Empty);
                                    usedSheets += 1;
                                    break;

                                case "":
                                    if (sheet.Shapes.Count > 0)
                                    {
                                        usedSheets += 1;
                                    }

                                    break;

                                default:
                                    SetWorkSheetPaperSize(sheet, printArea);
                                    usedSheets += 1;
                                    break;
                            }

                            _ = Marshal.ReleaseComObject(sheet);
                            continue;
                    }

                    if (sheetObject is not ExcelInterop.Chart chart)
                    {
                        continue;
                    }

                    SetChartPaperSize(chart);
                    _ = Marshal.ReleaseComObject(chart);
                }

                _ = Marshal.ReleaseComObject(activeWindow);

                // It is not possible in Excel to export an empty workbook
                if (usedSheets != 0)
                {
                    Logger.WriteToLog($"Exporting worksheets to PDF file '{outputFile}'");
                    workbook.ExportAsFixedFormat(ExcelInterop.XlFixedFormatType.xlTypePDF, outputFile);
                    Logger.WriteToLog("Worksheets exported to PDF");
                }
                else
                {
                    const string message = "The file contains no data";
                    Logger.WriteToLog(message);
                    throw new OCFileContainsNoData(message);
                }
            }
            catch (Exception)
            {
                StopExcel();
                throw;
            }
            finally
            {
                try
                {
                    CloseWorkbook(workbook);
                }
                catch (Exception exception)
                {
                    Logger.WriteToLog("Error closing workbook, error: " + ExceptionHelpers.GetInnerException(exception));
                }

                if (_tempDirectory != null)
                {
                    _tempDirectory.Refresh();
                    if (_tempDirectory.Exists && !DoNotDeleteTempDirectory)
                    {
                        Logger.WriteToLog($"Deleting temporary folder '{_tempDirectory.FullName}'");
                        _tempDirectory.Delete(true);
                    }
                }
            }
        }

        #endregion Convert

        #region GetCsvSeperator

        /// <summary>
        /// Returns the separator and text qualifier that is used in the CSV file
        /// </summary>
        /// <param name="inputFile"> The input file </param>
        /// <param name="separator"> The separator that is used </param>
        /// <param name="textQualifier"> The text qualifier </param>
        /// <returns> </returns>
        private static void GetCsvSeparator(string inputFile, out string separator,
            out ExcelInterop.XlTextQualifier textQualifier)
        {
            separator = string.Empty;
            textQualifier = ExcelInterop.XlTextQualifier.xlTextQualifierNone;

            using (StreamReader streamReader = new(inputFile))
            {
                string line = string.Empty;
                while (string.IsNullOrEmpty(line))
                {
                    line = streamReader.ReadLine();
                }

                if (line.Contains(";"))
                {
                    separator = ";";
                }
                else if (line.Contains(","))
                {
                    separator = ",";
                }
                else if (line.Contains("\t"))
                {
                    separator = "\t";
                }
                else if (line.Contains(" "))
                {
                    separator = " ";
                }

                if (line.Contains("\""))
                {
                    textQualifier = ExcelInterop.XlTextQualifier.xlTextQualifierDoubleQuote;
                }
                else if (line.Contains("'"))
                {
                    textQualifier = ExcelInterop.XlTextQualifier.xlTextQualifierSingleQuote;
                }
            }
        }

        #endregion GetCsvSeperator

        #region OpenWorkbook

        /// <summary>
        /// Opens the <paramref name="inputFile"/> and returns it as an <see
        /// cref="ExcelInterop.Workbook"/> object
        /// </summary>
        /// <param name="inputFile"> The file to open </param>
        /// <param name="extension"> The file extension </param>
        /// <param name="repairMode">
        /// When true the <paramref name="inputFile"/> is opened in repair mode
        /// </param>
        /// <returns> </returns>
        /// <exception cref="OCCsvFileLimitExceeded">
        /// Raised when a CSV <paramref name="inputFile"/> has to many rows
        /// </exception>
        private ExcelInterop.Workbook OpenWorkbook(string inputFile, string extension, bool repairMode)
        {
            Logger.WriteToLog($"Opening workbook '{inputFile}'{(repairMode ? " with repair mode" : string.Empty)}");

            try
            {
                switch (extension.ToUpperInvariant())
                {
                    case ".CSV":

                        int count = File.ReadLines(inputFile).Count();
                        int excelMaxRows = _maxRows;
                        if (count > excelMaxRows)
                        {
                            throw new OCCsvFileLimitExceeded("The input CSV file has more then " + excelMaxRows +
                                                             " rows, the installed Excel version supports only " +
                                                             excelMaxRows + " rows");
                        }

                        GetCsvSeparator(inputFile, out string separator, out Microsoft.Office.Interop.Excel.XlTextQualifier textQualifier);
                        Logger.WriteToLog($"Separator for CSV file set to '{separator}' and text qualifier to '{textQualifier}'");

                        switch (separator)
                        {
                            case ";":
                                _excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited,
                                    textQualifier, true, false, true);
                                break;

                            case ",":
                                _excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, false, true);
                                break;

                            case "\t":
                                _excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, true);
                                break;

                            case " ":
                                _excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, false, false, true);
                                break;

                            default:
                                _excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, true);
                                break;
                        }

                        Logger.WriteToLog("Workbook opened");
                        return _excel.ActiveWorkbook;

                    default:

                        ExcelInterop.Workbook workbook;

                        workbook = repairMode
                            ? _excel.Workbooks.Open(inputFile, false, true,
                                Password: "dummy password",
                                IgnoreReadOnlyRecommended: true,
                                AddToMru: false,
                                CorruptLoad: ExcelInterop.XlCorruptLoad.xlRepairFile)
                            : _excel.Workbooks.Open(inputFile, false, true,
                                Password: "dummy password",
                                IgnoreReadOnlyRecommended: true,
                                AddToMru: false);

                        Logger.WriteToLog("Workbook opened");
                        return workbook;
                }
            }
            catch (COMException comException)
            {
                if (comException.ErrorCode == -2146827284)
                {
                    throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                        "' is password protected");
                }

                throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                          "' could not be opened, error: " +
                                          ExceptionHelpers.GetInnerException(comException));
            }
            catch (Exception exception)
            {
                Logger.WriteToLog(
                    $"ERROR: Failed to open worksheet, exception: '{ExceptionHelpers.GetInnerException(exception)}'");

                return repairMode
                    ? throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                              "' could not be opened, error: " +
                                              ExceptionHelpers.GetInnerException(exception))
                    : OpenWorkbook(inputFile, extension, true);
            }
        }

        #endregion OpenWorkbook

        #region CloseWorkbook

        /// <summary>
        /// Closes the opened workbook and releases any allocated resources
        /// </summary>
        /// <param name="workbook"> The Excel workbook </param>
        private static void CloseWorkbook(ExcelInterop.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            Logger.WriteToLog("Closing workbook");
            workbook.Saved = true;
            workbook.Close(false);
            _ = Marshal.ReleaseComObject(workbook);
            Logger.WriteToLog("Workbook closed");
        }

        #endregion CloseWorkbook

        #region DeleteResiliencyKeys

        /// <summary>
        /// This method will delete the automatic created Resiliency key. Excel uses this registry
        /// key to make entries to corrupted workbooks. If there are to many entries under this key
        /// Excel will get slower and slower to start. To prevent this we just delete this key when
        /// it exists
        /// </summary>
        private void DeleteResiliencyKeys()
        {
            Logger.WriteToLog("Deleting Excel resiliency keys from the registry");

            try
            {
                // HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Excel\Resiliency\DocumentRecovery
                string key = $@"Software\Microsoft\Office\{_versionNumber}.0\Excel\Resiliency";

                if (Registry.CurrentUser.OpenSubKey(key, false) != null)
                {
                    Registry.CurrentUser.DeleteSubKeyTree(key);
                    Logger.WriteToLog("Resiliency keys deleted");
                }
                else
                {
                    Logger.WriteToLog("There are no keys to delete");
                }
            }
            catch (Exception exception)
            {
                Logger.WriteToLog($"Failed to delete resiliency keys, error: {ExceptionHelpers.GetInnerException(exception)}");
            }
        }

        #endregion DeleteResiliencyKeys

        #region Dispose

        /// <summary>
        /// Disposes the running <see cref="_excel"/>
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            StopExcel();
        }

        #endregion Dispose
    }
}