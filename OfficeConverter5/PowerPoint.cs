using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

using Microsoft.Office.Core;
using Microsoft.Win32;

using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;

using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;

// PowerPoint.cs
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
    /// <summary>
    /// This class is used as a placeholder for all PowerPoint related methods
    /// </summary>
    internal class PowerPoint : IDisposable
    {
        #region Fields

        /// <summary>
        /// PowerPoint version number
        /// </summary>
        private readonly int _versionNumber;

        /// <summary>
        /// Keeps track is we already disposed our resources
        /// </summary>
        private bool _disposed;

        /// <summary>
        /// <see cref="PowerPointInterop.ApplicationClass"/>
        /// </summary>
        private PowerPointInterop.ApplicationClass _powerPoint;

        /// <summary>
        /// A <see cref="Process"/> object to PowerPoint
        /// </summary>
        private Process _powerPointProcess;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Returns <c> true </c> when PowerPoint is running
        /// </summary>
        /// <returns> </returns>
        private bool IsPowerPointRunning
        {
            get
            {
                if (_powerPointProcess == null)
                {
                    return false;
                }

                _powerPointProcess.Refresh();
                return !_powerPointProcess.HasExited;
            }
        }

        #endregion Properties

        #region Constructor

        /// <summary>
        /// This constructor checks to see if all requirements for a successful conversion are here.
        /// </summary>
        /// <exception cref="OCConfiguration">
        /// Raised when the registry could not be read to determine PowerPoint version
        /// </exception>
        internal PowerPoint()
        {
            Logger.WriteToLog("Checking what version of PowerPoint is installed");

            try
            {
                RegistryKey baseKey = Registry.ClassesRoot;
                RegistryKey subKey = baseKey.OpenSubKey(@"PowerPoint.Application\CurVer");
                _versionNumber = subKey != null
                    ? subKey.GetValue(string.Empty).ToString().ToUpperInvariant() switch
                    {
                        // PowerPoint 2003
                        "POWERPOINT.APPLICATION.11" => 11,
                        // PowerPoint 2007
                        "POWERPOINT.APPLICATION.12" => 12,
                        // PowerPoint 2010
                        "POWERPOINT.APPLICATION.14" => 14,
                        // PowerPoint 2013
                        "POWERPOINT.APPLICATION.15" => 15,
                        // PowerPoint 2016
                        "POWERPOINT.APPLICATION.16" => 16,
                        _ => throw new OCConfiguration("Could not determine PowerPoint version"),
                    }
                    : throw new OCConfiguration("Could not find registry key PowerPoint.Application\\CurVer");
            }
            catch (Exception exception)
            {
                throw new OCConfiguration("Could not read registry to check PowerPoint version", exception);
            }
        }

        #endregion Constructor

        #region StartPowerPoint

        /// <summary>
        /// Starts PowerPoint
        /// </summary>
        private void StartPowerPoint()
        {
            if (IsPowerPointRunning)
            {
                Logger.WriteToLog($"Powerpoint is already running on PID {_powerPointProcess.Id}... skipped");
                return;
            }

            Logger.WriteToLog("Starting PowerPoint");

            _powerPoint = new PowerPointInterop.ApplicationClass
            {
                DisplayAlerts = PowerPointInterop.PpAlertLevel.ppAlertsNone,
                DisplayDocumentInformationPanel = false,
                AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
            };

            _ = ProcessHelpers.GetWindowThreadProcessId(_powerPoint.HWND, out int processId);
            _powerPointProcess = Process.GetProcessById(processId);

            Logger.WriteToLog($"PowerPoint started with process id {_powerPointProcess.Id}");
        }

        #endregion StartPowerPoint

        #region StopPowerPoint

        /// <summary>
        /// Stops PowerPoint
        /// </summary>
        private void StopPowerPoint()
        {
            if (IsPowerPointRunning)
            {
                Logger.WriteToLog("Stopping PowerPoint");

                try
                {
                    _powerPoint.Quit();
                }
                catch (Exception exception)
                {
                    Logger.WriteToLog($"PowerPoint did not shutdown gracefully, exception: {ExceptionHelpers.GetInnerException(exception)}");
                }

                int counter = 0;

                // Give PowerPoint 2 seconds to close
                while (counter < 200)
                {
                    if (!IsPowerPointRunning)
                    {
                        break;
                    }

                    counter++;
                    Thread.Sleep(10);
                }

                if (IsPowerPointRunning)
                {
                    Logger.WriteToLog(
                        $"PowerPoint did not shutdown gracefully in 2 seconds ... killing it on process id {_powerPointProcess.Id}");
                    _powerPointProcess.Kill();
                    Logger.WriteToLog("PowerPoint process killed");
                }
                else
                {
                    Logger.WriteToLog("PowerPoint stopped");
                }
            }

            if (_powerPoint != null)
            {
                _ = Marshal.ReleaseComObject(_powerPoint);
                _powerPoint = null;
            }

            _powerPointProcess = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        #endregion StopPowerPoint

        #region Convert

        /// <summary>
        /// Converts a PowerPoint document to PDF
        /// </summary>
        /// <param name="inputFile"> The PowerPoint input file </param>
        /// <param name="outputFile"> The PDF output file </param>
        /// <returns> </returns>
        internal void Convert(string inputFile, string outputFile)
        {
            DeleteResiliencyKeys();

            PowerPointInterop.Presentation presentation = null;

            try
            {
                StartPowerPoint();

                presentation = OpenPresentation(inputFile, false);

                Logger.WriteToLog($"Exporting presentation to PDF file '{outputFile}'");
                presentation.ExportAsFixedFormat(outputFile, PowerPointInterop.PpFixedFormatType.ppFixedFormatTypePDF);
                Logger.WriteToLog("Presentation exported to PDF");
            }
            catch (Exception)
            {
                StopPowerPoint();
                throw;
            }
            finally
            {
                ClosePresentation(presentation);
            }
        }

        #endregion Convert

        #region OpenPresentation

        /// <summary>
        /// Opens the <paramref name="inputFile"/> and returns it as an <see
        /// cref="PowerPointInterop.Presentation"/> object
        /// </summary>
        /// <param name="inputFile"> The file to open </param>
        /// <param name="repairMode">
        /// When true the <paramref name="inputFile"/> is opened in repair mode
        /// </param>
        /// <returns> </returns>
        /// <exception cref="OCFileIsCorrupt">
        /// Raised when the <paramref name="inputFile"/> is corrupt and can't be opened in repair mode
        /// </exception>
        private PowerPointInterop.Presentation OpenPresentation(string inputFile, bool repairMode)
        {
            try
            {
                return _powerPoint.Presentations.Open(inputFile, MsoTriState.msoTrue, MsoTriState.msoTrue,
                    MsoTriState.msoFalse);
            }
            catch (Exception exception)
            {
                return repairMode
                    ? throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                              "' seems to be corrupt, error: " +
                                              ExceptionHelpers.GetInnerException(exception))
                    : OpenPresentation(inputFile, true);
            }
        }

        #endregion OpenPresentation

        #region ClosePresentation

        /// <summary>
        /// Closes the opened presentation and releases any allocated resources
        /// </summary>
        private static void ClosePresentation(PowerPointInterop.Presentation presentation)
        {
            if (presentation == null)
            {
                return;
            }

            Logger.WriteToLog("Closing presentation");
            presentation.Saved = MsoTriState.msoFalse;
            presentation.Close();
            _ = Marshal.ReleaseComObject(presentation);
            Logger.WriteToLog("Presentation closed");
        }

        #endregion ClosePresentation

        #region DeleteResiliencyKeys

        /// <summary>
        /// This method will delete the automatic created Resiliency key. PowerPoint uses this
        /// registry key to make entries to corrupted presentations. If there are to many entries
        /// under this key PowerPoint will get slower and slower to start. To prevent this we just
        /// delete this key when it exists
        /// </summary>
        private void DeleteResiliencyKeys()
        {
            Logger.WriteToLog("Deleting PowerPoint resiliency keys from the registry");

            try
            {
                // HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\PowerPoint\Resiliency\DocumentRecovery
                string key = $@"Software\Microsoft\Office\{_versionNumber}.0\PowerPoint\Resiliency";

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
        /// Disposes the running <see cref="_powerPoint"/>
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            StopPowerPoint();
        }

        #endregion Dispose
    }
}