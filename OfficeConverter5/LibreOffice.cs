using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

using Microsoft.Win32;

using OfficeConverter.Helpers;

using uno;
using uno.util;

using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.bridge;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.util;

using Exception = System.Exception;

// LibreOffice.cs
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
    /// This class is used as a placeholder for all Libre office related methods
    /// </summary>
    /// <remarks>
    /// - https://api.libreoffice.org/examples/examples.html
    /// - https://api.libreoffice.org/docs/install.html
    /// - https://www.libreoffice.org/download/download/
    /// </remarks>
    internal class LibreOffice : IDisposable
    {
        #region Fields

        /// <summary>
        /// <see cref="XComponentLoader"/>
        /// </summary>
        private XComponentLoader _componentLoader;

        /// <summary>
        /// Keeps track is we already disposed our resources
        /// </summary>
        private bool _disposed;

        /// <summary>
        /// A <see cref="Process"/> object to LibreOffice
        /// </summary>
        private Process _libreOfficeProcess;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Returns the full path to LibreOffice, when not found <c> null </c> is returned
        /// </summary>
        private static string GetInstallPath
        {
            get
            {
                using (RegistryKey hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                using (RegistryKey regkey64 = hklm.OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath", false))
                {
                    string installPath = (string)regkey64?.GetValue(string.Empty);

                    if (installPath != null)
                    {
                        return installPath;
                    }

                    using (RegistryKey regkey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath", false))
                    {
                        installPath = (string)regkey?.GetValue(string.Empty);
                        return installPath;
                    }
                }
            }
        }

        #region Properties

        /// <summary>
        /// Returns <c> true </c> when LibreOffice is running
        /// </summary>
        /// <returns> </returns>
        private bool IsLibreOfficeRunning
        {
            get
            {
                if (_libreOfficeProcess == null)
                {
                    return false;
                }

                _libreOfficeProcess.Refresh();
                return !_libreOfficeProcess.HasExited;
            }
        }

        #endregion Properties

        #endregion Properties

        #region StartLibreOffice

        /// <summary>
        /// Checks if LibreOffice is started and if not starts it
        /// </summary>
        private void StartLibreOffice()
        {
            if (IsLibreOfficeRunning)
            {
                Logger.WriteToLog($"LibreOffice is already running on PID {_libreOfficeProcess.Id}... skipped");
                return;
            }

            string installPath = GetInstallPath;
            if (string.IsNullOrEmpty(installPath))
            {
                throw new InvalidProgramException("LibreOffice is not installed");
            }

            string path = installPath.Replace('\\', '/');

            string ureBootStrap = $"vnd.sun.star.pathname:{path}/fundamental.ini";
            Logger.WriteToLog($"Setting environment variable URE_BOOTSTRAP to '{ureBootStrap}'");
            Environment.SetEnvironmentVariable("URE_BOOTSTRAP", $"vnd.sun.star.pathname:{path}/fundamental.ini", EnvironmentVariableTarget.Process);

            string environmentPath = Environment.GetEnvironmentVariable("PATH");
            Logger.WriteToLog($"Setting environment variable UNO_PATH to '{path}'");
            Environment.SetEnvironmentVariable("UNO_PATH", path, EnvironmentVariableTarget.Process);

            if (environmentPath != null && !environmentPath.Contains(path))
            {
                Logger.WriteToLog($"Adding '{path}' to PATH environment variable");
                Environment.SetEnvironmentVariable("PATH", Environment.GetEnvironmentVariable("PATH") + @";" + path,
                    EnvironmentVariableTarget.Process);
            }

            Logger.WriteToLog("Starting LibreOffice");

            string pipeName = Guid.NewGuid().ToString().Replace("-", string.Empty);

            Process process = new()
            {
                StartInfo =
                {
                    // -env:UserInstallation=file:///{_userFolder}
                    Arguments = $"-invisible -nofirststartwizard -minimized -nologo -nolockcheck --accept=pipe,name={pipeName};urp;StarOffice.ComponentContext",
                    FileName = installPath + @"\soffice.exe",
                    CreateNoWindow = true
                }
            };

            if (!process.Start())
            {
                throw new InvalidProgramException("Could not start LibreOffice");
            }

            _libreOfficeProcess = process;

            Logger.WriteToLog($"LibreOffice started with process id {process.Id}");

            OpenLibreOfficePipe(pipeName);
        }

        #endregion StartLibreOffice

        #region OpenLibreOfficePipe

        /// <summary>
        /// Opens a pipe to LibreOffice
        /// </summary>
        /// <param name="pipeName"> </param>
        private void OpenLibreOfficePipe(string pipeName)
        {
            XComponentContext localContext = Bootstrap.defaultBootstrap_InitialComponentContext();
            XMultiComponentFactory localServiceManager = localContext.getServiceManager();
            XUnoUrlResolver urlResolver = (XUnoUrlResolver)localServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext);
            XComponentContext remoteContext;

            int i = 0;

            Logger.WriteToLog($"Connecting to LibreOffice with pipe '{pipeName}'");

            while (true)
            {
                try
                {
                    remoteContext = (XComponentContext)urlResolver.resolve($"uno:pipe,name={pipeName};urp;StarOffice.ComponentContext");
                    Logger.WriteToLog("Connected to LibreOffice");
                    break;
                }
                catch (Exception exception)
                {
                    if (i == 20 || !exception.Message.Contains("couldn't connect to pipe"))
                    {
                        throw;
                    }

                    Thread.Sleep(100);
                    i++;
                }
            }

            // ReSharper disable once SuspiciousTypeConversion.Global
            XMultiServiceFactory remoteFactory = (XMultiServiceFactory)remoteContext.getServiceManager();
            _componentLoader = (XComponentLoader)remoteFactory.createInstance("com.sun.star.frame.Desktop");
        }

        #endregion OpenLibreOfficePipe

        #region StopLibreOffice

        /// <summary>
        /// Stops LibreOffice
        /// </summary>
        private void StopLibreOffice()
        {
            if (IsLibreOfficeRunning)
            {
                Logger.WriteToLog($"LibreOffice did not shutdown gracefully in 2 seconds ... killing it on process id {_libreOfficeProcess.Id}");
                _libreOfficeProcess.Kill();
                Logger.WriteToLog("Word process killed");
            }
            else
            {
                Logger.WriteToLog("LibreOffice stopped");
            }

            _libreOfficeProcess = null;
        }

        #endregion StopLibreOffice

        #region ConvertToUrl

        /// <summary>
        /// Convert the give file path to the format LibreOffice needs
        /// </summary>
        /// <param name="file"> </param>
        /// <returns> </returns>
        private static string ConvertToUrl(string file)
        {
            return $"file:///{file.Replace(@"\", "/")}";
        }

        #endregion ConvertToUrl

        #region Convert

        /// <summary>
        /// Converts the given <paramref name="inputFile"/> to PDF format and saves it as <paramref name="outputFile"/>
        /// </summary>
        /// <param name="inputFile"> The input file </param>
        /// <param name="outputFile"> The output file </param>
        public void Convert(string inputFile, string outputFile)
        {
            if (GetFilterType(Path.GetExtension(inputFile)) == null)
            {
                throw new InvalidProgramException($"Unknown file type '{Path.GetFileName(inputFile)}' for LibreOffice");
            }

            StartLibreOffice();

            XComponent component = InitDocument(_componentLoader, ConvertToUrl(inputFile), "_blank");

            // Save/export the document
            // http://herbertniemeyerblog.blogspot.com/2011/11/have-to-start-somewhere.html https://forum.openoffice.org/en/forum/viewtopic.php?t=73098

            ExportToPdf(component, inputFile, outputFile);

            CloseDocument(component);
        }

        #endregion Convert

        #region InitDocument

        /// <summary>
        /// Creates a new document in LibreOffice and opens the given <paramref name="inputFile"/>
        /// </summary>
        /// <param name="aLoader"> </param>
        /// <param name="inputFile"> </param>
        /// <param name="target"> </param>
        /// <returns> </returns>
        private static XComponent InitDocument(XComponentLoader aLoader, string inputFile, string target)
        {
            Logger.WriteToLog($"Loading document '{inputFile}'");

            PropertyValue[] openProps = new PropertyValue[2];
            openProps[0] = new PropertyValue { Name = "Hidden", Value = new Any(true) };
            openProps[1] = new PropertyValue { Name = "ReadOnly", Value = new Any(true) };

            XComponent xComponent = aLoader.loadComponentFromURL(
                inputFile, target, 0,
                openProps);

            Logger.WriteToLog("Document loaded");

            return xComponent;
        }

        #endregion InitDocument

        #region ExportToPdf

        /// <summary>
        /// Exports the loaded document to PDF format
        /// </summary>
        /// <param name="component"> </param>
        /// <param name="inputFile"> </param>
        /// <param name="outputFile"> </param>
        private static void ExportToPdf(XComponent component, string inputFile, string outputFile)
        {
            Logger.WriteToLog($"Exporting document to PDF file '{outputFile}'");

            PropertyValue[] propertyValues = new PropertyValue[3];
            PropertyValue[] filterData = new PropertyValue[5];

            filterData[0] = new PropertyValue
            {
                Name = "UseLosslessCompression",
                Value = new Any(false)
            };

            filterData[1] = new PropertyValue
            {
                Name = "Quality",
                Value = new Any(90)
            };

            filterData[2] = new PropertyValue
            {
                Name = "ReduceImageResolution",
                Value = new Any(true)
            };

            filterData[3] = new PropertyValue
            {
                Name = "MaxImageResolution",
                Value = new Any(300)
            };

            filterData[4] = new PropertyValue
            {
                Name = "ExportBookmarks",
                Value = new Any(false)
            };

            // Setting the filter name
            propertyValues[0] = new PropertyValue
            {
                Name = "FilterName",
                Value = new Any(GetFilterType(inputFile))
            };

            // Setting the flag for overwriting
            propertyValues[1] = new PropertyValue { Name = "Overwrite", Value = new Any(true) };

            PolymorphicType polymorphicType = PolymorphicType.GetType(typeof(PropertyValue[]), "unoidl.com.sun.star.beans.PropertyValue[]");

            propertyValues[2] = new PropertyValue { Name = "FilterData", Value = new Any(polymorphicType, filterData) };

            // ReSharper disable once SuspiciousTypeConversion.Global
            ((XStorable)component).storeToURL(ConvertToUrl(outputFile), propertyValues);

            Logger.WriteToLog("Document exported to PDF");
        }

        #endregion ExportToPdf

        #region CloseDocument

        /// <summary>
        /// Closes the document and frees any used resources
        /// </summary>
        private static void CloseDocument(XComponent component)
        {
            Logger.WriteToLog("Closing document");
            XCloseable closeable = (XCloseable)component;
            closeable?.close(false);
            Logger.WriteToLog("Document closed");
        }

        #endregion CloseDocument

        #region GetFilterType

        /// <summary>
        /// Returns the filter that is needed to convert the given <paramref name="fileName"/>, <c>
        /// null </c> is returned when the file cannot be converted
        /// </summary>
        /// <param name="fileName"> The file to check </param>
        /// <returns> </returns>
        private static string GetFilterType(string fileName)
        {
            string extension = Path.GetExtension(fileName);
            extension = extension?.ToUpperInvariant();

            return extension switch
            {
                ".DOC" or ".DOT" or ".DOCM" or ".DOCX" or ".DOTM" or ".ODT" or ".RTF" or ".MHT" or ".WPS" or ".WRI" => "writer_pdf_Export",
                ".XLS" or ".XLT" or ".XLW" or ".XLSB" or ".XLSM" or ".XLSX" or ".XLTM" or ".XLTX" => "calc_pdf_Export",
                ".POT" or ".PPT" or ".PPS" or ".POTM" or ".POTX" or ".PPSM" or ".PPSX" or ".PPTM" or ".PPTX" or ".ODP" => "impress_pdf_Export",
                _ => null,
            };
        }

        #endregion GetFilterType

        #region Dispose

        /// <summary>
        /// Disposes the running <see cref="_libreOfficeProcess"/>
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            StopLibreOffice();
        }

        #endregion Dispose
    }
}