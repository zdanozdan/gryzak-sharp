using System;
using System.Collections.Generic;
using System.Linq;
using System.Printing;
using System.Runtime.InteropServices;
using System.Text;

namespace Gryzak.Services
{
    /// <summary>
    /// Wysyłka surowych bajtów (ZPL) na drukarkę Windows w trybie RAW.
    /// </summary>
    public static class RawPrinterHelper
    {
        public static IReadOnlyList<string> GetInstalledPrinterNames()
        {
            try
            {
                using var server = new LocalPrintServer();
                var queues = server.GetPrintQueues(new[]
                {
                    EnumeratedPrintQueueTypes.Local,
                    EnumeratedPrintQueueTypes.Connections
                });

                return queues
                    .Select(q => q.FullName)
                    .Where(n => !string.IsNullOrWhiteSpace(n))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch
            {
                return Array.Empty<string>();
            }
        }

        public static string? PickPreferredPrinter(string? preferredName, IReadOnlyList<string>? printers = null)
        {
            printers ??= GetInstalledPrinterNames();
            if (printers.Count == 0)
            {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(preferredName))
            {
                var exact = printers.FirstOrDefault(p =>
                    p.Equals(preferredName.Trim(), StringComparison.OrdinalIgnoreCase));
                if (exact != null)
                {
                    return exact;
                }
            }

            var zebra = printers.FirstOrDefault(p =>
                p.Contains("zebra", StringComparison.OrdinalIgnoreCase)
                || p.Contains("zdesigner", StringComparison.OrdinalIgnoreCase)
                || p.Contains("gk420", StringComparison.OrdinalIgnoreCase)
                || p.Contains("zd4", StringComparison.OrdinalIgnoreCase)
                || p.Contains("zt4", StringComparison.OrdinalIgnoreCase));
            if (zebra != null)
            {
                return zebra;
            }

            try
            {
                using var server = new LocalPrintServer();
                var defaultName = server.DefaultPrintQueue?.FullName;
                if (!string.IsNullOrWhiteSpace(defaultName))
                {
                    var match = printers.FirstOrDefault(p =>
                        p.Equals(defaultName, StringComparison.OrdinalIgnoreCase));
                    if (match != null)
                    {
                        return match;
                    }
                }
            }
            catch
            {
                // ignore
            }

            return printers[0];
        }

        public static bool SendBytes(string printerName, byte[] data, string documentName, out string? error)
        {
            error = null;
            if (string.IsNullOrWhiteSpace(printerName))
            {
                error = "Nie wybrano drukarki.";
                return false;
            }

            if (data == null || data.Length == 0)
            {
                error = "Brak danych do druku.";
                return false;
            }

            var di = new DOCINFOA
            {
                pDocName = string.IsNullOrWhiteSpace(documentName) ? "Gryzak GLS" : documentName,
                pDataType = "RAW"
            };

            var hPrinter = IntPtr.Zero;
            try
            {
                if (!OpenPrinter(printerName.Trim(), ref hPrinter, IntPtr.Zero))
                {
                    error = $"Nie udało się otworzyć drukarki „{printerName}” (Win32 {Marshal.GetLastWin32Error()}).";
                    return false;
                }

                if (!StartDocPrinter(hPrinter, 1, di))
                {
                    error = $"StartDocPrinter nie powiódł się (Win32 {Marshal.GetLastWin32Error()}).";
                    return false;
                }

                try
                {
                    if (!StartPagePrinter(hPrinter))
                    {
                        error = $"StartPagePrinter nie powiódł się (Win32 {Marshal.GetLastWin32Error()}).";
                        return false;
                    }

                    try
                    {
                        var handle = GCHandle.Alloc(data, GCHandleType.Pinned);
                        try
                        {
                            var written = 0;
                            if (!WritePrinter(hPrinter, handle.AddrOfPinnedObject(), data.Length, ref written)
                                || written != data.Length)
                            {
                                error = $"WritePrinter nie powiódł się (zapisano {written}/{data.Length}, Win32 {Marshal.GetLastWin32Error()}).";
                                return false;
                            }
                        }
                        finally
                        {
                            handle.Free();
                        }
                    }
                    finally
                    {
                        EndPagePrinter(hPrinter);
                    }
                }
                finally
                {
                    EndDocPrinter(hPrinter);
                }

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            finally
            {
                if (hPrinter != IntPtr.Zero)
                {
                    ClosePrinter(hPrinter);
                }
            }
        }

        public static bool SendUtf8Text(string printerName, string text, string documentName, out string? error)
        {
            var bytes = Encoding.UTF8.GetBytes(text ?? "");
            return SendBytes(printerName, bytes, documentName, out error);
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        private class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)]
            public string? pDocName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string? pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)]
            public string? pDataType;
        }

        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern bool OpenPrinter(string szPrinter, ref IntPtr hPrinter, IntPtr pd);

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true)]
        private static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern bool StartDocPrinter(IntPtr hPrinter, int level, [In] DOCINFOA di);

        [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true)]
        private static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true)]
        private static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true)]
        private static extern bool EndPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true)]
        private static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, int dwCount, ref int dwWritten);
    }
}
