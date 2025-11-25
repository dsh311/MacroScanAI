/*
 * Copyright (C) 2025 David S. Shelley <davidsmithshelley@gmail.com>
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License 
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 */

using System.Diagnostics;

namespace MacroScanAI.Cli
{
    internal class Program
    {
        static int Main(string[] args)
        {
            try
            {
                string exeDir = AppContext.BaseDirectory;
                string targetExe = Path.Combine(exeDir, "MacroScanAI.exe");

                if (!File.Exists(targetExe))
                {
                    Console.Error.WriteLine($"Error: Could not find MacroScanAI.exe at: {targetExe}");
                    return -1;
                }

                string argString = string.Join(" ", args.Select(a => QuoteIfNeeded(a)));

                var psi = new ProcessStartInfo
                {
                    FileName = targetExe,
                    Arguments = argString,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                var process = new Process
                {
                    StartInfo = psi,
                    EnableRaisingEvents = true
                };

                process.OutputDataReceived += (s, e) =>
                {
                    if (e.Data != null)
                        Console.WriteLine(e.Data);
                };

                process.ErrorDataReceived += (s, e) =>
                {
                    if (e.Data != null)
                        Console.Error.WriteLine(e.Data);
                };

                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                process.WaitForExit();
                return process.ExitCode;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Fatal error: " + ex.Message);
                return -1;
            }
        }

        private static string QuoteIfNeeded(string arg)
        {
            if (arg.Contains(' ') || arg.Contains('"'))
                return "\"" + arg.Replace("\"", "\\\"") + "\"";
            return arg;
        }
    }
}