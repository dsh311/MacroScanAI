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

using System.Runtime.InteropServices;

namespace MacroScanAI.Utils
{
    internal static class ConsoleHelperUtil
    {
        // The maximum number of characters in a console title.
        private const int MAX_CONSOLE_TITLE = 1024;

        // Win32 API declarations for managing the console.
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AttachConsole(int dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetConsoleWindow();


        [DllImport("kernel32.dll")]
        private static extern uint GetConsoleProcessList(uint[] processList, uint processCount);


        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern int GetConsoleTitle(IntPtr lpConsoleTitle, int nSize);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        // The special process ID used to attach to the parent's console.
        private const int ATTACH_PARENT_PROCESS = -1;
        // Window show command constants.
        private const int SW_SHOW = 5;
        private const int SW_HIDE = 0;

        public static void HideConsoleIfOwned()
        {
            IntPtr hWnd = GetConsoleWindow();
            if (hWnd == IntPtr.Zero)
            {
                return; // There is no console
            }

            // Find out how many processes share this console
            uint[] processIds = new uint[4];
            uint count = GetConsoleProcessList(processIds, (uint)processIds.Length);

            // If only THIS process is attached, then we created the console so it is safe to hide
            if (count == 1)
            {
                ShowWindow(hWnd, SW_HIDE);
            }
            else
            {
                // Console is shared with a parent (launch from cmd, PowerShell, etc)
                // Do NOT hide it.
            }
        }

        public static void ShowConsole()
        {
            // Check if a console window is already attached.
            if (GetConsoleWindow() != IntPtr.Zero)
            {
                // If we have a console, make sure it's visible.
                ShowWindow(GetConsoleWindow(), SW_SHOW);
                return;
            }

            // Try to attach to the console of the parent process.
            // This is useful if the app is started from a command prompt.
            bool attached = AttachConsole(ATTACH_PARENT_PROCESS);

            if (!attached)
            {
                // If attachment fails, allocate a new console for this process.
                AllocConsole();
                Console.WriteLine("New console allocated.");
            }
        }

        public static void AllocateConsoleCleanly()
        {
            if (GetConsoleWindow() == IntPtr.Zero)
            {
                AllocConsole();
            }
        }
    }
}
