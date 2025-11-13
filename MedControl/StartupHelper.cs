using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

namespace MedControl
{
    internal static class StartupHelper
    {
        private const string RunKeyPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
        private const string ValueName = "MedControl";

        // Garante que a aplicação inicia com o Windows (escopo usuário)
        internal static void EnsureAutoStart()
        {
            try
            {
                using var runKey = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true) ??
                                   Registry.CurrentUser.CreateSubKey(RunKeyPath, writable: true);
                if (runKey == null) return;

                var exePath = Application.ExecutablePath;
                var current = runKey.GetValue(ValueName) as string;
                if (string.IsNullOrWhiteSpace(current) || !StringEqualsIgnoreCase(current, exePath))
                {
                    runKey.SetValue(ValueName, exePath);
                }
            }
            catch { }
        }

        private static bool StringEqualsIgnoreCase(string? a, string? b) =>
            string.Equals(a?.Trim(), b?.Trim(), StringComparison.OrdinalIgnoreCase);

        // Permite remover se necessário futuramente
        internal static void RemoveAutoStart()
        {
            try
            {
                using var runKey = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true);
                runKey?.DeleteValue(ValueName, false);
            }
            catch { }
        }
    }
}