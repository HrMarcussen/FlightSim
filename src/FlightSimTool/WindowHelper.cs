using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using System.IO;

namespace FlightSimTool.Core
{
    /// <summary>
    /// Represents basic information about an open window.
    /// </summary>
    public class WindowInfo
    {
        public IntPtr Handle { get; set; }
        public string Title { get; set; } = "";
        public string ClassName { get; set; } = "";
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }

    /// <summary>
    /// Represents a saved window layout configuration entry.
    /// </summary>
    public class LayoutEntry
    {
        public string TitleLike { get; set; } = "";
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public int BorderThickness { get; set; } = 8;
        public int BorderTopExtra { get; set; } = 0;
        public int BorderCover { get; set; } = 2;
        public int BorderTopCoverExtra { get; set; } = 0;
        public bool StripTitleBar { get; set; } = false;
        public bool Follow { get; set; } = true;
        public bool FullScreen { get; set; } = false;
    }

    /// <summary>
    /// Provides helper methods for enumerating windows, finding them by title, and managing layouts.
    /// </summary>
    public static class WindowHelper
    {
        /// <summary>
        /// Retrieves a list of currently open windows.
        /// </summary>
        /// <param name="includeHidden">If true, includes invisible windows.</param>
        /// <returns>A list of WindowInfo objects.</returns>
        public static List<WindowInfo> GetOpenWindows(bool includeHidden = false)
        {
            var list = new List<WindowInfo>();
            NativeMethods.EnumWindows((h, p) =>
            {
                if (!includeHidden && !NativeMethods.IsWindowVisible(h)) return true;

                int len = NativeMethods.GetWindowTextLength(h);
                if (len <= 0) return true;

                var sb = new StringBuilder(len + 1);
                NativeMethods.GetWindowText(h, sb, sb.Capacity);
                string title = sb.ToString();
                if (string.IsNullOrWhiteSpace(title)) return true;

                var csb = new StringBuilder(256);
                NativeMethods.GetClassName(h, csb, csb.Capacity);
                
                NativeMethods.GetWindowRect(h, out var r);
                if (r.Width <= 0 || r.Height <= 0) return true;

                list.Add(new WindowInfo
                {
                    Handle = h,
                    Title = title,
                    ClassName = csb.ToString(),
                    X = r.Left,
                    Y = r.Top,
                    Width = r.Width,
                    Height = r.Height
                });
                return true;
            }, IntPtr.Zero);
            return list;
        }

        /// <summary>
        /// Saves a list of layout entries to a JSON file.
        /// </summary>
        /// <param name="entries">The entries to save.</param>
        /// <param name="path">The file path.</param>
        public static void SaveLayout(List<LayoutEntry> entries, string path)
        {
            var options = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(entries, options);
            File.WriteAllText(path, json);
        }

        /// <summary>
        /// Loads layout entries from a JSON file.
        /// </summary>
        /// <param name="path">The file path.</param>
        /// <returns>A list of LayoutEntry objects.</returns>
        public static List<LayoutEntry> LoadLayout(string path)
        {
            if (!File.Exists(path)) return new List<LayoutEntry>();
            string json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<List<LayoutEntry>>(json) ?? new List<LayoutEntry>();
        }

        /// <summary>
        /// Finds the first window handle that matches the given title (partial match, case-insensitive).
        /// </summary>
        /// <param name="titleLike">The title substring to search for.</param>
        /// <returns>The window handle (IntPtr.Zero if not found).</returns>
        public static IntPtr FindWindow(string titleLike)
        {
            // Simple linear search for best match
            var all = GetOpenWindows();
            foreach (var w in all)
            {
                if (w.Title.Contains(titleLike, StringComparison.OrdinalIgnoreCase)) return w.Handle;
            }
            return IntPtr.Zero;
        }

        /// <summary>
        /// Restores a window's position and style based on the provided layout entry.
        /// Handles stripping title bars and full-screen maximization.
        /// </summary>
        /// <param name="handle">The window handle.</param>
        /// <param name="entry">The layout configuration.</param>
        public static void RestoreWindow(IntPtr handle, LayoutEntry entry)
        {
            NativeMethods.ShowWindowAsync(handle, NativeMethods.SW_RESTORE);

            if (entry.StripTitleBar || entry.FullScreen)
            {
                // FullScreen implies stripping title bar
                NativeMethods.StripTitleBarKeepBounds(handle, entry.X, entry.Y, entry.Width, entry.Height);
            }

            if (entry.FullScreen)
            {
                // Real Fullscreen: Maximize and remove decoration
                // Often better to set exact bounds to monitor, but SW_MAXIMIZE works if style is popup
                NativeMethods.ShowWindowAsync(handle, NativeMethods.SW_MAXIMIZE);
            }
            else
            {
                NativeMethods.SetWindowPos(handle, NativeMethods.HWND_TOP, entry.X, entry.Y, entry.Width, entry.Height, 
                    NativeMethods.SWP_NOZORDER | NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_FRAMECHANGED);
            }
        }
    }
}
