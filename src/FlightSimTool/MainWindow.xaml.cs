using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using FlightSimTool.Core;
using Microsoft.Win32;

namespace FlightSimTool
{
    /// <summary>
    /// Represents a window item displayed in the list view.
    /// contains helper properties for data binding.
    /// </summary>
    public class WindowDisplayItem
    {
        public bool IsSelected { get; set; }
        public IntPtr Handle { get; set; }
        public string Title { get; set; } = "";
        public string ClassName { get; set; } = "";
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string BoundsString => $"{X},{Y}  {Width}x{Height}";
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml.
    /// Handles the main UI flow for capturing, editing, and applying window layouts.
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<OverlayWindow> _overlays = new List<OverlayWindow>();
        private string _currentLayoutPath = "WindowLayout.json";
        
        /// <summary>
        /// Collection of layout entries bound to the Editor DataGrid.
        /// </summary>
        public System.Collections.ObjectModel.ObservableCollection<LayoutEntry> EditorEntries { get; set; } = new System.Collections.ObjectModel.ObservableCollection<LayoutEntry>();

        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this; // Allow binding to EditorEntries
            
            // Hook up theme integration
            this.SourceInitialized += (s, e) => UpdateTitleBarTheme();
            Microsoft.Win32.SystemEvents.UserPreferenceChanged += (s, e) => 
            {
                if (e.Category == Microsoft.Win32.UserPreferenceCategory.General)
                {
                    Dispatcher.Invoke(() => UpdateTitleBarTheme());
                }
            };
            
            RefreshList();
            TxtProfilePath.Text = System.IO.Path.GetFullPath(_currentLayoutPath);
            
            // Initial load of editor if file exists
            LoadEditorEntries(_currentLayoutPath);
        }

        /// <summary>
        /// Applies the Windows immersive dark mode attribute to the window title bar if the system theme is Dark.
        /// </summary>
        private void UpdateTitleBarTheme()
        {
            var theme = ThemeManager.GetCurrentSystemTheme();
            var handle = new System.Windows.Interop.WindowInteropHelper(this).Handle;
            FlightSimTool.Core.NativeMethods.UseImmersiveDarkMode(handle, theme == ThemeManager.Theme.Dark);
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            RefreshList();
        }

        /// <summary>
        /// Refreshes the list of currently open windows in the capture tab.
        /// </summary>
        private void RefreshList()
        {
            var windows = WindowHelper.GetOpenWindows(includeHidden: false);
            var displayItems = windows.Select(w => new WindowDisplayItem
            {
                IsSelected = false,
                Handle = w.Handle,
                Title = w.Title,
                ClassName = w.ClassName,
                X = w.X,
                Y = w.Y,
                Width = w.Width,
                Height = w.Height
            }).OrderBy(w => w.Title).ToList();

            ListWindows.ItemsSource = displayItems;
            TxtStatus.Text = $"Found {displayItems.Count} open windows.";
        }

        // Renamed from BtnSave_Click. Now creates a new layout from selected windows.
        private void BtnCreateLayout_Click(object sender, RoutedEventArgs e)
        {
            var items = ListWindows.ItemsSource as List<WindowDisplayItem>;
            if (items == null) return;

            var selected = items.Where(i => i.IsSelected).ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show("No windows selected to create layout from.", "Create Layout", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Create new layout entries with default "safe" values (Follow=false, Border=0)
            var entries = selected.Select(s => new LayoutEntry
            {
                TitleLike = Core.WindowHelper.FindWindow(s.Title) != IntPtr.Zero ? s.Title : s.Title, 
                X = s.X,
                Y = s.Y,
                Width = s.Width,
                Height = s.Height,
                Follow = false, // Default to false per user request
                BorderThickness = 0, // Default to 0 per user request
                BorderCover = 0,
                BorderTopExtra = 0,
                FullScreen = false
            }).ToList();

            var dlg = new SaveFileDialog
            {
                FileName = "WindowLayout",
                DefaultExt = ".json",
                Filter = "JSON Files|*.json"
            };

            if (dlg.ShowDialog() == true)
            {
                _currentLayoutPath = dlg.FileName;
                WindowHelper.SaveLayout(entries, _currentLayoutPath);
                TxtProfilePath.Text = _currentLayoutPath;
                
                // Switch to editor and load
                LoadEditorEntries(_currentLayoutPath);
                TabEditor.IsSelected = true;
                
                TxtStatus.Text = $"Created new layout with {entries.Count} entries.";
            }
        }

        private void BtnLoad_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                DefaultExt = ".json",
                Filter = "JSON Files|*.json"
            };

            if (dlg.ShowDialog() == true)
            {
                _currentLayoutPath = dlg.FileName;
                TxtProfilePath.Text = _currentLayoutPath;
                LoadEditorEntries(_currentLayoutPath);
                
                // If we are in live capture, maybe stay there? User might want to edit. Let's switch to editor.
                TabEditor.IsSelected = true;
                
                TxtStatus.Text = "Layout loaded into Editor.";
            }
        }

        private void LoadEditorEntries(string path)
        {
            EditorEntries.Clear();
            var loaded = WindowHelper.LoadLayout(path);
            foreach (var entry in loaded)
            {
                EditorEntries.Add(entry);
            }
            GridEditor.ItemsSource = EditorEntries; // Re-bind to ensure updates
        }

        private void BtnSaveEditor_Click(object sender, RoutedEventArgs e)
        {
            try 
            {
                var entries = EditorEntries.ToList();
                WindowHelper.SaveLayout(entries, _currentLayoutPath);
                TxtStatus.Text = $"Saved changes to {_currentLayoutPath}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving layout: {ex.Message}");
            }
        }

        private void BtnRestore_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Use the entries currently in the Editor (memory) instead of file
                var entries = EditorEntries.ToList();
                
                int restored = 0;
                foreach (var entry in entries)
                {
                    // Find window by TitleLike
                    IntPtr handle = WindowHelper.FindWindow(entry.TitleLike);
                    if (handle != IntPtr.Zero)
                    {
                        // Use updated helper which handles StripTitleBar and FullScreen inside
                        WindowHelper.RestoreWindow(handle, entry);
                        restored++;
                    }
                }
                TxtStatus.Text = $"Restored {restored} / {entries.Count} windows.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error restoring layout: {ex.Message}");
            }
        }

        private void BtnOverlays_Click(object sender, RoutedEventArgs e)
        {
            StopOverlays();
            try
            {
                // Use entries from Editor
                var entries = EditorEntries.ToList();
                
                int count = 0;
                foreach (var entry in entries)
                {
                    IntPtr handle = WindowHelper.FindWindow(entry.TitleLike);
                    if (handle != IntPtr.Zero)
                    {
                        var overlay = new OverlayWindow(handle, entry);
                        overlay.Show();
                        _overlays.Add(overlay);
                        count++;
                    }
                }
                TxtStatus.Text = $"Started {count} overlays.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting overlays: {ex.Message}");
            }
        }

        private void BtnStopOverlays_Click(object sender, RoutedEventArgs e)
        {
            StopOverlays();
            TxtStatus.Text = "Overlays stopped.";
        }

        private void StopOverlays()
        {
            foreach (var ov in _overlays)
            {
                ov.Stop();
            }
            _overlays.Clear();
        }

        protected override void OnClosed(EventArgs e)
        {
            StopOverlays();
            base.OnClosed(e);
        }
    }
}
