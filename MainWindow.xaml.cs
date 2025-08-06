using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace RittalMåling
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private int _sections = 5;
        private int _rows = 7;
        private int _cols = 4;
        private string _selectedExcelFile;
        private string _lastMeasurement;
        private string _excelDisplayText;
        private Button _startPoint;
        private Button _endPoint;
        private List<Cell> _lastPath;
        private Excel.Application excelApp;
        private Excel.Workbook workbook;
        private Excel.Worksheet worksheet;
        private readonly Dictionary<(int globalRow, int globalCol), Cell> allCells = new Dictionary<(int, int), Cell>();
        private bool[,] customGridLayout;
        private int _currentExcelRow = 1; // Start with row 1 (B1, C1)

        public int Sections
        {
            get => _sections;
            set { _sections = value; OnPropertyChanged(nameof(Sections)); }
        }

        public int Rows
        {
            get => _rows;
            set { _rows = value; OnPropertyChanged(nameof(Rows)); }
        }

        public int Cols
        {
            get => _cols;
            set { _cols = value; OnPropertyChanged(nameof(Cols)); }
        }

        public string SelectedExcelFile
        {
            get => _selectedExcelFile;
            set { _selectedExcelFile = value; OnPropertyChanged(nameof(SelectedExcelFile)); }
        }

        public string LastMeasurement
        {
            get => _lastMeasurement;
            set { _lastMeasurement = value; OnPropertyChanged(nameof(LastMeasurement)); }
        }

        public string ExcelDisplayText
        {
            get => _excelDisplayText;
            set { _excelDisplayText = value; OnPropertyChanged(nameof(ExcelDisplayText)); }
        }

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            UpdateExcelDisplayText();
            BuildAllSections();
        }

        private void UpdateExcelDisplayText()
        {
            if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
            {
                ExcelDisplayText = "";
                return;
            }
            try
            {
                string cellB = (worksheet.Cells[_currentExcelRow, 2] as Excel.Range)?.Value?.ToString() ?? "";
                string cellC = (worksheet.Cells[_currentExcelRow, 3] as Excel.Range)?.Value?.ToString() ?? "";
                ExcelDisplayText = string.IsNullOrEmpty(cellB) && string.IsNullOrEmpty(cellC) ? "" : $"{cellB} - {cellC}".Trim();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading Excel data for row {_currentExcelRow}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ExcelDisplayText = "";
            }
        }

        private bool InitializeExcel(string filePath)
        {
            try
            {
                CleanupExcel();
                excelApp = new Excel.Application { Visible = true };
                if (File.Exists(filePath))
                {
                    workbook = excelApp.Workbooks.Open(filePath);
                }
                else
                {
                    workbook = excelApp.Workbooks.Add();
                    workbook.SaveAs(filePath);
                }
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to initialize Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            CleanupExcel();
        }

        private void CleanupExcel()
        {
            try
            {
                if (workbook != null)
                {
                    workbook.Save();
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error cleaning up Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                worksheet = null;
                workbook = null;
                excelApp = null;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private void Rebuild_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(SectionBox.Text, out int s) || s <= 0 || s > 10)
            {
                MessageBox.Show("Please enter a valid number of sections (1-10).", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!int.TryParse(RowBox.Text, out int r) || r <= 0 || r > 20)
            {
                MessageBox.Show("Please enter a valid number of rows (1-20).", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!int.TryParse(ColBox.Text, out int c) || c <= 0 || c > 10)
            {
                MessageBox.Show("Please enter a valid number of columns (1-10).", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Sections = s;
            Rows = r;
            Cols = c;
            customGridLayout = null;
            _currentExcelRow = 1; // Reset Excel row
            BuildAllSections();
            UpdateExcelDisplayText();
            if (_lastPath != null && IsValidPath(_lastPath))
            {
                HighlightPath(_lastPath);
            }
        }

        private void SelectExcelFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                Title = "Select Excel File for Measurements"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                SelectedExcelFile = openFileDialog.FileName;
                if (InitializeExcel(SelectedExcelFile))
                {
                    MessageBox.Show($"Excel file '{SelectedExcelFile}' selected and opened.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    _currentExcelRow = FindFirstEmptyRow(); // Set to first row where A is empty
                    UpdateExcelDisplayText();
                    BuildAllSections(); // Rebuild to clear any previous state
                }
                else
                {
                    SelectedExcelFile = null;
                    ExcelDisplayText = "";
                }
            }
        }

        private int FindFirstEmptyRow()
        {
            if (worksheet == null) return 1;
            int row = 1;
            while ((worksheet.Cells[row, 1] as Excel.Range)?.Value != null)
            {
                row++;
            }
            return row;
        }

        private void DeleteLast_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
                {
                    MessageBox.Show("Please select an Excel file first.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                int lastRow = 1;
                int row = 1;
                while ((worksheet.Cells[row, 1] as Excel.Range)?.Value != null)
                {
                    lastRow = row;
                    row++;
                }
                if (lastRow >= 1)
                {
                    ((Excel.Range)worksheet.Cells[lastRow, 1]).Clear();
                    if (lastRow == 1)
                    {
                        LastMeasurement = "";
                        _lastPath = null;
                        _currentExcelRow = 1; // Reset to row 1
                        BuildAllSections();
                        UpdateExcelDisplayText();
                    }
                    else
                    {
                        _currentExcelRow = lastRow; // Set to the last row with a measurement
                        object lastValueObj = (worksheet.Cells[_currentExcelRow, 1] as Excel.Range)?.Value;
                        double? lastValue = lastValueObj != null ? Convert.ToDouble(lastValueObj) : (double?)null;
                        LastMeasurement = lastValue.HasValue ? $"Last measurement: {lastValue:F2} mm" : "";
                        UpdateExcelDisplayText();
                    }
                }
                else
                {
                    MessageBox.Show("No measurements to delete.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    LastMeasurement = "";
                    _lastPath = null;
                    _currentExcelRow = 1; // Reset to row 1
                    BuildAllSections();
                    UpdateExcelDisplayText();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error deleting last measurement: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void EditGridLayout_Click(object sender, RoutedEventArgs e)
        {
            var editor = new GridLayoutEditor(Rows, Cols * Sections, customGridLayout)
            {
                Owner = this
            };
            if (editor.ShowDialog() == true)
            {
                customGridLayout = editor.Result;
                _currentExcelRow = 1; // Reset Excel row
                BuildAllSections();
                UpdateExcelDisplayText();
                if (_lastPath != null && IsValidPath(_lastPath))
                {
                    HighlightPath(_lastPath);
                }
            }
        }

        private bool IsValidPath(List<Cell> path)
        {
            if (path == null || path.Count == 0) return false;
            foreach (var cell in path)
            {
                if (!allCells.ContainsKey((cell.Row, cell.Col)))
                    return false;
            }
            return true;
        }

        private void BuildAllSections()
        {
            MainPanel.Children.Clear();
            allCells.Clear();
            _startPoint = null;
            _endPoint = null;

            for (int s = 0; s < Sections; s++)
            {
                var sectionPanel = new StackPanel
                {
                    Margin = new Thickness(15),
                    Orientation = Orientation.Vertical,
                    Background = Brushes.Transparent
                };
                var grid = new Grid
                {
                    Background = Brushes.Transparent,
                    Margin = new Thickness(0, 10, 0, 10)
                };
                sectionPanel.Children.Add(grid);
                grid.RowDefinitions.Clear();
                grid.ColumnDefinitions.Clear();
                for (int r = 0; r < Rows; r++)
                    grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                for (int c = 0; c < Cols; c++)
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                for (int r = 0; r < Rows; r++)
                {
                    for (int c = 0; c < Cols; c++)
                    {
                        int globalCol = s * Cols + c;
                        if (customGridLayout == null || customGridLayout[r, globalCol])
                        {
                            AddCell(grid, r, c, s);
                        }
                    }
                }
                MainPanel.Children.Add(sectionPanel);
            }
        }

        private void AddCell(Grid grid, int localRow, int localCol, int sectionIndex)
        {
            var btn = new Button
            {
                Style = (Style)FindResource("GridButtonStyle"),
                ToolTip = $"Cell ({localRow}, {localCol}) in section {sectionIndex + 1}",
                Content = ""
            };
            btn.Click += Cell_Click;
            Grid.SetRow(btn, localRow);
            Grid.SetColumn(btn, localCol);
            grid.Children.Add(btn);
            int globalRow = localRow;
            int globalCol = sectionIndex * Cols + localCol;
            allCells[(globalRow, globalCol)] = new Cell(globalRow, globalCol, btn);
        }

        private void Cell_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            var cell = allCells.Values.FirstOrDefault(c => c.ButtonRef == btn);
            if (cell == null) return;

            if (_startPoint == null)
            {
                ResetSelection();
                _startPoint = btn;
                btn.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 120, 212));
            }
            else if (_endPoint == null && btn != _startPoint)
            {
                _endPoint = btn;
                btn.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(232, 17, 35));
                ProcessPath();
            }
        }

        private void ProcessPath()
        {
            var startCell = allCells.Values.FirstOrDefault(c => c.ButtonRef == _startPoint);
            var endCell = allCells.Values.FirstOrDefault(c => c.ButtonRef == _endPoint);
            if (startCell == null || endCell == null)
            {
                LastMeasurement = "";
                ResetSelection();
                return;
            }
            var path = PathFinder.FindShortestPath(startCell, endCell, allCells);
            if (path == null)
            {
                LastMeasurement = "";
                ResetSelection();
                return;
            }
            _lastPath = path;
            double totalDistance = PathFinder.CalculateDistance(path);
            HighlightPath(path);
            LastMeasurement = $"Last measurement: {totalDistance:F2} mm";
            LogMeasurementToExcel(totalDistance);
            _currentExcelRow = FindFirstEmptyRow(); // Move to next empty row
            UpdateExcelDisplayText();
            _startPoint = null;
            _endPoint = null;
        }

        private void ResetSelection()
        {
            foreach (var cell in allCells.Values)
            {
                cell.ButtonRef.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(74, 90, 91));
                cell.ButtonRef.Content = "";
            }
            _startPoint = null;
            _endPoint = null;
            _lastPath = null;
        }

        private void HighlightPath(List<Cell> path)
        {
            foreach (var cell in allCells.Values)
            {
                cell.ButtonRef.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(74, 90, 91));
                cell.ButtonRef.Content = "";
            }
            foreach (var cell in path)
            {
                cell.ButtonRef.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 178, 148));
                cell.ButtonRef.Content = "";
            }
            if (path.Count > 0)
            {
                path.First().ButtonRef.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 120, 212));
                path.First().ButtonRef.Content = "";
                path.Last().ButtonRef.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(232, 17, 35));
                path.Last().ButtonRef.Content = "";
            }
        }

        private void LogMeasurementToExcel(double distance)
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
                {
                    MessageBox.Show("Please select an Excel file first.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                worksheet.Cells[_currentExcelRow, 1] = distance; // Log to current row (A1 for _currentExcelRow=1)
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error writing to Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    public static class PathFinder
    {
        public static List<Cell> FindShortestPath(Cell start, Cell end, Dictionary<(int, int), Cell> allCells)
        {
            var dist = new Dictionary<Cell, double>();
            dist[start] = 200;
            var prev = new Dictionary<Cell, Cell>();
            var queue = new PriorityQueue<Cell, double>();
            queue.Enqueue(start, 200);
            foreach (var cell in allCells.Values)
            {
                if (cell != start)
                    dist[cell] = double.PositiveInfinity;
                prev[cell] = null;
            }
            while (queue.Count > 0)
            {
                var u = queue.Dequeue();
                if (u == end)
                {
                    dist[u] = dist[u] + 200; // Add 200mm for end point
                    break;
                }
                foreach (var neighbor in GetNeighbors(u, allCells))
                {
                    double weight = 100; // Each intermediate box = 100mm
                    double alt = dist[u] + weight;
                    if (alt < dist[neighbor])
                    {
                        dist[neighbor] = alt;
                        prev[neighbor] = u;
                        queue.Enqueue(neighbor, alt);
                    }
                }
            }
            if (double.IsInfinity(dist[end])) return null;
            var path = new List<Cell>();
            for (var curr = end; curr != null; curr = prev[curr])
                path.Add(curr);
            path.Reverse();
            return path;
        }

        private static List<Cell> GetNeighbors(Cell cell, Dictionary<(int, int), Cell> allCells)
        {
            var directions = new[] { (-1, 0), (1, 0), (0, -1), (0, 1) };
            var neighbors = new List<Cell>();
            foreach (var (dr, dc) in directions)
            {
                int nr = cell.Row + dr;
                int nc = cell.Col + dc;
                if (allCells.TryGetValue((nr, nc), out var neighbor))
                    neighbors.Add(neighbor);
            }
            return neighbors;
        }

        public static double CalculateDistance(List<Cell> path)
        {
            if (path == null || path.Count == 0) return 0;
            if (path.Count == 1) return 400; // Same cell: 200mm start + 200mm end
            return 200 + 200 + (path.Count - 2) * 100; // Start + end + intermediate
        }
    }

    public class Cell
    {
        public int Row { get; }
        public int Col { get; }
        public Button ButtonRef { get; }
        public Cell(int row, int col, Button button)
        {
            Row = row;
            Col = col;
            ButtonRef = button;
        }
    }

    public class PriorityQueue<TItem, TPriority> where TPriority : IComparable<TPriority>
    {
        private readonly List<(TItem Item, TPriority Priority)> elements = new List<(TItem, TPriority)>();
        public int Count => elements.Count;
        public void Enqueue(TItem item, TPriority priority)
        {
            elements.Add((item, priority));
        }
        public TItem Dequeue()
        {
            int bestIndex = 0;
            for (int i = 1; i < elements.Count; i++)
            {
                if (elements[i].Priority.CompareTo(elements[bestIndex].Priority) < 0)
                    bestIndex = i;
            }
            var result = elements[bestIndex].Item;
            elements.RemoveAt(bestIndex);
            return result;
        }
    }
}