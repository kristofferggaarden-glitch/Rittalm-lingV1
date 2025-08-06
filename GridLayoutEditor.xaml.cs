using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace RittalMåling
{
    public partial class GridLayoutEditor : Window
    {
        private readonly int rows;
        private readonly int cols;
        private readonly bool[,] layout;

        public bool[,] Result { get; private set; }

        public GridLayoutEditor(int rows, int cols, bool[,] existingLayout)
        {
            InitializeComponent();
            this.rows = rows;
            this.cols = cols;
            // Initialize layout, copying existing or creating new
            layout = new bool[rows, cols];
            if (existingLayout != null && existingLayout.GetLength(0) == rows && existingLayout.GetLength(1) == cols)
            {
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        layout[r, c] = existingLayout[r, c];
            }
            else
            {
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        layout[r, c] = true; // Default: all cells active
            }
            BuildGrid();
        }

        private void BuildGrid()
        {
            LayoutGrid.Children.Clear();
            LayoutGrid.RowDefinitions.Clear();
            LayoutGrid.ColumnDefinitions.Clear();

            for (int r = 0; r < rows; r++)
                LayoutGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            for (int c = 0; c < cols; c++)
                LayoutGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    var btn = new Button
                    {
                        Style = (Style)FindResource("GridCellStyle"),
                        Background = layout[r, c] ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 178, 148)) : new SolidColorBrush(System.Windows.Media.Color.FromRgb(74, 90, 91)),
                        Tag = (r, c)
                    };
                    btn.Click += Cell_Click;
                    Grid.SetRow(btn, r);
                    Grid.SetColumn(btn, c);
                    LayoutGrid.Children.Add(btn);
                }
            }
        }

        private void Cell_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn?.Tag is (int r, int c))
            {
                layout[r, c] = !layout[r, c];
                btn.Background = layout[r, c] ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 178, 148)) : new SolidColorBrush(System.Windows.Media.Color.FromRgb(74, 90, 91));
            }
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            Result = layout;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}