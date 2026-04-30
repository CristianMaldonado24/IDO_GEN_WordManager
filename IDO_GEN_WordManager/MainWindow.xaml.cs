using IDO_GEN_WordManager.Models;
using IDO_GEN_WordManager.ViewModels;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace IDO_GEN_WordManager;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        TrySetWindowIcon();
        Loaded += (s, e) =>
        {
            HeadingsGrid.SelectionChanged += HeadingsGrid_SelectionChanged;
            if (DataContext is MainViewModel vm)
                vm.ClearSelectionRequested += (_, _) => HeadingsGrid.UnselectAll();
        };
    }

    private void TrySetWindowIcon()
    {
        try
        {
            var visual = new DrawingVisual();
            using (var ctx = visual.RenderOpen())
            {
                // Fondo azul redondeado
                ctx.DrawRoundedRectangle(
                    new SolidColorBrush(Color.FromRgb(26, 86, 219)),
                    null,
                    new Rect(0, 0, 32, 32), 6, 6);

                // Cuerpo de la pluma
                var fill  = Brushes.White;
                var stroke = new Pen(fill, 0) { LineJoin = PenLineJoin.Round };
                var body   = Geometry.Parse("M 24,4 C 28,0 32,4 28,8 L 13,23 L 8,24 L 9,19 Z");
                ctx.DrawGeometry(fill, stroke, body);

                // Punta de escritura
                var tip = Geometry.Parse("M 8,24 L 5,27 L 7,27 L 9,25 Z");
                ctx.DrawGeometry(fill, stroke, tip);

                // Línea de tinta
                var ink = new Pen(new SolidColorBrush(Color.FromRgb(26, 86, 219)), 1.2)
                    { StartLineCap = PenLineCap.Round, EndLineCap = PenLineCap.Round };
                ctx.DrawLine(ink, new Point(12, 20), new Point(20, 12));
            }

            var bmp = new RenderTargetBitmap(32, 32, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(visual);
            Icon = bmp;
        }
        catch { }
    }

    private void HeadingsGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Limpiar selecciones previas
        foreach (var item in e.RemovedItems.Cast<DocumentHeading>())
        {
            item.IsSelected = false;
        }

        // Marcar nuevas selecciones
        foreach (var item in e.AddedItems.Cast<DocumentHeading>())
        {
            item.IsSelected = true;
        }
    }
}