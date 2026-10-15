using MigraDoc.DocumentObjectModel;
using PdfSharp.Drawing;

namespace ConverterReport.Core.Pdf;

/// <summary>Цвета отчета — как в книге Excel (Всего — золотистый, Устранено — зеленый, Устраняется — красный, Рассм. — синий).</summary>
internal static class Palette
{
    public static readonly Color Total = Color.FromRgb(0x80, 0x60, 0x00);
    public static readonly Color Fixed = Color.FromRgb(0x00, 0xB0, 0x50);
    public static readonly Color Fixing = Color.FromRgb(0xFF, 0x00, 0x00);
    public static readonly Color Review = Color.FromRgb(0x00, 0x70, 0xC0);
    public static readonly Color Text = Color.FromRgb(0x00, 0x00, 0x00);
    public static readonly Color TotalsRow = Color.FromRgb(0xF2, 0xF2, 0xF2);
    public static readonly Color Error = Color.FromRgb(0xFF, 0xC7, 0xCE);
    public static readonly Color CurrentStage = Color.FromRgb(0xE2, 0xEF, 0xDA);
    public static readonly Color Muted = Color.FromRgb(0x59, 0x59, 0x59);

    // Диаграммы
    public static readonly XColor ChartFixing = XColor.FromArgb(0xFF, 0x99, 0x99);
    public static readonly XColor ChartReview = XColor.FromArgb(0x00, 0x70, 0xC0);
    public static readonly XColor ChartFixedPerDay = XColor.FromArgb(0x92, 0xD0, 0x50);
    public static readonly XColor ChartTotal = XColor.FromArgb(0xBF, 0x90, 0x00);
    public static readonly XColor ChartFixed = XColor.FromArgb(0x00, 0xB0, 0x50);
    public static readonly XColor ChartLoaded = XColor.FromArgb(0x00, 0x70, 0xC0);
    public static readonly XColor ChartGrid = XColor.FromArgb(0xD9, 0xD9, 0xD9);
    public static readonly XColor ChartAxis = XColor.FromArgb(0x7F, 0x7F, 0x7F);
    public static readonly XColor ChartText = XColor.FromArgb(0x40, 0x40, 0x40);
    public static readonly XColor ChartDeadline = XColor.FromArgb(0xC0, 0x00, 0x00);
    public static readonly XColor ChartFrame = XColor.FromArgb(0xBF, 0xBF, 0xBF);
}
