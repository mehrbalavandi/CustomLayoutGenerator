using DocumentFormat.OpenXml.Packaging;

namespace CustomLayoutGenerator; // نام namespace را با پروژه‌ی خودتان یکی کنید

public static class CyrillicFileTool
{
    /// <summary>
    /// از فایل ورودی یک کپی می‌سازد، حروف سیریلیک هم‌شکل را در آن اصلاح می‌کند
    /// و گزارش را کنار فایل خروجی می‌نویسد.
    /// </summary>
    public static CyrillicReport FixFile(string inputPath, string outputPath)
    {
        File.Copy(inputPath, outputPath, overwrite: true);

        using var doc = WordprocessingDocument.Open(outputPath, true);
        var report = CyrillicFixer.Fix(doc);

        // گزارش با نام فایل ورودی ساخته می‌شود (اگر موردی نباشد فایلی نوشته نمی‌شود)
        string reportPath = Path.Combine(
            Path.GetDirectoryName(outputPath)!,
            Path.GetFileNameWithoutExtension(inputPath) + "_cyrillic_report.txt");
        CyrillicFixer.WriteReport(report, reportPath, Path.GetFileName(inputPath));

        return report;
    }
}