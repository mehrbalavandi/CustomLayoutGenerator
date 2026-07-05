using System.Collections.Generic;

namespace WordToJsonParser
{
    public class PageData
    {
        public int PageNumber { get; set; }
        public List<ParagraphData> Paragraphs { get; set; } = new List<ParagraphData>();
    }

    public class ParagraphData
    {
        public string Alignment { get; set; }
        public string Direction { get; set; }
        public string FillColor { get; set; }

        public string HasBorders { get; set; }
        public string BorderColor { get; set; }
        public string BorderStyle { get; set; }

        public double SpaceAfter { get; set; }
        public double SpaceBefore { get; set; }
        public double? IndentLeft { get; set; }
        public double? IndentRight { get; set; }
        public double? IndentFirstLine { get; set; }

        // 🌟 فیلدهای اضافه‌شده برای پشتیبانی از سیستم همگام‌سازی فایل صوتی (Karaoke)
        public int? StartMs { get; set; }
        public int? EndMs { get; set; }
        public string AudioTrackName { get; set; }

        public List<SpanData> Spans { get; set; } = new List<SpanData>();
    }

    public class SpanData
    {
        public string Type { get; set; }
        public string Content { get; set; }
        public List<string> Markers { get; set; } = new List<string>();
        public string Url { get; set; }

        public string FloatPosition { get; set; }
        public string FillColor { get; set; }
        public string TextColor { get; set; }
        public string BorderColor { get; set; }
        public string BorderStyle { get; set; }
        public double? BorderWidth { get; set; }

        public string TableStyleId { get; set; }
        public string TableStyleName { get; set; }
        public string TableAlignment { get; set; }
        public string HasBorders { get; set; }
        public double? TableWidthPercent { get; set; }

        public List<SpanData> InnerSpans { get; set; } = new List<SpanData>();
        public List<TableRowData> TableRows { get; set; } = new List<TableRowData>();
    }

    public class TableRowData
    {
        public bool IsHeader { get; set; }
        public List<TableCellData> Cells { get; set; } = new List<TableCellData>();
    }
    // 🌟 کلاس‌های جدید برای نگهداری جزئیات هر خط مرزی
    public class BorderDetail
    {
        public double? Width { get; set; } // ضخامت بر اساس Point
        public string Color { get; set; }  // کد رنگ Hex
        public string Val { get; set; }    // نوع خط (single, dashed و ...)
    }

    public class CellBorders
    {
        public BorderDetail Top { get; set; }
        public BorderDetail Bottom { get; set; }
        public BorderDetail Left { get; set; }
        public BorderDetail Right { get; set; }
    }
    public class TableCellData
    {
        public List<ParagraphData> Paragraphs { get; set; } = new List<ParagraphData>();
        public string FillColor { get; set; }
        public double? WidthPercent { get; set; }
        public string VAlign { get; set; }
        public int? ColSpan { get; set; }
        public string RowMerge { get; set; }
        public bool IsHeaderCell { get; set; }
        public CellBorders Borders { get; set; }
        public double? PaddingTop { get; set; }
        public double? PaddingBottom { get; set; }
        public double? PaddingLeft { get; set; }
        public double? PaddingRight { get; set; }
    }
}