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

        // 🌟 ارتقاء: تجمیع ویژگی‌های بوردر در کلاس مشترک
        public BorderDetail Borders { get; set; }

        public double SpaceAfter { get; set; }
        public double SpaceBefore { get; set; }
        public double? LineSpacing { get; set; }       // ضریبِ فاصله‌ی خطوط (240=تک → 1.0)
        public double? IndentLeft { get; set; }
        public double? IndentRight { get; set; }
        public double? IndentFirstLine { get; set; }

        // 🌟 لیست‌های شماره‌دار/بولت (numbering که در C# به مارکرِ آماده تبدیل شده)
        public string ListType { get; set; }           // "ordered" | "bullet"
        public int? ListLevel { get; set; }            // 0 = سطح اول
        public string ListMarker { get; set; }         // "1." , "a)" , "•"

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

        public int? ImageWidth { get; set; }
        public int? ImageHeight { get; set; }

        public string FloatPosition { get; set; }
        public string FillColor { get; set; }
        public string TextColor { get; set; }

        // 🌟 ارتقاء: اضافه شدن کلاس مشترک بوردر در سطح متن (Span)
        public BorderDetail Borders { get; set; }

        public string TableStyleId { get; set; }
        public string TableStyleName { get; set; }
        public string TableAlignment { get; set; }
        public string HasBorders { get; set; }
        public double? TableWidthPercent { get; set; }

        public List<SpanData> InnerSpans { get; set; } = new List<SpanData>();
        public List<TableRowData> TableRows { get; set; } = new List<TableRowData>();
        // 🌟 ریسپانسیو (نتیجهٔ lowering از روی نام استایلِ Word)
        public string ResponsiveStrategy { get; set; }  // "horizontalScroll" | "collapseToCards"
        public string LayoutDirection { get; set; }      // برای Type=="layout": "row" | "column"
        public string LayoutReflow { get; set; }         // "stack" | "wrap" | "none"
    }

    public class TableRowData
    {
        public bool IsHeader { get; set; }
        public List<TableCellData> Cells { get; set; } = new List<TableCellData>();
    }

    public class BorderDetail
    {
        public double? Width { get; set; }
        public string Color { get; set; }
        public string Val { get; set; }
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