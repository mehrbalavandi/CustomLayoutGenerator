using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordToJsonParser
{
    /// <summary>
    /// Word شماره‌ی دیده‌شده‌ی لیست‌ها را در متن ذخیره نمی‌کند؛ آن را از numbering.xml و
    /// شمارنده‌های در حال اجرا محاسبه می‌کند. این کلاس همان کار را در C# انجام می‌دهد و
    /// یک مارکرِ آماده مثل "1." , "a)" , "iv." یا بولت برمی‌گرداند.
    /// یک نمونه به‌ازای هر سند؛ برای هر پاراگرافِ لیست به‌ترتیبِ سند Next() صدا زده می‌شود.
    /// </summary>
    public class NumberingResolver
    {
        private readonly Dictionary<int, AbstractNum> _absByNum = new Dictionary<int, AbstractNum>();
        private readonly Dictionary<int, Dictionary<int, int>> _startOverride = new Dictionary<int, Dictionary<int, int>>();
        private readonly Dictionary<int, Dictionary<int, int>> _counters = new Dictionary<int, Dictionary<int, int>>();

        public NumberingResolver(MainDocumentPart mainPart)
        {
            var numbering = mainPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering == null) return;

            var absById = numbering.Elements<AbstractNum>()
                .Where(a => a.AbstractNumberId?.Value != null)
                .ToDictionary(a => a.AbstractNumberId.Value, a => a);

            foreach (var num in numbering.Elements<NumberingInstance>())
            {
                if (num.NumberID?.Value is not int numId) continue;

                var absRef = num.GetFirstChild<AbstractNumId>();
                if (absRef?.Val != null && absById.TryGetValue(absRef.Val.Value, out var abs))
                    _absByNum[numId] = abs;

                foreach (var ov in num.Elements<LevelOverride>())
                {
                    var start = ov.StartOverrideNumberingValue?.Val?.Value;
                    if (ov.LevelIndex?.Value is int lvl && start.HasValue)
                    {
                        if (!_startOverride.TryGetValue(numId, out var m))
                            _startOverride[numId] = m = new Dictionary<int, int>();
                        m[lvl] = start.Value;
                    }
                }
            }
        }

        public bool HasNumbering => _absByNum.Count > 0;

        /// <summary>numId/level مؤثرِ یک پاراگراف: اول numPr مستقیم، بعد ارثی از استایل. null اگر لیست نباشد.</summary>
        public static (int NumId, int Level)? ReadNumPr(Paragraph p, MainDocumentPart mainPart)
        {
            var numPr = p.ParagraphProperties?.NumberingProperties;
            int? numId = numPr?.NumberingId?.Val?.Value;
            int level = numPr?.NumberingLevelReference?.Val?.Value ?? 0;

            if (numId == null)
            {
                var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                var style = mainPart?.StyleDefinitionsPart?.Styles?.Elements<Style>()
                    .FirstOrDefault(s => s.StyleId?.Value == styleId);
                numId = style?.StyleParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
            }

            if (numId == null || numId == 0) return null;
            return (numId.Value, level);
        }

        /// <summary>شمارنده را جلو می‌برد و (مارکر، نوع) را برمی‌گرداند. نوع = "ordered" | "bullet".</summary>
        public (string Marker, string Kind)? Next(int numId, int level)
        {
            if (!_absByNum.TryGetValue(numId, out var abs)) return null;

            if (!_counters.TryGetValue(numId, out var counters))
                _counters[numId] = counters = new Dictionary<int, int>();

            if (!counters.ContainsKey(level))
                counters[level] = StartOf(numId, abs, level) - 1;
            counters[level]++;

            // سطوحِ عمیق‌تر دفعه‌ی بعد از نو شروع شوند
            foreach (var deeper in counters.Keys.Where(k => k > level).ToList())
                counters[deeper] = StartOf(numId, abs, deeper) - 1;

            var lvlDef = LevelOf(abs, level);
            var fmt = lvlDef?.NumberingFormat?.Val?.Value ?? NumberFormatValues.Decimal;
            if (fmt == NumberFormatValues.Bullet)
                return (BulletFor(level), "bullet");

            var template = lvlDef?.LevelText?.Val?.Value ?? $"%{level + 1}.";
            return (FormatTemplate(abs, numId, counters, template), "ordered");
        }

        private string FormatTemplate(AbstractNum abs, int numId, Dictionary<int, int> counters, string template)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < template.Length; i++)
            {
                if (template[i] == '%' && i + 1 < template.Length && char.IsDigit(template[i + 1]))
                {
                    int refLevel = (template[i + 1] - '0') - 1; // %1 -> level 0
                    int val = counters.TryGetValue(refLevel, out var v) ? v : StartOf(numId, abs, refLevel);
                    var f = LevelOf(abs, refLevel)?.NumberingFormat?.Val?.Value ?? NumberFormatValues.Decimal;
                    sb.Append(FormatNumber(val, f));
                    i++;
                }
                else sb.Append(template[i]);
            }
            return sb.ToString();
        }

        private int StartOf(int numId, AbstractNum abs, int level)
        {
            if (_startOverride.TryGetValue(numId, out var m) && m.TryGetValue(level, out var s)) return s;
            return LevelOf(abs, level)?.StartNumberingValue?.Val?.Value ?? 1;
        }

        private static Level LevelOf(AbstractNum abs, int level) =>
            abs.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == level);

        private static string BulletFor(int level)
        {
            string[] b = { "\u2022", "\u25E6", "\u25AA", "\u2023", "\u00B7" }; // • ◦ ▪ ‣ ·
            return b[level % b.Length];
        }

        private static string FormatNumber(int n, NumberFormatValues fmt)
        {
            if (n < 1) n = 1;
            if (fmt == NumberFormatValues.LowerLetter) return ToLetters(n, false);
            if (fmt == NumberFormatValues.UpperLetter) return ToLetters(n, true);
            if (fmt == NumberFormatValues.LowerRoman) return ToRoman(n).ToLowerInvariant();
            if (fmt == NumberFormatValues.UpperRoman) return ToRoman(n);
            return n.ToString();
        }

        private static string ToLetters(int n, bool upper)
        {
            var sb = new StringBuilder();
            while (n > 0) { n--; sb.Insert(0, (char)((upper ? 'A' : 'a') + (n % 26))); n /= 26; }
            return sb.ToString();
        }

        private static string ToRoman(int n)
        {
            int[] vals = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] syms = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            var sb = new StringBuilder();
            for (int i = 0; i < vals.Length && n > 0; i++)
                while (n >= vals[i]) { sb.Append(syms[i]); n -= vals[i]; }
            return sb.ToString();
        }
    }
}