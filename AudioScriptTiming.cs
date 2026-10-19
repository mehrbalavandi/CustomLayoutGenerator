using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace CustomLayoutGenerator
{
    /// <summary>
    /// 🌟 A-2، A-3: پس‌پردازشِ زمان‌بندیِ اسپن‌های اسکریپتِ صوتی، بعد از اینکه مارکرها
    /// (A-1) به StartMs/EndMs تبدیل شدند.
    ///
    /// A-2 «گفته‌نشده»: برچسبِ گوینده در ابتدای پاراگراف («Man:»، «Bank clerk:»،
    /// «Candidate 1:»؛ ۱ تا ۳ کلمه، حداکثر ۲۴ نویسه، بی کلمه‌ی دستوری — <see cref="IsLabel"/>) و هر متنی با استایلِ کاراکتریِ «NotSpoken» در فایلِ صوتی خوانده
    /// نمی‌شود. چنین اسپنی Unspoken=true می‌گیرد؛ زمانش می‌ماند (فاصله‌ی سکوت/صدای
    /// بی‌متن را نگه می‌دارد) ولی اپ هرگز هایلایتش نمی‌کند.
    ///
    /// A-3 «کلمه‌ی بی‌زمان»: aligner عدد و رقم را نمی‌تواند تراز کند («1»، «$10.50»،
    /// «1985.»)، پس طولِ صفرِ ثانیه می‌گیرند و هیچ‌وقت هایلایت نمی‌شوند، در حالی که خوانده
    /// می‌شوند («Conversation 1»). چنین کلمه‌ای بازه‌ی نزدیک‌ترین کلمه‌ی گفته‌شده‌ی قبلی (یا
    /// اگر نیست، بعدیِ) همان پاراگراف را می‌گیرد و با آن هم‌زمان هایلایت می‌شود.
    /// </summary>
    public static class AudioScriptTiming
    {
        /// <summary>مارکرِ اسپن برای استایلِ کاراکتریِ NotSpoken (در ParseRun گذاشته می‌شود).</summary>
        public const string NotSpokenMarker = "ns";

        // برچسب: ۱ تا ۳ کلمه در ابتدای پاراگراف، اولی با حرفِ بزرگ، آخری با «:» تمام می‌شود.
        private static readonly Regex LabelRegex = new Regex(
            @"^\s*\p{Lu}[\p{L}\p{N}'’.\-]*(?:[ \t]+[\p{L}\p{N}'’.\-]+){0,2}:(?=\s|$)",
            RegexOptions.Compiled);
        private const int MaxLabelLength = 24;
        // عبارتِ گفته‌شده‌ای که به «:» ختم می‌شود («The answer is:») برچسب نیست: برچسب نام و نقش است.
        // ⚠️ عیناً همان LABEL_STOPWORDS در align_engine/speech_text.py.
        private static readonly HashSet<string> LabelStopwords = new HashSet<string>(
            ("is are was were be been am the a an my your his her its our their this that these those " +
             "it i you we they he she and or but to of in on for with at as here there what which who " +
             "so if then do does did have has had not no yes").Split(' '), StringComparer.Ordinal);

        /// <summary>«Man:»، «A:»، «Bank clerk:»، «Student A:» بله؛ «The answer is:» نه.</summary>
        public static bool IsLabel(string prefix)
        {
            string p = prefix.Trim();
            if (p.Length > MaxLabelLength) return false;
            var words = p.TrimEnd(':').Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length == 1) return true;
            return !words.Any(w =>
                LabelStopwords.Contains(w.ToLowerInvariant().Trim('\'', '’', '.', '-')) &&
                !(w.Length == 1 && char.IsUpper(w[0])));
        }

        public static void Apply(IEnumerable<AudioScriptTrack> tracks)
        {
            foreach (var track in tracks)
            {
                foreach (var para in track.Paragraphs)
                {
                    MarkUnspoken(para);
                    FillZeroLength(para);
                }
            }
        }

        private static IEnumerable<SpanData> TextSpans(ParagraphData p) =>
            p.Spans.Where(s => s.Type == "text" && s.Content != null);

        private static void MarkUnspoken(ParagraphData para)
        {
            foreach (var s in TextSpans(para))
                if (s.Markers != null && s.Markers.Contains(NotSpokenMarker))
                    s.Unspoken = true;

            // برچسبِ گوینده: فقط اسپن‌هایی که *کاملاً* داخلِ برچسب‌اند
            var spans = TextSpans(para).ToList();
            string text = string.Concat(spans.Select(s => s.Content));
            var m = LabelRegex.Match(text);
            if (!m.Success || !IsLabel(m.Value)) return;
            int labelEnd = m.Index + m.Length;
            int pos = 0;
            foreach (var s in spans)
            {
                int end = pos + s.Content.Length;
                if (end <= labelEnd && s.Content.Trim().Length > 0) s.Unspoken = true;
                if (end >= labelEnd) break;
                pos = end;
            }
        }

        private static bool IsTimedWord(SpanData s) =>
            s.StartMs.HasValue && s.EndMs.HasValue && s.Unspoken != true &&
            s.Content.Any(char.IsLetterOrDigit);

        private static void FillZeroLength(ParagraphData para)
        {
            var spans = TextSpans(para).ToList();
            for (int i = 0; i < spans.Count; i++)
            {
                var s = spans[i];
                if (!IsTimedWord(s) || s.StartMs != s.EndMs) continue;
                SpanData donor = null;
                for (int j = i - 1; j >= 0 && donor == null; j--)
                    if (IsTimedWord(spans[j]) && spans[j].EndMs > spans[j].StartMs) donor = spans[j];
                for (int j = i + 1; j < spans.Count && donor == null; j++)
                    if (IsTimedWord(spans[j]) && spans[j].EndMs > spans[j].StartMs) donor = spans[j];
                if (donor == null) continue;
                s.StartMs = donor.StartMs;
                s.EndMs = donor.EndMs;
            }
        }
    }
}
