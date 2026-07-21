using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WordToJsonParser
{

    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private int _currentPage = 1;
        private int _currentSection = 1;
        private int _imageCounter = 1;

        // متغیرهای مدیریت صوتی و جای‌خالی
        private string _activeAudioTrack = null;
        private ParagraphData _lastAudioParagraph = null;
        private HashSet<ParagraphData> _blankWord2Set = new HashSet<ParagraphData>();

        public void ResetCounters()
        {
            _currentPage = 1;
            _currentSection = 1;
            _imageCounter = 1;
            _activeAudioTrack = null;
            _lastAudioParagraph = null;
            _blankWord2Set.Clear();
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "انتخاب فایل اصلی کتاب (Word)";
                openFileDialog.Filter = "Word Documents (*.docx)|*.docx";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string originalFile = openFileDialog.FileName;

                    // ساخت فایل موقت
                    string tempFile = Path.Combine(
                        Path.GetTempPath(),
                        Guid.NewGuid().ToString() + Path.GetExtension(originalFile));

                    File.Copy(originalFile, tempFile, true);

                    string filePath = openFileDialog.FileName;
                    try
                    {
                        string outputDir = Path.GetDirectoryName(filePath);

                        ResetCounters();

                        // ۱. پردازش کتاب اصلی
                        List<PageData> pages = ProcessWordDocument(tempFile, outputDir);

                        // ادغام پاراگراف‌های BlankWord2 برای کتاب اصلی
                        foreach (var page in pages)
                        {
                            page.Paragraphs = MergeBlankWord2Paragraphs(page.Paragraphs);
                        }

                        List<ParagraphData> audioScripts = new List<ParagraphData>();

                        // ۲. دریافت فایل صوتی در صورت نیاز
                        DialogResult hasAudio = MessageBox.Show(
                            "آیا فایل اسکریپت صوتی (Audio Script) هم برای این کتاب دارید؟",
                            "اسکریپت صوتی",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (hasAudio == DialogResult.Yes)
                        {
                            using (OpenFileDialog audioDialog = new OpenFileDialog())
                            {
                                audioDialog.Title = "انتخاب فایل اسکریپت صوتی (Word)";
                                audioDialog.Filter = "Word Documents (*.docx)|*.docx";
                                if (audioDialog.ShowDialog() == DialogResult.OK)
                                {
                                    audioScripts = ProcessAudioScriptWordFile(audioDialog.FileName, outputDir);
                                    // ادغام پاراگراف‌های BlankWord2 برای اسکریپت‌های صوتی
                                    audioScripts = MergeBlankWord2Paragraphs(audioScripts);
                                }
                            }
                        }

                        // 🌟 ۱. ریسپانسیو: نام استایل‌های جدول → primitiveهای declarative
                        ResponsiveLowering.Apply(pages);

                        // 🌟 ۲. خروجی per-page + index.json (نسخهٔ هر صفحه = هَشِ محتوا)
                        BookOutputWriter.Write(outputDir, pages, audioScripts);

                        MessageBox.Show(
                            $"کتاب با موفقیت به {pages.Count} فایلِ صفحه + index.json تبدیل شد!",
                            "عملیات موفق", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    finally
                    {
                        // حذف فایل موقت
                        if (File.Exists(tempFile))
                        {
                            File.Delete(tempFile);
                        }
                    }
                }
            }
        }

        private List<PageData> ProcessWordDocument(string filePath, string outputDir)
        {
            using var wordDocument = WordprocessingDocument.Open(filePath, false);
            var resolver = new FontResolver(wordDocument);
            List<PageData> pages = new List<PageData>();
            PageData currentPage = new PageData { PageNumber = 1 };
            pages.Add(currentPage);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;

                foreach (var element in body.Elements())
                {
                    if (element.Descendants<Break>().Any(b => b.Type != null && b.Type.Value == BreakValues.Page))
                    {
                        currentPage = new PageData { PageNumber = pages.Count + 1 };
                        pages.Add(currentPage);
                        continue;
                    }

                    if (element is Paragraph paragraph)
                    {
                        bool isBlankWord2 = IsTargetStyle(paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value, wordDoc.MainDocumentPart, "BlankWord2");

                        var paraDataList = ParseParagraph(paragraph, wordDoc.MainDocumentPart, resolver, outputDir, false);

                        // 🌟 فیلتر پاراگراف‌های زباله: حذف پاراگراف‌هایی که فقط کاراکترهای \n یا فضای خالی دارند
                        paraDataList.RemoveAll(p =>
                            p.Spans.All(s => s.Type == "text" && string.IsNullOrWhiteSpace(s.Content)) &&
                            p.StartMs == null);

                        if (paraDataList.Count > 0)
                        {
                            currentPage.Paragraphs.AddRange(paraDataList);

                            if (isBlankWord2)
                            {
                                foreach (var p in paraDataList) _blankWord2Set.Add(p);
                            }
                        }
                    }
                    else if (element is Table table)
                    {
                        var tableSpan = ParseTable(table, wordDoc.MainDocumentPart, resolver, outputDir);
                        var para = new ParagraphData();
                        para.Spans.Add(tableSpan);
                        currentPage.Paragraphs.Add(para);
                    }
                }
            }

            return pages;
        }

        private List<ParagraphData> ProcessAudioScriptWordFile(string filepath, string outputDir)
        {
            var audioScripts = new List<ParagraphData>();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filepath, false))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                var resolver = new FontResolver(wordDoc);
                var paragraphs = mainPart.Document.Body.Elements<Paragraph>();

                foreach (var p in paragraphs)
                {
                    bool isBlankWord2 = IsTargetStyle(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value, mainPart, "BlankWord2");

                    var parsedParas = ParseParagraph(p, mainPart, resolver, outputDir, false);

                    // 🌟 فیلتر پاراگراف‌های زباله و نامرئی
                    parsedParas.RemoveAll(pr =>
                        pr.Spans.All(s => s.Type == "text" && string.IsNullOrWhiteSpace(s.Content)) &&
                        pr.StartMs == null);

                    if (parsedParas.Count > 0)
                    {
                        audioScripts.AddRange(parsedParas);

                        if (isBlankWord2)
                        {
                            foreach (var cp in parsedParas) _blankWord2Set.Add(cp);
                        }
                    }
                }
            }

            return audioScripts;
        }

        // ==========================================
        // موتور قدرتمند برای جستجوی نام استایل
        // ==========================================
        private bool IsTargetStyle(string styleId, MainDocumentPart mainPart, string targetName)
        {
            if (string.IsNullOrEmpty(styleId)) return false;

            if (styleId.Replace(" ", "").Equals(targetName, StringComparison.OrdinalIgnoreCase))
                return true;

            var stylesPart = mainPart.StyleDefinitionsPart;
            if (stylesPart?.Styles != null)
            {
                var style = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleId);
                var styleName = style?.StyleName?.Val?.Value;
                if (styleName != null && styleName.Replace(" ", "").Equals(targetName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        // ==========================================
        // موتور ادغام پاراگراف‌های BlankWord2
        // ==========================================
        private List<ParagraphData> MergeBlankWord2Paragraphs(List<ParagraphData> input)
        {
            var result = new List<ParagraphData>();
            List<ParagraphData> group = new List<ParagraphData>();

            void flushGroup()
            {
                if (group.Count == 0) return;

                var merged = CloneParagraphProperties(group[0]);
                merged.Spans = new List<SpanData>();
                var nonTextSpans = new List<SpanData>();

                var blankParentSpan = new SpanData
                {
                    Type = "text",
                    InnerSpans = new List<SpanData>()
                };

                string combinedRawText = "";

                for (int i = 0; i < group.Count; i++)
                {
                    var p = group[i];

                    // 🐞 رفع باگِ گم‌شدنِ شماره‌های ۲ به بعد: merged فقط مارکرِ
                    // group[0] را نگه می‌دارد (در CloneParagraphProperties)، پس
                    // بدونِ این تزریق، ListMarker خودِ این پاراگراف (که
                    // NumberingResolver/Numbering(...).Next(...) پیش‌تر برایش
                    // به‌درستی محاسبه کرده، مثلاً "2:"، "3:") برای همیشه دور
                    // ریخته می‌شد و در JSON اصلاً ظاهر نمی‌شد. این‌جا همان
                    // مارکرِ واقعیِ Word را مستقیماً به‌صورتِ متنِ معمولی، درست
                    // قبل از محتوای همان پاراگراف، در ترکیب می‌گذاریم — همینه
                    // که JSON خودش کامل و خودکفا می‌ماند و سمتِ Flutter مجبور
                    // نیست شماره را حدس بزند یا افزایش بدهد (که برای بولت/حرف
                    // اصلاً معنا ندارد، ولی این‌جا چون مارکرِ واقعیِ همان خط
                    // است، برای هر نوع لیستی درست کار می‌کند).
                    if (i > 0 && !string.IsNullOrEmpty(p.ListMarker))
                    {
                        // 🐞 رفع باگِ گزارش‌شده (بولد‌نبودنِ شماره‌های ۲ به بعد):
                        // چون این SpanData یک متنِ معمولیِ تازه‌ساز است، نه یک
                        // run واقعیِ کپی‌شده از سند، بدونِ این مارکر به‌صورتِ
                        // پیش‌فرض normal-weight رندر می‌شد؛ در حالی‌که در Word
                        // شماره‌های لیست بولد بودند (و شماره‌ی خطِ اول هم چون
                        // در Flutter هاردکد بولد است، بولد دیده می‌شد — همین
                        // ناهم‌خوانی، فقط شماره‌ی اول بولد/بقیه عادی، دقیقاً
                        // چیزی بود که در اسکرین‌شات دیده شد).
                        var markerSpan = new SpanData { Type = "text", Content = p.ListMarker + " ", Markers = new List<string> { "b" } };
                        blankParentSpan.InnerSpans.Add(markerSpan);
                        combinedRawText += p.ListMarker + " ";
                    }

                    foreach (var span in p.Spans)
                    {
                        if (span.Type == "text")
                        {
                            var cleanSpan = CloneSpan(span);
                            if (!string.IsNullOrEmpty(cleanSpan.Content))
                            {
                                cleanSpan.Content = cleanSpan.Content.Replace("{blk}", "").Replace("{/blk}", "");
                                blankParentSpan.InnerSpans.Add(cleanSpan);
                                combinedRawText += cleanSpan.Content;
                            }
                        }
                        else
                        {
                            nonTextSpans.Add(span);
                        }
                    }

                    if (i < group.Count - 1)
                    {
                        blankParentSpan.InnerSpans.Add(new SpanData { Type = "text", Content = "\n" });
                        combinedRawText += "\n";
                    }
                }

                blankParentSpan.Content = "{blk}" + combinedRawText + "{/blk}";

                var firstTextSpan = blankParentSpan.InnerSpans.FirstOrDefault();
                if (firstTextSpan != null)
                {
                    // 🌟 اصلاح شد: مارکرهای ساختاری متنی مثل b، i و u فیلتر می‌شوند تا به کل دکمه والد ارث نرسند
                    blankParentSpan.Markers = firstTextSpan.Markers != null
                        ? firstTextSpan.Markers.Where(m => m != "b" && m != "i" && m != "u").ToList()
                        : new List<string>();
                    // 🌟 هر دو مقدار رنگ متن و رنگ پس‌زمینه را از دکمه والد می‌گیریم
                    blankParentSpan.FillColor = null;
                    blankParentSpan.TextColor = null;
                    // 🌟 مسدود کردن ارث‌بری بوردر برای دکمه اصلی
                    //blankParentSpan.HasBorders = null;
                    blankParentSpan.Borders = null;
                }

                merged.Spans.Add(blankParentSpan);
                merged.Spans.AddRange(nonTextSpans);

                merged.StartMs = group.First().StartMs;
                merged.EndMs = group.Last().EndMs;
                merged.AudioTrackName = group.First().AudioTrackName;

                result.Add(merged);
                group.Clear();
            }

            foreach (var p in input)
            {
                if (_blankWord2Set.Contains(p))
                {
                    group.Add(p);
                }
                else
                {
                    flushGroup();
                    result.Add(p);
                }
            }
            flushGroup();

            return result;
        }

        public List<ParagraphData> ParseParagraph(Paragraph p, MainDocumentPart mainPart, FontResolver resolver, string outputDir, bool inTable = false)
        {
            var basePara = new ParagraphData();

            // 🌟 لیست‌ها: numId/level را بخوان و مارکرِ آماده را از روی شمارنده‌ها بساز
            var _np = NumberingResolver.ReadNumPr(p, mainPart);
            if (_np != null)
            {
                var _li = Numbering(mainPart).Next(_np.Value.NumId, _np.Value.Level);
                if (_li != null)
                {
                    basePara.ListMarker = _li.Value.Marker;
                    basePara.ListType = _li.Value.Kind;
                    basePara.ListLevel = _np.Value.Level;
                }
            }

            // استخراج استایل و خواص پایه پاراگراف
            if (p.ParagraphProperties != null)
            {
                var bidi = p.ParagraphProperties.BiDi;
                basePara.Direction = (bidi != null) ? "RTL" : "LTR";

                var justification = p.ParagraphProperties.Justification;
                if (justification?.Val != null)
                    basePara.Alignment = MapAlignment(justification.Val.Value);

                // 🌟 استخراج دقیق تورفتگی‌های پاراگراف (Indentation)
                if (p.ParagraphProperties.Indentation != null)
                {
                    var ind = p.ParagraphProperties.Indentation;
                    if (ind.Left != null && double.TryParse(ind.Left.Value, out double lTwips)) basePara.IndentLeft = lTwips / 20.0;
                    if (ind.Right != null && double.TryParse(ind.Right.Value, out double rTwips)) basePara.IndentRight = rTwips / 20.0;

                    // تورفتگی خط اول (First Line)
                    if (ind.FirstLine != null && double.TryParse(ind.FirstLine.Value, out double flTwips)) basePara.IndentFirstLine = flTwips / 20.0;
                    // اگر تورفتگی معکوس (Hanging) باشد، آن را به عنوان تورفتگی خط اول منفی در نظر می‌گیریم
                    else if (ind.Hanging != null && double.TryParse(ind.Hanging.Value, out double hTwips)) basePara.IndentFirstLine = -(hTwips / 20.0);
                }

                var pShading = p.ParagraphProperties.Shading?.Fill?.Value;
                if (!string.IsNullOrEmpty(pShading) && pShading != "auto")
                {
                    basePara.FillColor = pShading;
                }

                double? spaceAfter = null;
                double? spaceBefore = null;

                var spacing = p.ParagraphProperties.SpacingBetweenLines;
                if (spacing != null)
                {
                    if (spacing.After != null && double.TryParse(spacing.After.Value, out double afterTwips))
                        spaceAfter = afterTwips / 20.0;

                    if (spacing.Before != null && double.TryParse(spacing.Before.Value, out double beforeTwips))
                        spaceBefore = beforeTwips / 20.0;

                    // 🌟 فاصله‌ی بین خطوطِ یک پاراگراف (line spacing)
                    if (spacing.Line != null && double.TryParse(spacing.Line.Value, out double lineVal))
                    {
                        var rule = spacing.LineRule?.Value;
                        // حالت auto/multiple: مقدار در 240امِ خط است (240=تک، 360=۱٫۵، 480=دوبل)
                        if (rule == null || rule == LineSpacingRuleValues.Auto)
                            basePara.LineSpacing = lineVal / 240.0;
                        // حالت exact/atLeast مقدارِ مطلق (twip) است؛ فعلاً رد می‌کنیم تا مقیاسِ اشتباه اعمال نشود
                    }
                }

                if (p.ParagraphProperties.ParagraphStyleId != null)
                {
                    string styleId = p.ParagraphProperties.ParagraphStyleId.Val.Value;
                    if (spaceAfter == null) spaceAfter = GetSpacingFromStyle(mainPart, styleId, true);
                    if (spaceBefore == null) spaceBefore = GetSpacingFromStyle(mainPart, styleId, false);
                }

                if (spaceAfter == null) spaceAfter = inTable ? 0 : (GetDefaultSpacing(mainPart, true) ?? 0);
                if (spaceBefore == null) spaceBefore = inTable ? 0 : (GetDefaultSpacing(mainPart, false) ?? 0);

                basePara.SpaceAfter = spaceAfter.Value;
                basePara.SpaceBefore = spaceBefore.Value;

                var pBorders = p.ParagraphProperties.ParagraphBorders;
                BorderType pBorder = null;
                if (pBorders != null)
                {
                    pBorder = (BorderType)pBorders.TopBorder ?? (BorderType)pBorders.BottomBorder ?? (BorderType)pBorders.LeftBorder ?? (BorderType)pBorders.RightBorder;
                }

                if (pBorder == null && p.ParagraphProperties.ParagraphStyleId != null)
                {
                    pBorder = GetParagraphBorderFromStyle(mainPart, p.ParagraphProperties.ParagraphStyleId.Val.Value);
                }

                if (pBorder != null && pBorder.Val != null && pBorder.Val.Value != BorderValues.None)
                {
                    basePara.Borders = ParseBorder(pBorder);
                }

                if (p.ParagraphProperties.SectionProperties != null)
                    _currentSection++;

                // 🌟 اضافه شده: پاک کردن رنگ و حاشیه بصری برای BlankWord2
                bool isBlankWord2 = IsTargetStyle(p.ParagraphProperties.ParagraphStyleId?.Val?.Value, mainPart, "BlankWord2");
                if (isBlankWord2)
                {
                    basePara.FillColor = null;
                    basePara.Borders = null;
                }
            }

            SpanData lastTextSpan = null;

            foreach (var element in p.Elements())
            {
                if (element is Hyperlink hyperlink)
                {
                    string url = GetHyperlinkUrl(mainPart, hyperlink.Id);
                    foreach (var run in hyperlink.Elements<Run>())
                    {
                        ProcessRun(run, mainPart, outputDir, basePara, p.ParagraphProperties, ref lastTextSpan, url);
                    }
                }
                else if (element is Run run)
                {
                    ProcessRun(run, mainPart, outputDir, basePara, p.ParagraphProperties, ref lastTextSpan, null);
                }
            }

            var result = new List<ParagraphData>();
            var currentPara = CloneParagraphProperties(basePara);
            bool trackEndedHere = false;
            string combinedPattern = @"(\[AudioStart:\s*.+?\]|\[(?:(?:\d{1,2}):)?\d{1,2}:\d{2}(?:\.\d+)?\]|\[\d+\]|\[AudioEnd:\s*(?:(?:\d{1,2}):)?\d{1,2}:\d{2}(?:\.\d+)?\]|\[AudioEnd\])";

            if (!basePara.Spans.Any())
            {
                result.Add(basePara);
                return result;
            }

            foreach (var span in basePara.Spans)
            {
                if (span.Type != "text" || string.IsNullOrEmpty(span.Content))
                {
                    currentPara.Spans.Add(span);
                    continue;
                }

                var parts = Regex.Split(span.Content, combinedPattern, RegexOptions.IgnoreCase);

                foreach (var part in parts)
                {
                    if (string.IsNullOrEmpty(part)) continue;

                    var startMatch = Regex.Match(part, @"\[AudioStart:\s*(.+?)\]", RegexOptions.IgnoreCase);
                    var endTimeMatch = Regex.Match(part, @"\[AudioEnd:\s*(?:(\d{1,2}):)?(\d{1,2}):(\d{2}(?:\.\d+)?)\]", RegexOptions.IgnoreCase);
                    var endMatch = Regex.Match(part, @"\[AudioEnd\]", RegexOptions.IgnoreCase);
                    var timeMatch = Regex.Match(part, @"\[(?:(\d{1,2}):)?(\d{1,2}):(\d{2}(?:\.\d+)?)\]");
                    var msMatch = Regex.Match(part, @"\[(\d+)\]");

                    if (startMatch.Success)
                    {
                        if (HasText(currentPara))
                        {
                            result.Add(currentPara);
                            currentPara = CloneParagraphProperties(basePara);
                        }

                        _activeAudioTrack = startMatch.Groups[1].Value.Trim();
                        currentPara.StartMs = 0;
                        currentPara.AudioTrackName = _activeAudioTrack;

                        if (_lastAudioParagraph != null && _lastAudioParagraph != currentPara)
                            _lastAudioParagraph.EndMs = 0;

                        _lastAudioParagraph = currentPara;
                    }
                    else if (timeMatch.Success || msMatch.Success)
                    {
                        if (HasText(currentPara))
                        {
                            result.Add(currentPara);
                            currentPara = CloneParagraphProperties(basePara);
                        }

                        int currentStartMs = 0;
                        if (timeMatch.Success)
                        {
                            int startHours = timeMatch.Groups[1].Success ? int.Parse(timeMatch.Groups[1].Value) : 0;
                            int startMin = int.Parse(timeMatch.Groups[2].Value);
                            double startSec = double.Parse(timeMatch.Groups[3].Value, System.Globalization.CultureInfo.InvariantCulture);
                            currentStartMs = (int)((startHours * 3600 + startMin * 60 + startSec) * 1000);
                        }
                        else
                        {
                            currentStartMs = int.Parse(msMatch.Groups[1].Value);
                        }

                        currentPara.StartMs = currentStartMs;
                        currentPara.AudioTrackName = _activeAudioTrack;

                        if (_lastAudioParagraph != null && _lastAudioParagraph != currentPara)
                            _lastAudioParagraph.EndMs = currentStartMs;

                        _lastAudioParagraph = currentPara;
                    }
                    else if (endTimeMatch.Success)
                    {
                        trackEndedHere = true;
                        int endHours = endTimeMatch.Groups[1].Success ? int.Parse(endTimeMatch.Groups[1].Value) : 0;
                        int endMin = int.Parse(endTimeMatch.Groups[2].Value);
                        double endSec = double.Parse(endTimeMatch.Groups[3].Value, System.Globalization.CultureInfo.InvariantCulture);
                        if (_lastAudioParagraph != null)
                            _lastAudioParagraph.EndMs = (int)((endHours * 3600 + endMin * 60 + endSec) * 1000);
                    }
                    else if (endMatch.Success)
                    {
                        trackEndedHere = true;
                        if (_lastAudioParagraph != null && _lastAudioParagraph.EndMs == null)
                            _lastAudioParagraph.EndMs = 9999999;
                    }
                    else
                    {
                        var newSpan = CloneSpan(span);
                        if (currentPara.Spans.Count == 0)
                            newSpan.Content = part.TrimStart();
                        else
                            newSpan.Content = part;

                        // فقط در صورتی اضافه کن که کاملاً خالی نباشد
                        if (!string.IsNullOrEmpty(newSpan.Content))
                            currentPara.Spans.Add(newSpan);
                    }
                }
            }

            if (HasText(currentPara) || currentPara.Spans.Any(s => s.Type != "text"))
            {
                result.Add(currentPara);
            }

            if (trackEndedHere)
            {
                _activeAudioTrack = null;
                _lastAudioParagraph = null;
            }

            if (result.Count == 0)
                result.Add(basePara);

            // ==========================================
            // 🌟 موتور پاک‌سازی (Garbage Collector) اسپن‌ها
            // ==========================================
            foreach (var rPara in result)
            {
                for (int i = rPara.Spans.Count - 1; i >= 0; i--)
                {
                    var span = rPara.Spans[i];
                    if (span.Type == "text" && span.Content != null)
                    {
                        // ۱. اصلاح تگ‌های به هم چسبیده
                        span.Content = span.Content.Replace("{/blk}{blk}", "");

                        // ۲. اگر بعد از حذف تگ‌ها، محتوای اسپن کاملاً خالی ("") شد، کل شیء Span را پاک کن!
                        if (string.IsNullOrEmpty(span.Content))
                        {
                            rPara.Spans.RemoveAt(i);
                        }
                    }
                }
            }

            return result;
        }

        private void ProcessRun(Run run, MainDocumentPart mainPart, string outputDir, ParagraphData paraData, ParagraphProperties pPr, ref SpanData lastTextSpan, string hyperlinkUrl)
        {
            if (run.Descendants<LastRenderedPageBreak>().Any() ||
                run.Elements<Break>().Any(b => b.Type != null && b.Type.Value == BreakValues.Page))
            {
                _currentPage++;
            }

            var drawing = run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().FirstOrDefault();
            if (drawing != null)
            {
                // 🌟 دریافت همزمان نام فایل و ابعاد
                var imgResult = ExtractAndSaveImage(drawing, mainPart, outputDir);

                if (imgResult != null)
                {
                    string floatPos = "none";
                    var anchor = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().FirstOrDefault();
                    if (anchor != null)
                    {
                        var hPos = anchor.Elements<DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalPosition>().FirstOrDefault();
                        var align = hPos?.Elements<DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignment>().FirstOrDefault();
                        if (align != null) floatPos = align.Text.ToLower();
                    }

                    paraData.Spans.Add(new SpanData
                    {
                        Type = "image",
                        Url = imgResult.Value.FileName,
                        FloatPosition = floatPos,
                        ImageWidth = imgResult.Value.Width,   // 🌟 ذخیره عرض در مدل
                        ImageHeight = imgResult.Value.Height  // 🌟 ذخیره ارتفاع در مدل
                    });
                }
                lastTextSpan = null;
                return;
            }

            string runText = "";
            foreach (var child in run.Elements())
            {
                if (child is Text textNode) runText += textNode.Text;
                else if (child is Break br && (br.Type == null || br.Type.Value == BreakValues.TextWrapping)) runText += "\n";
                else if (child is TabChar) runText += "\t";
                else if (child is DocumentFormat.OpenXml.Wordprocessing.SymbolChar sym)
                {
                    if (sym.Char != null && sym.Char.HasValue)
                    {
                        try { runText += (char)Convert.ToInt32(sym.Char.Value, 16); } catch { }
                    }
                }
            }

            if (string.IsNullOrEmpty(runText)) return;

            var runStyleId = run.RunProperties?.RunStyle?.Val?.Value;
            var pStyleId = pPr?.ParagraphStyleId?.Val?.Value;

            if (IsAllCaps(run.RunProperties, runStyleId, pStyleId, mainPart)) runText = runText.ToUpper();

            // تله‌گذاری ایمن برای BlankWord1
            bool isBlankWord1 = IsTargetStyle(runStyleId, mainPart, "BlankWord1");
            if (isBlankWord1)
            {
                runText = "{blk}" + runText + "{/blk}";
            }

            List<string> currentMarkers = ExtractRunMarkers(run, pPr, mainPart);
            string runShading = run.RunProperties?.Shading?.Fill?.Value;
            if (runShading == "auto") runShading = null;

            string runTextColor = null;
            if (run.RunProperties?.Color?.Val?.Value != null && run.RunProperties.Color.Val.Value != "auto")
                runTextColor = run.RunProperties.Color.Val.Value;

            if (string.IsNullOrEmpty(runTextColor))
            {
                if (!string.IsNullOrEmpty(runStyleId)) runTextColor = GetColorFromStyleId(mainPart, runStyleId);
                if (string.IsNullOrEmpty(runTextColor) && pStyleId != null) runTextColor = GetColorFromStyleId(mainPart, pStyleId);
            }

            // 🌟 اصلاح مهم: کادر متنی (Character Border) فقط باید از خود کلمه یا استایلِ مستقیمِ کلمه خوانده شود. 
            // ارث‌بری از استایل پاراگراف (pStyleId) حذف شد تا کادر به تمام کلمات نشت نکند!
            Border rBorder = run.RunProperties?.Border;
            if (rBorder == null && !string.IsNullOrEmpty(runStyleId)) rBorder = GetRunBorderFromStyle(mainPart, runStyleId);

            BorderDetail runBorder = null;
            if (rBorder != null && rBorder.Val != null && rBorder.Val.Value != BorderValues.None && rBorder.Val.Value != BorderValues.Nil)
            {
                runBorder = ParseBorder(rBorder);
            }

            string runBorderColor = null, runHasBorders = null, runBorderStyle = null;
            if (rBorder != null && rBorder.Val != null && rBorder.Val.Value != BorderValues.None && rBorder.Val.Value != BorderValues.Nil)
            {
                runHasBorders = "true";
                runBorderColor = rBorder.Color?.Value;
                if (rBorder.Color != null && runBorderColor != "auto" && string.IsNullOrEmpty(runBorderColor))
                {
                    runBorderColor = rBorder.Color.Value;
                }
                runBorderStyle = rBorder.Val.Value.ToString();
            }
            else
            {
                // 🌟 اصلاح شد: به جای "false"، کلاً تهی (null) می‌شود تا در JSON وارد نشده و باعث خطای پیش‌فرض فلاتر نگردد
                runHasBorders = null;
                runBorderColor = null;
                runBorderStyle = null;
            }

            // 🌟 مسدود کردن نشت رنگ و حاشیه برای کلمات جای‌خالی (BlankWord1)
            bool isParentBlankWord2 = IsTargetStyle(pStyleId, mainPart, "BlankWord2");
            if (isBlankWord1)
            {
                runBorder = null;
            }
            bool bordersAreEqual = (lastTextSpan?.Borders == null && runBorder == null) ||
                                          (lastTextSpan?.Borders != null && runBorder != null &&
                                           lastTextSpan.Borders.Val == runBorder.Val &&
                                           lastTextSpan.Borders.Width == runBorder.Width &&
                                           lastTextSpan.Borders.Color == runBorder.Color);

            if (lastTextSpan != null && lastTextSpan.Type == "text" &&
                lastTextSpan.Url == hyperlinkUrl &&
                lastTextSpan.Markers.SequenceEqual(currentMarkers) &&
                lastTextSpan.FillColor == runShading &&
                lastTextSpan.TextColor == runTextColor &&
                bordersAreEqual)
            {
                lastTextSpan.Content += runText;
            }
            else
            {
                var newTextSpan = new SpanData { Type = "text", Content = runText, Markers = currentMarkers, Url = hyperlinkUrl };

                if (!string.IsNullOrEmpty(runShading)) newTextSpan.FillColor = runShading;
                if (!string.IsNullOrEmpty(runTextColor)) newTextSpan.TextColor = runTextColor;
                if (runBorder != null) newTextSpan.Borders = runBorder; // تزریق مستقیم شیء مشترک بوردر

                paraData.Spans.Add(newTextSpan);
                lastTextSpan = newTextSpan;
            }
        }

        public SpanData ParseTable(Table table, MainDocumentPart mainPart, FontResolver resolver, string outputDir)
        {
            var tableSpan = new SpanData { Type = "table" };
            var tableProps = ExtractTableProperties(table, mainPart);

            if (tableProps.ContainsKey("shading")) tableSpan.FillColor = tableProps["shading"];
            if (tableProps.ContainsKey("tableStyleName")) tableSpan.TableStyleName = tableProps["tableStyleName"];
            if (tableProps.ContainsKey("tableStyleId")) tableSpan.TableStyleId = tableProps["tableStyleId"];
            if (tableProps.ContainsKey("alignment")) tableSpan.TableAlignment = tableProps["alignment"];
            bool hasBorderInfo = tableProps.ContainsKey("borderColor") || tableProps.ContainsKey("borderWidth");

            if (hasBorderInfo)
            {
                // اگر شیء Borders هنوز ساخته نشده، آن را ایجاد می‌کنیم
                if (tableSpan.Borders == null)
                    tableSpan.Borders = new BorderDetail();

                // تنظیم رنگ
                if (tableProps.ContainsKey("borderColor"))
                    tableSpan.Borders.Color = tableProps["borderColor"];

                // تنظیم ضخامت
                if (tableProps.ContainsKey("borderWidth") && double.TryParse(tableProps["borderWidth"], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double bw))
                    tableSpan.Borders.Width = bw;

                // تنظیم یک استایل پیش‌فرض برای کادر تا در فلاتر قابل شناسایی باشد
                if (string.IsNullOrEmpty(tableSpan.Borders.Val))
                    tableSpan.Borders.Val = "single";
            }



            if (tableProps.ContainsKey("tableWidthPercent") && double.TryParse(tableProps["tableWidthPercent"], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double twp))
                tableSpan.TableWidthPercent = twp;

            // 🌟 شبکه اصلی جدول فقط به عنوان پشتیبان (فال‌بک) استخراج می‌شود
            var gridCols = table.Elements<TableGrid>().FirstOrDefault()?.Elements<GridColumn>().ToList();
            List<double> baseGridWidths = new List<double>();

            if (gridCols != null)
            {
                foreach (var col in gridCols)
                {
                    double w = 0;
                    if (col.Width != null && double.TryParse(col.Width.Value, out double parsedWidth)) w = parsedWidth;
                    baseGridWidths.Add(w);
                }
            }

            foreach (var row in table.Elements<TableRow>())
            {
                var rowData = new TableRowData();
                var trPr = row.TableRowProperties;
                if (trPr != null && trPr.Elements<TableHeader>().Any()) rowData.IsHeader = true;

                var cells = row.Elements<TableCell>().ToList();

                // 🌟 جادوی جدید: محاسبه مستقل پهنای سلول‌ها برای همین ردیف (پشتیبانی از ColSpan و عرض متفاوت ردیف‌ها)
                List<double> rowCellWidths = new List<double>();
                double rowTotalWidth = 0;
                int currentGridIndex = 0;

                foreach (var cell in cells)
                {
                    var tcPr = cell.Elements<TableCellProperties>().FirstOrDefault();
                    double cellWidth = 0;

                    int colSpan = 1;
                    if (tcPr?.GridSpan != null && tcPr.GridSpan.Val != null)
                        colSpan = tcPr.GridSpan.Val.Value;

                    // اولویت اول: عرض صریح خود سلول در این ردیف
                    var tcw = tcPr?.TableCellWidth;
                    if (tcw != null && tcw.Width != null && double.TryParse(tcw.Width.Value, out double wVal))
                    {
                        cellWidth = wVal;
                    }

                    // اولویت دوم (فال‌بک): استفاده از شبکه اصلی جدول با احتساب ادغام سلول‌ها
                    if (cellWidth <= 0 && baseGridWidths.Count > 0)
                    {
                        for (int c = 0; c < colSpan; c++)
                        {
                            if (currentGridIndex + c < baseGridWidths.Count)
                                cellWidth += baseGridWidths[currentGridIndex + c];
                        }
                    }

                    rowCellWidths.Add(cellWidth);
                    rowTotalWidth += cellWidth;
                    currentGridIndex += colSpan;
                }

                for (int i = 0; i < cells.Count; i++)
                {
                    var cell = cells[i];
                    var cellData = new TableCellData();
                    ExtractSmartCellPadding(cell, cellData); // 🌟 تزریق پدینگ‌های استخراج‌شده

                    if (rowData.IsHeader) cellData.IsHeaderCell = true;

                    // 🌟 اعمال پهنای اختصاصی محاسبه شده برای همین ردیف
                    if (rowTotalWidth > 0 && i < rowCellWidths.Count)
                        cellData.WidthPercent = Math.Round((rowCellWidths[i] / rowTotalWidth) * 100, 2);

                    var cellProps = ExtractCellProperties(cell);
                    if (cellProps.ContainsKey("shading")) cellData.FillColor = cellProps["shading"];
                    if (cellProps.ContainsKey("vAlign")) cellData.VAlign = cellProps["vAlign"];
                    if (cellProps.ContainsKey("colSpan")) cellData.ColSpan = int.Parse(cellProps["colSpan"]);
                    if (cellProps.ContainsKey("rowMerge")) cellData.RowMerge = cellProps["rowMerge"];
                    cellData.Borders = ExtractSmartCellBorders(cell); // 🌟 تزریق مرزهای استخراج‌شده

                    // 🌟 بررسی تمام فرزندان سلول (پاراگراف و جدول‌های تودرتو)
                    foreach (var element in cell.Elements())
                    {
                        if (element is Paragraph p)
                        {
                            bool isBlankWord2 = IsTargetStyle(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value, mainPart, "BlankWord2");

                            var cellParaDataList = ParseParagraph(p, mainPart, resolver, outputDir, true);

                            cellParaDataList.RemoveAll(pr =>
                                pr.Spans.All(s => s.Type == "text" && string.IsNullOrWhiteSpace(s.Content)) &&
                                pr.StartMs == null);

                            if (cellParaDataList.Count > 0)
                            {
                                cellData.Paragraphs.AddRange(cellParaDataList);
                                if (isBlankWord2)
                                {
                                    foreach (var cp in cellParaDataList) _blankWord2Set.Add(cp);
                                }
                            }
                        }
                        else if (element is Table nestedTable)
                        {
                            // 🌟 استخراج بازگشتی جدول‌های تودرتو
                            var nestedTableSpan = ParseTable(nestedTable, mainPart, resolver, outputDir);
                            var nestedPara = new ParagraphData();
                            nestedPara.Spans.Add(nestedTableSpan);
                            cellData.Paragraphs.Add(nestedPara);
                        }
                    }

                    cellData.Paragraphs = MergeBlankWord2Paragraphs(cellData.Paragraphs);

                    rowData.Cells.Add(cellData);
                }
                tableSpan.TableRows.Add(rowData);
            }

            return tableSpan;
        }

        // 🌟 متد تبدیل مرز خام Word به مدل بهینه‌سازی شده
        private BorderDetail ParseBorder(DocumentFormat.OpenXml.Wordprocessing.BorderType border)
        {
            if (border == null || border.Val == null || border.Val == DocumentFormat.OpenXml.Wordprocessing.BorderValues.None || border.Val == DocumentFormat.OpenXml.Wordprocessing.BorderValues.Nil)
                return null;

            return new BorderDetail
            {
                Val = border.Val.ToString(),
                // در ورد، سایز خطوط بر اساس 1/8 Point ذخیره می‌شود. آن را استاندارد می‌کنیم
                Width = border.Size != null && border.Size.HasValue ? Math.Round((double)border.Size.Value / 8.0, 1) : 1.0,
                Color = border.Color != null && border.Color.Value != "auto" ? border.Color.Value : null
            };
        }
        private BorderDetail ParseBorder(DocumentFormat.OpenXml.Wordprocessing.Border border)
        {
            if (border == null || border.Val == null || border.Val == DocumentFormat.OpenXml.Wordprocessing.BorderValues.None || border.Val == DocumentFormat.OpenXml.Wordprocessing.BorderValues.Nil)
                return null;

            return new BorderDetail
            {
                Val = border.Val.ToString(),
                Width = border.Size != null && border.Size.HasValue ? Math.Round((double)border.Size.Value / 8.0, 1) : 1.0,
                Color = border.Color != null && border.Color.Value != "auto" ? border.Color.Value : null
            };
        }
        // 🌟 هسته اصلی استخراج مرزهای سلول با پشتیبانی کامل از ارث‌بری جدول
        private CellBorders ExtractSmartCellBorders(DocumentFormat.OpenXml.Wordprocessing.TableCell cell)
        {
            var row = cell.Ancestors<DocumentFormat.OpenXml.Wordprocessing.TableRow>().FirstOrDefault();
            var table = row?.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault();

            // تشخیص موقعیت سلول برای استخراج مرزهای خارجی جدول
            bool isFirstRow = table?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().FirstOrDefault() == row;
            bool isLastRow = table?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().LastOrDefault() == row;
            bool isFirstCol = row?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().FirstOrDefault() == cell;
            bool isLastCol = row?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().LastOrDefault() == cell;

            var tcBorders = cell.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCellProperties>().FirstOrDefault()?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCellBorders>().FirstOrDefault();
            var tblBorders = table?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableProperties>().FirstOrDefault()?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableBorders>().FirstOrDefault();

            // قانون ارث‌بری: اول مرز خود سلول، دوم مرز خارجی/داخلی کل جدول
            BorderDetail GetBorder(DocumentFormat.OpenXml.Wordprocessing.BorderType cellB, DocumentFormat.OpenXml.Wordprocessing.BorderType tblOuter, DocumentFormat.OpenXml.Wordprocessing.BorderType tblInner, bool isEdge)
            {
                if (cellB != null) return ParseBorder(cellB);
                if (isEdge && tblOuter != null) return ParseBorder(tblOuter);
                if (!isEdge && tblInner != null) return ParseBorder(tblInner);
                return null;
            }

            // ادامه قطعه کد آخر سی‌شارپ شما جهت تکمیل متد:
            var borders = new CellBorders
            {
                Top = GetBorder(tcBorders?.TopBorder, tblBorders?.TopBorder, tblBorders?.InsideHorizontalBorder, isFirstRow),
                Bottom = GetBorder(tcBorders?.BottomBorder, tblBorders?.BottomBorder, tblBorders?.InsideHorizontalBorder, isLastRow),
                Left = GetBorder(tcBorders?.LeftBorder, tblBorders?.LeftBorder, tblBorders?.InsideVerticalBorder, isFirstCol),
                Right = GetBorder(tcBorders?.RightBorder, tblBorders?.RightBorder, tblBorders?.InsideVerticalBorder, isLastCol)
            };

            return borders;
        }
        // ==========================================
        // متدهای استخراج جزئیات، فونت و استایل‌ها
        // ==========================================
        // 🌟 نگاشتِ درستِ Justification به کدِ تک‌حرفیِ فلاتر (به‌جای Substring که "Both" را "B" می‌کرد)
        private static string MapAlignment(JustificationValues v)
        {
            if (v == JustificationValues.Center) return "C";
            if (v == JustificationValues.Right) return "R";
            if (v == JustificationValues.End) return "R";          // انتهای خط (RTL-relative)
            if (v == JustificationValues.Both) return "J";         // justify
            if (v == JustificationValues.Distribute) return "J";
            return "L";                                            // Left / Start / پیش‌فرض
        }

        // 🌟 کشِ NumberingResolver که با تغییرِ سند (mainPart) خودکار ری‌ست می‌شود
        private NumberingResolver _numbering;
        private MainDocumentPart _numberingPart;
        private NumberingResolver Numbering(MainDocumentPart mainPart)
        {
            if (!ReferenceEquals(_numberingPart, mainPart))
            {
                _numbering = new NumberingResolver(mainPart);
                _numberingPart = mainPart;
            }
            return _numbering;
        }

        private ParagraphData CloneParagraphProperties(ParagraphData source)
        {
            return new ParagraphData
            {
                Direction = source.Direction,
                Alignment = source.Alignment,
                FillColor = source.FillColor,
                SpaceAfter = source.SpaceAfter,
                SpaceBefore = source.SpaceBefore,
                LineSpacing = source.LineSpacing,
                IndentFirstLine = source.IndentFirstLine,
                IndentLeft = source.IndentLeft,
                IndentRight = source.IndentRight,
                Borders = source.Borders != null ? new BorderDetail { Val = source.Borders.Val, Width = source.Borders.Width, Color = source.Borders.Color } : null,
                // 🌟 لیست‌ها: بدون این‌ها، مارکرِ ست‌شده روی basePara هنگام clone گم می‌شد
                // و اصلاً در JSON نمی‌آمد (علتِ نمایش‌نشدنِ شماره‌ها/بولت‌ها).
                ListType = source.ListType,
                ListLevel = source.ListLevel,
                ListMarker = source.ListMarker,
                Spans = new List<SpanData>()
            };
        }

        private SpanData CloneSpan(SpanData source)
        {
            return new SpanData
            {
                Type = source.Type,
                Content = source.Content,
                Markers = source.Markers != null ? new List<string>(source.Markers) : new List<string>(),
                Url = source.Url,

                ImageWidth = source.ImageWidth,   // 🌟 کپی کردن عرض
                ImageHeight = source.ImageHeight, // 🌟 کپی کردن ارتفاع

                FillColor = source.FillColor,
                TextColor = source.TextColor,
                Borders = source.Borders != null ? new BorderDetail { Val = source.Borders.Val, Width = source.Borders.Width, Color = source.Borders.Color } : null,
                FloatPosition = source.FloatPosition,
                TableStyleName = source.TableStyleName,
                TableStyleId = source.TableStyleId,
                TableAlignment = source.TableAlignment,
                TableWidthPercent = source.TableWidthPercent,
                TableRows = source.TableRows
            };
        }

        private bool HasText(ParagraphData para)
        {
            return para.Spans.Any(s => s.Type == "text" && !string.IsNullOrWhiteSpace(s.Content));
        }

        private string GetFontFromStyleId(MainDocumentPart mainPart, string styleId)
        {
            if (mainPart?.StyleDefinitionsPart?.Styles == null || string.IsNullOrEmpty(styleId)) return null;
            var style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleId);
            while (style != null)
            {
                var font = style.StyleRunProperties?.RunFonts?.Ascii?.Value ??
                           style.StyleRunProperties?.RunFonts?.HighAnsi?.Value ??
                           style.StyleRunProperties?.RunFonts?.ComplexScript?.Value;
                if (!string.IsNullOrEmpty(font)) return font;
                var basedOn = style.BasedOn?.Val?.Value;
                if (string.IsNullOrEmpty(basedOn)) break;
                style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == basedOn);
            }
            return null;
        }

        private string GetColorFromStyleId(MainDocumentPart mainPart, string styleId)
        {
            if (mainPart?.StyleDefinitionsPart?.Styles == null || string.IsNullOrEmpty(styleId)) return null;
            var style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleId);
            while (style != null)
            {
                var colorVal = style.StyleRunProperties?.Color?.Val?.Value;
                if (!string.IsNullOrEmpty(colorVal) && colorVal != "auto") return colorVal;
                var basedOn = style.BasedOn?.Val?.Value;
                if (string.IsNullOrEmpty(basedOn)) break;
                style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == basedOn);
            }
            return null;
        }

        private bool IsAllCaps(RunProperties rPr, string runStyleId, string pStyleId, MainDocumentPart mainPart)
        {
            if (rPr?.Caps != null) return rPr.Caps.Val == null || rPr.Caps.Val.Value;
            if (mainPart?.StyleDefinitionsPart?.Styles == null) return false;
            if (!string.IsNullOrEmpty(runStyleId))
            {
                var style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == runStyleId);
                while (style != null)
                {
                    if (style.StyleRunProperties?.Caps != null) return style.StyleRunProperties.Caps.Val == null || style.StyleRunProperties.Caps.Val.Value;
                    var basedOn = style.BasedOn?.Val?.Value;
                    if (string.IsNullOrEmpty(basedOn)) break;
                    style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == basedOn);
                }
            }
            if (!string.IsNullOrEmpty(pStyleId))
            {
                var style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == pStyleId);
                while (style != null)
                {
                    if (style.StyleRunProperties?.Caps != null) return style.StyleRunProperties.Caps.Val == null || style.StyleRunProperties.Caps.Val.Value;
                    var basedOn = style.BasedOn?.Val?.Value;
                    if (string.IsNullOrEmpty(basedOn)) break;
                    style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == basedOn);
                }
            }
            return false;
        }

        private bool IsBold(RunProperties rPr, string runStyleId, string pStyleId, MainDocumentPart mainPart)
        {
            if (rPr?.Bold != null) return rPr.Bold.Val == null || rPr.Bold.Val.Value;
            if (mainPart?.StyleDefinitionsPart?.Styles == null) return false;
            if (!string.IsNullOrEmpty(runStyleId))
            {
                var style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == runStyleId);
                while (style != null)
                {
                    if (style.StyleRunProperties?.Bold != null) return style.StyleRunProperties.Bold.Val == null || style.StyleRunProperties.Bold.Val.Value;
                    var basedOn = style.BasedOn?.Val?.Value;
                    if (string.IsNullOrEmpty(basedOn)) break;
                    style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == basedOn);
                }
            }
            if (!string.IsNullOrEmpty(pStyleId))
            {
                var style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == pStyleId);
                while (style != null)
                {
                    if (style.StyleRunProperties?.Bold != null) return style.StyleRunProperties.Bold.Val == null || style.StyleRunProperties.Bold.Val.Value;
                    var basedOn = style.BasedOn?.Val?.Value;
                    if (string.IsNullOrEmpty(basedOn)) break;
                    style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == basedOn);
                }
            }
            return false;
        }

        private bool IsItalic(RunProperties rPr, string runStyleId, string pStyleId, MainDocumentPart mainPart)
        {
            if (rPr?.Italic != null) return rPr.Italic.Val == null || rPr.Italic.Val.Value;
            if (mainPart?.StyleDefinitionsPart?.Styles == null) return false;
            if (!string.IsNullOrEmpty(runStyleId))
            {
                var style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == runStyleId);
                while (style != null)
                {
                    if (style.StyleRunProperties?.Italic != null) return style.StyleRunProperties.Italic.Val == null || style.StyleRunProperties.Italic.Val.Value;
                    var basedOn = style.BasedOn?.Val?.Value;
                    if (string.IsNullOrEmpty(basedOn)) break;
                    style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == basedOn);
                }
            }
            if (!string.IsNullOrEmpty(pStyleId))
            {
                var style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == pStyleId);
                while (style != null)
                {
                    if (style.StyleRunProperties?.Italic != null) return style.StyleRunProperties.Italic.Val == null || style.StyleRunProperties.Italic.Val.Value;
                    var basedOn = style.BasedOn?.Val?.Value;
                    if (string.IsNullOrEmpty(basedOn)) break;
                    style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == basedOn);
                }
            }
            return false;
        }

        private BorderType GetParagraphBorderFromStyle(MainDocumentPart mainPart, string styleId)
        {
            if (mainPart?.StyleDefinitionsPart?.Styles == null || string.IsNullOrEmpty(styleId)) return null;
            var style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleId);
            while (style != null)
            {
                var pb = style.StyleParagraphProperties?.ParagraphBorders;
                if (pb != null)
                {
                    return (BorderType)pb.TopBorder ?? (BorderType)pb.BottomBorder ?? (BorderType)pb.LeftBorder ?? (BorderType)pb.RightBorder;
                }
                var basedOn = style.BasedOn?.Val?.Value;
                if (string.IsNullOrEmpty(basedOn)) break;
                style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == basedOn);
            }
            return null;
        }

        private Border GetRunBorderFromStyle(MainDocumentPart mainPart, string styleId)
        {
            if (mainPart?.StyleDefinitionsPart?.Styles == null || string.IsNullOrEmpty(styleId)) return null;
            var style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleId);
            while (style != null)
            {
                if (style.StyleRunProperties?.Border != null) return style.StyleRunProperties.Border;
                var basedOn = style.BasedOn?.Val?.Value;
                if (string.IsNullOrEmpty(basedOn)) break;
                style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == basedOn);
            }
            return null;
        }

        private string GetFontSizeFromStyleId(MainDocumentPart mainPart, string styleId)
        {
            if (mainPart?.StyleDefinitionsPart?.Styles == null || string.IsNullOrEmpty(styleId)) return null;
            var style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleId);
            while (style != null)
            {
                var fontSize = style.StyleRunProperties?.FontSize?.Val?.Value;
                if (!string.IsNullOrEmpty(fontSize)) return fontSize;
                var basedOn = style.BasedOn?.Val?.Value;
                if (string.IsNullOrEmpty(basedOn)) break;
                style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == basedOn);
            }
            return null;
        }

        private double? GetSpacingFromStyle(MainDocumentPart mainPart, string styleId, bool isAfter)
        {
            if (mainPart?.StyleDefinitionsPart?.Styles == null || string.IsNullOrEmpty(styleId)) return null;

            var style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleId);
            while (style != null)
            {
                var spacing = style.StyleParagraphProperties?.SpacingBetweenLines;
                if (spacing != null)
                {
                    var targetAttr = isAfter ? spacing.After : spacing.Before;
                    if (targetAttr != null && double.TryParse(targetAttr.Value, out double twips))
                    {
                        return twips / 20.0;
                    }
                }
                var basedOn = style.BasedOn?.Val?.Value;
                if (string.IsNullOrEmpty(basedOn)) break;
                style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == basedOn);
            }
            return null;
        }

        private double? GetDefaultSpacing(MainDocumentPart mainPart, bool isAfter)
        {
            var docDefaults = mainPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
            var pPrDefault = docDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
            var spacing = pPrDefault?.SpacingBetweenLines;
            if (spacing != null)
            {
                var targetAttr = isAfter ? spacing.After : spacing.Before;
                if (targetAttr != null && double.TryParse(targetAttr.Value, out double twips))
                {
                    return twips / 20.0;
                }
            }
            return null;
        }

        private List<string> ExtractRunMarkers(Run run, ParagraphProperties pPr, MainDocumentPart mainPart)
        {
            var markers = new List<string>();
            var rPr = run.RunProperties;

            var runStyleId = rPr?.RunStyle?.Val?.Value;
            var pStyleId = pPr?.ParagraphStyleId?.Val?.Value;

            if (IsBold(rPr, runStyleId, pStyleId, mainPart)) markers.Add("b");
            if (IsItalic(rPr, runStyleId, pStyleId, mainPart)) markers.Add("i");

            if (rPr?.Underline != null)
            {
                var uVal = rPr.Underline.Val?.Value ?? UnderlineValues.Single;
                if (uVal != UnderlineValues.None) markers.Add("u");
            }

            var shading = rPr?.Shading?.Fill?.Value;
            if (shading != null && shading != "auto") markers.Add($"bg:{shading}");

            string fontName = null;
            if (rPr?.RunFonts != null)
            {
                fontName = rPr.RunFonts.Ascii?.Value ?? rPr.RunFonts.ComplexScript?.Value ?? rPr.RunFonts.HighAnsi?.Value;
                if (string.IsNullOrEmpty(fontName))
                {
                    var themeFont = rPr.RunFonts.AsciiTheme ?? rPr.RunFonts.ComplexScriptTheme;
                    if (themeFont != null)
                    {
                        string themeVal = themeFont.Value.ToString().ToLower();
                        if (themeVal.Contains("major")) fontName = "Times New Roman";
                        else if (themeVal.Contains("minor")) fontName = "Calibri";
                    }
                }
            }

            if (string.IsNullOrEmpty(fontName) && !string.IsNullOrEmpty(runStyleId)) fontName = GetFontFromStyleId(mainPart, runStyleId);
            if (string.IsNullOrEmpty(fontName) && pPr?.ParagraphStyleId?.Val?.Value != null)
            {
                fontName = GetFontFromStyleId(mainPart, pPr.ParagraphStyleId.Val.Value);
                if (string.IsNullOrEmpty(fontName))
                {
                    string pStyle = pPr.ParagraphStyleId.Val.Value.ToLower();
                    if (pStyle.Contains("heading") || pStyle.Contains("title")) fontName = "Times New Roman";
                    else if (pStyle.Contains("normal") || pStyle.Contains("body") || pStyle.Contains("list")) fontName = "Calibri";
                }
            }

            if (string.IsNullOrEmpty(fontName)) fontName = "Source Sans 3";
            markers.Add($"fn:{fontName}");

            string fontSizeStr = null;
            if (rPr?.FontSize?.Val?.Value != null) fontSizeStr = rPr.FontSize.Val.Value;
            if (string.IsNullOrEmpty(fontSizeStr))
            {
                if (!string.IsNullOrEmpty(runStyleId)) fontSizeStr = GetFontSizeFromStyleId(mainPart, runStyleId);
                if (string.IsNullOrEmpty(fontSizeStr) && pPr?.ParagraphStyleId?.Val?.Value != null) fontSizeStr = GetFontSizeFromStyleId(mainPart, pPr.ParagraphStyleId.Val.Value);
                if (string.IsNullOrEmpty(fontSizeStr) && pPr?.ParagraphMarkRunProperties != null)
                {
                    var pSize = pPr.ParagraphMarkRunProperties.GetFirstChild<FontSize>();
                    if (pSize?.Val?.Value != null) fontSizeStr = pSize.Val.Value;
                }
            }

            if (!string.IsNullOrEmpty(fontSizeStr)) markers.Add($"sz:{fontSizeStr}");
            return markers;
        }

        private (string FileName, int Width, int Height)? ExtractAndSaveImage(DocumentFormat.OpenXml.Wordprocessing.Drawing drawing, MainDocumentPart mainPart, string outputDir)
        {
            var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
            if (blip == null || string.IsNullOrEmpty(blip.Embed?.Value)) return null;

            var part = mainPart.GetPartById(blip.Embed.Value);
            if (part == null) return null;

            string imageDir = Path.Combine(outputDir, "images");
            if (!Directory.Exists(imageDir)) Directory.CreateDirectory(imageDir);

            string extension = part.Uri.ToString().Split('.').Last();
            string fileName = $"p{_currentPage}_s{_currentSection}_img{_imageCounter}.{extension}";
            string filePath = Path.Combine(imageDir, fileName);
            _imageCounter++;

            using (var stream = part.GetStream())
            using (var originalImage = System.Drawing.Image.FromStream(stream))
            {
                var srcRect = drawing.Descendants<DocumentFormat.OpenXml.Drawing.SourceRectangle>().FirstOrDefault();
                if (srcRect != null)
                {
                    double leftPct = (srcRect.Left != null ? srcRect.Left.Value : 0) / 100000.0;
                    double topPct = (srcRect.Top != null ? srcRect.Top.Value : 0) / 100000.0;
                    double rightPct = (srcRect.Right != null ? srcRect.Right.Value : 0) / 100000.0;
                    double bottomPct = (srcRect.Bottom != null ? srcRect.Bottom.Value : 0) / 100000.0;

                    int origWidth = originalImage.Width;
                    int origHeight = originalImage.Height;

                    int x = (int)(origWidth * leftPct);
                    int y = (int)(origHeight * topPct);
                    int width = origWidth - x - (int)(origWidth * rightPct);
                    int height = origHeight - y - (int)(origHeight * bottomPct);

                    if (width > 0 && height > 0)
                    {
                        using (var bitmap = new System.Drawing.Bitmap(originalImage))
                        {
                            var cropRect = new System.Drawing.Rectangle(x, y, width, height);
                            using (var croppedBitmap = bitmap.Clone(cropRect, bitmap.PixelFormat))
                            {
                                croppedBitmap.Save(filePath);
                            }
                        }
                        // 🌟 بازگرداندن ابعاد تصویر کراپ شده
                        return (fileName, width, height);
                    }
                }
                originalImage.Save(filePath);
                // 🌟 بازگرداندن ابعاد تصویر اصلی
                return (fileName, originalImage.Width, originalImage.Height);
            }
        }

        private string GetHyperlinkUrl(MainDocumentPart mainPart, string id)
        {
            if (mainPart == null || string.IsNullOrEmpty(id)) return null;
            try
            {
                var rel = mainPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == id);
                return rel?.Uri?.ToString();
            }
            catch
            {
                return null;
            }
        }
        private Dictionary<string, string> ExtractTableProperties(Table table, MainDocumentPart mainPart)
        {
            var props = new Dictionary<string, string>();
            var tblPr = table.Elements<TableProperties>().FirstOrDefault();

            if (tblPr == null) return props;

            var styleId = tblPr.TableStyle?.Val?.Value;
            if (!string.IsNullOrEmpty(styleId))
            {
                props.Add("tableStyleId", styleId);

                var stylesPart = mainPart?.StyleDefinitionsPart;
                if (stylesPart?.Styles != null)
                {
                    var style = stylesPart.Styles.Elements<Style>()
                        .FirstOrDefault(s => s.StyleId == styleId && s.Type == StyleValues.Table);

                    var friendlyName = style?.StyleName?.Val?.Value;
                    if (!string.IsNullOrEmpty(friendlyName))
                    {
                        props.Add("tableStyleName", friendlyName);
                    }
                }
            }

            var alignment = tblPr.TableJustification?.Val?.Value.ToString();
            if (!string.IsNullOrEmpty(alignment)) props.Add("alignment", alignment.ToLower());

            var shading = tblPr.Shading?.Fill?.Value;
            if (!string.IsNullOrEmpty(shading) && shading != "auto") props.Add("shading", shading);

            double totalWidthPercent = 100;
            bool widthFoundFromProp = false;

            if (tblPr.TableWidth != null)
            {
                var widthUnit = tblPr.TableWidth.Type?.Value;
                if (double.TryParse(tblPr.TableWidth.Width?.Value, out double wVal))
                {
                    if (widthUnit == TableWidthUnitValues.Pct)
                    {
                        totalWidthPercent = wVal / 50.0;
                        widthFoundFromProp = true;
                    }
                    else if (widthUnit == TableWidthUnitValues.Dxa && wVal > 0)
                    {
                        totalWidthPercent = (wVal / 9360.0) * 100;
                        widthFoundFromProp = true;
                    }
                }
            }

            var tblGrid = table.Elements<TableGrid>().FirstOrDefault();
            if (tblGrid != null)
            {
                double sumGridWidth = 0;
                foreach (var col in tblGrid.Elements<GridColumn>())
                {
                    if (col.Width != null && double.TryParse(col.Width.Value, out double cWidth))
                    {
                        sumGridWidth += cWidth;
                    }
                }

                if (sumGridWidth > 0)
                {
                    double gridPercent = (sumGridWidth / 9360.0) * 100;
                    if (!widthFoundFromProp || gridPercent < totalWidthPercent)
                    {
                        totalWidthPercent = gridPercent;
                    }
                }
            }

            if (totalWidthPercent > 100) totalWidthPercent = 100;
            if (totalWidthPercent < 10) totalWidthPercent = 10;
            props.Add("tableWidthPercent", totalWidthPercent.ToString(System.Globalization.CultureInfo.InvariantCulture));

            TableBorders inlineBorders = tblPr.TableBorders;
            TableBorders styleBorders = null;

            if (!string.IsNullOrEmpty(styleId) && mainPart?.StyleDefinitionsPart?.Styles != null)
            {
                var style = mainPart.StyleDefinitionsPart.Styles.Elements<Style>()
                    .FirstOrDefault(s => s.StyleId == styleId && s.Type == StyleValues.Table);
                styleBorders = style?.StyleTableProperties?.TableBorders;
            }

            BorderType border = null;

            if (inlineBorders != null)
            {
                border = (BorderType)inlineBorders.TopBorder ?? (BorderType)inlineBorders.LeftBorder ??
                         (BorderType)inlineBorders.BottomBorder ?? (BorderType)inlineBorders.RightBorder ??
                         (BorderType)inlineBorders.InsideHorizontalBorder ?? (BorderType)inlineBorders.InsideVerticalBorder;
            }

            if (border == null && styleBorders != null)
            {
                border = (BorderType)styleBorders.TopBorder ?? (BorderType)styleBorders.LeftBorder ??
                         (BorderType)styleBorders.BottomBorder ?? (BorderType)styleBorders.RightBorder ??
                         (BorderType)styleBorders.InsideHorizontalBorder ?? (BorderType)styleBorders.InsideVerticalBorder;
            }

            if (border != null)
            {
                var hasBorders = border.Val != null && border.Val.Value != BorderValues.None && border.Val.Value != BorderValues.Nil;
                props.Add("hasBorders", hasBorders.ToString().ToLower());

                if (hasBorders)
                {
                    if (border.Color != null && border.Color.Value != "auto")
                    {
                        props.Add("borderColor", border.Color.Value);
                    }
                    if (border.Size != null)
                    {
                        double widthPt = border.Size.Value / 8.0;
                        props.Add("borderWidth", widthPt.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    }
                }
            }
            else
            {
                if (styleId == "BorderedTable")
                {
                    props.Add("hasBorders", "true");
                    props.Add("borderColor", "000000");
                    props.Add("borderWidth", "1.0");
                }
                else
                {
                    props.Add("hasBorders", "false");
                }
            }

            return props;
        }

        private Dictionary<string, string> ExtractCellProperties(TableCell cell)
        {
            var props = new Dictionary<string, string>();
            var tcPr = cell.Elements<TableCellProperties>().FirstOrDefault();

            if (tcPr == null) return props;

            var shading = tcPr.Shading?.Fill?.Value;
            if (!string.IsNullOrEmpty(shading) && shading != "auto") props.Add("shading", shading);

            var vAlign = tcPr.TableCellVerticalAlignment?.Val?.Value.ToString();
            if (!string.IsNullOrEmpty(vAlign)) props.Add("vAlign", vAlign.ToLower());

            var gridSpan = tcPr.GridSpan?.Val?.Value;
            if (gridSpan != null && gridSpan > 1) props.Add("colSpan", gridSpan.ToString());

            var vMergeVal = tcPr.VerticalMerge?.Val?.Value.ToString();
            var vMerge = vMergeVal ?? (tcPr.VerticalMerge != null ? "continue" : null);

            if (!string.IsNullOrEmpty(vMerge)) props.Add("rowMerge", vMerge.ToLower());

            var width = tcPr.TableCellWidth?.Type?.Value == TableWidthUnitValues.Dxa ? tcPr.TableCellWidth?.Width?.Value : null;
            if (!string.IsNullOrEmpty(width)) props.Add("widthDxa", width);

            return props;
        }
        private void ExtractSmartCellPadding(DocumentFormat.OpenXml.Wordprocessing.TableCell cell, TableCellData cellData)
        {
            var row = cell.Ancestors<DocumentFormat.OpenXml.Wordprocessing.TableRow>().FirstOrDefault();
            var table = row?.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault();

            var tcMargin = cell.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCellProperties>().FirstOrDefault()?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCellMargin>().FirstOrDefault();
            var tblMargin = table?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableProperties>().FirstOrDefault()?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCellMarginDefault>().FirstOrDefault();

            double? GetMargin<T>(T cellW, T tblW)
                where T : DocumentFormat.OpenXml.Wordprocessing.TableWidthType
            {
                var target = cellW ?? tblW;

                if (target?.Width != null &&
                    double.TryParse(target.Width.Value, out double twips))
                {
                    return twips / 20.0;
                }

                return null;
            }

            cellData.PaddingTop = GetMargin(
                tcMargin?.GetFirstChild<TopMargin>(),
                tblMargin?.GetFirstChild<TopMargin>());

            cellData.PaddingBottom = GetMargin(
                tcMargin?.GetFirstChild<BottomMargin>(),
                tblMargin?.GetFirstChild<BottomMargin>());

            cellData.PaddingLeft = GetMargin(
                tcMargin?.GetFirstChild<StartMargin>(),
                tblMargin?.GetFirstChild<StartMargin>());

            cellData.PaddingRight = GetMargin(
                tcMargin?.GetFirstChild<EndMargin>(),
                tblMargin?.GetFirstChild<EndMargin>());
        }
        private List<string> ExtractCellBorders(TableCell cell)
        {
            var borders = new List<string>();
            var tcPr = cell.Elements<TableCellProperties>().FirstOrDefault();

            if (tcPr != null)
            {
                var tcBorders = tcPr.Elements<TableCellBorders>().FirstOrDefault();
                // ۱. بررسی مرزهای اختصاصی خود سلول
                if (tcBorders != null)
                {
                    if (tcBorders.TopBorder != null && tcBorders.TopBorder.Val != BorderValues.None) borders.Add("top");
                    if (tcBorders.BottomBorder != null && tcBorders.BottomBorder.Val != BorderValues.None) borders.Add("bottom");
                    if (tcBorders.LeftBorder != null && tcBorders.LeftBorder.Val != BorderValues.None) borders.Add("left");
                    if (tcBorders.RightBorder != null && tcBorders.RightBorder.Val != BorderValues.None) borders.Add("right");

                    return borders; // اگر سلول مرز اختصاصی داشت، همین را برمی‌گردانیم
                }
            }

            // ۲. فالبک (Fallback): اگر سلول مرز اختصاصی نداشت، بررسی مرزهای کل جدول
            var table = cell.Ancestors<Table>().FirstOrDefault();
            var tblPr = table?.Elements<TableProperties>().FirstOrDefault();
            var tblBorders = tblPr?.Elements<TableBorders>().FirstOrDefault();

            if (tblBorders != null)
            {
                if (tblBorders.TopBorder != null && tblBorders.TopBorder.Val != BorderValues.None) borders.Add("top");
                if (tblBorders.BottomBorder != null && tblBorders.BottomBorder.Val != BorderValues.None) borders.Add("bottom");
                if (tblBorders.LeftBorder != null && tblBorders.LeftBorder.Val != BorderValues.None) borders.Add("left");
                if (tblBorders.RightBorder != null && tblBorders.RightBorder.Val != BorderValues.None) borders.Add("right");

                // مرزهای داخلی جدول (اگر جدول خطوط شطرنجی داخلی داشته باشد)
                if (tblBorders.InsideHorizontalBorder != null && tblBorders.InsideHorizontalBorder.Val != BorderValues.None)
                {
                    borders.Add("top");
                    borders.Add("bottom");
                }
                if (tblBorders.InsideVerticalBorder != null && tblBorders.InsideVerticalBorder.Val != BorderValues.None)
                {
                    borders.Add("left");
                    borders.Add("right");
                }
            }

            return borders.Distinct().ToList(); // حذف جهت‌های تکراری احتمالی
        }
    }


    // 🌟 ساختار یکپارچه خروجی JSON (نرمال‌سازی داده‌ها)
    public class BookExportData
    {
        public List<PageData> Pages { get; set; } = new List<PageData>();
        public List<ParagraphData> AudioScripts { get; set; } = new List<ParagraphData>();
    }
}