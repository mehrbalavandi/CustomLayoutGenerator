// توسط ChatGPT
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
namespace WordToJsonParser
{
    /*
        public class FontResolver
        {
            private readonly WordprocessingDocument _doc;

            public FontResolver(
                WordprocessingDocument doc)
            {
                _doc = doc;
            }
            public string? GetEffectiveFontName(
                Run run)
            {
                string runText = run.InnerText ?? "";

                bool isComplexScript = ContainsPersianOrArabic(runText);

                // 1- Run Properties
                var font = ResolveFontFromRunProperties(
                    run.RunProperties,
                    isComplexScript);

                if (!string.IsNullOrWhiteSpace(font))
                    return font;

                // 2- Character Style
                var charStyleId =
                    run.RunProperties?
                       .RunStyle?
                       .Val?
                       .Value;

                if (!string.IsNullOrEmpty(charStyleId))
                {
                    font = ResolveFontFromStyle(
                        charStyleId,
                        isComplexScript);

                    if (!string.IsNullOrWhiteSpace(font))
                        return font;
                }

                // 3- Paragraph Style
                var paragraph =
                    run.Ancestors<Paragraph>()
                       .FirstOrDefault();

                var paraStyleId =
                    paragraph?
                       .ParagraphProperties?
                       .ParagraphStyleId?
                       .Val?
                       .Value;

                if (!string.IsNullOrEmpty(paraStyleId))
                {
                    font = ResolveFontFromStyle(
                        paraStyleId,
                        isComplexScript);

                    if (!string.IsNullOrWhiteSpace(font))
                        return font;
                }

                // 4- Document Defaults
                font = ResolveFromDocumentDefaults(
                    isComplexScript);

                return font;
            }

            private string? ResolveFontFromStyle(
                string styleId,
                bool isComplexScript)
            {
                var styles =
                    _doc.MainDocumentPart?
                       .StyleDefinitionsPart?
                       .Styles;

                if (styles == null)
                    return null;

                var style =
                    styles.Elements<Style>()
                          .FirstOrDefault(x => x.StyleId == styleId);

                while (style != null)
                {
                    var font = ResolveFontFromRunProperties(
                        style.StyleRunProperties,
                        isComplexScript);

                    if (!string.IsNullOrWhiteSpace(font))
                        return font;

                    var basedOn =
                        style.BasedOn?
                             .Val?
                             .Value;

                    if (string.IsNullOrEmpty(basedOn))
                        break;

                    style =
                        styles.Elements<Style>()
                              .FirstOrDefault(x => x.StyleId == basedOn);
                }

                return null;
            }

            private string? ResolveFromDocumentDefaults(
                bool isComplexScript)
            {
                var rPr =
                    _doc.MainDocumentPart?
                       .StyleDefinitionsPart?
                       .Styles?
                       .DocDefaults?
                       .RunPropertiesDefault?
                       .RunPropertiesBaseStyle;

                return ResolveFontFromRunProperties(
                    rPr,
                    isComplexScript);
            }

            private string? ResolveFontFromRunProperties(
        OpenXmlCompositeElement? rPr,
        bool isComplexScript)
            {
                var fonts = rPr?.GetFirstChild<RunFonts>();

                if (fonts == null)
                    return null;

                if (isComplexScript)
                {
                    if (!string.IsNullOrWhiteSpace(fonts.ComplexScript?.Value))
                        return fonts.ComplexScript!.Value;

                    if (fonts.ComplexScriptTheme != null)
                        return ResolveThemeFont(
                            fonts.ComplexScriptTheme.Value.ToString());
                }

                if (!string.IsNullOrWhiteSpace(fonts.Ascii?.Value))
                    return fonts.Ascii!.Value;

                if (!string.IsNullOrWhiteSpace(fonts.HighAnsi?.Value))
                    return fonts.HighAnsi!.Value;

                if (fonts.AsciiTheme != null)
                {
                    Console.WriteLine(rPr.InnerText);
                    Console.WriteLine(rPr.OuterXml);

                    return ResolveThemeFont(
                        fonts.AsciiTheme.Value.ToString());
                }

                if (fonts.HighAnsiTheme != null)
                {
                    return ResolveThemeFont(
                        fonts.HighAnsiTheme.Value.ToString());
                }

                return null;
            }

            private string? ResolveThemeFont(string themeValue)
            {
                var theme =
                    _doc.MainDocumentPart?
                        .ThemePart?
                        .Theme;

                var fontScheme =
                    theme?.ThemeElements?.FontScheme;

                if (fontScheme == null)
                    return null;

                switch (themeValue)
                {
                    case "majorAscii":
                    case "majorHAnsi":

                        return fontScheme.MajorFont?
                            .LatinFont?
                            .Typeface?
                            .Value;

                    case "minorAscii":
                    case "minorHAnsi":

                        return fontScheme.MinorFont?
                            .LatinFont?
                            .Typeface?
                            .Value;

                    case "majorBidi":
                        {
                            var cs =
                                fontScheme.MajorFont?
                                    .ComplexScriptFont?
                                    .Typeface?
                                    .Value;

                            if (!string.IsNullOrWhiteSpace(cs))
                                return cs;

                            var arab =
                                fontScheme.MajorFont?
                                    .Elements<DocumentFormat.OpenXml.Drawing.SupplementalFont>()
                                    .FirstOrDefault(x => x.Script?.Value == "Arab");

                            return arab?.Typeface?.Value;
                        }

                    case "minorBidi":
                        {
                            var cs =
                                fontScheme.MinorFont?
                                    .ComplexScriptFont?
                                    .Typeface?
                                    .Value;

                            if (!string.IsNullOrWhiteSpace(cs))
                                return cs;

                            var arab =
                                fontScheme.MinorFont?
                                    .Elements<DocumentFormat.OpenXml.Drawing.SupplementalFont>()
                                    .FirstOrDefault(x => x.Script?.Value == "Arab");

                            return arab?.Typeface?.Value;
                        }
                }

                return null;
            }
            private bool ContainsPersianOrArabic(string text)
            {
                return text.Any(ch =>
                    (ch >= 0x0600 && ch <= 0x06FF) ||
                    (ch >= 0x0750 && ch <= 0x077F) ||
                    (ch >= 0x08A0 && ch <= 0x08FF) ||
                    (ch >= 0xFB50 && ch <= 0xFDFF) ||
                    (ch >= 0xFE70 && ch <= 0xFEFF));
            }
        }

    */

    public class FontResolver
    {
        private readonly WordprocessingDocument _doc;

        public FontResolver(WordprocessingDocument doc)
        {
            _doc = doc;
        }

        public string? GetEffectiveFontName(Run run)
        {
            // 1- فونت مستقیم روی Run
            var font = GetFontFromRunFonts(
                run.RunProperties?.RunFonts);

            if (!string.IsNullOrWhiteSpace(font))
                return font;

            // 2- Character Style
            var charStyleId =
                run.RunProperties?
                   .RunStyle?
                   .Val?
                   .Value;

            if (!string.IsNullOrEmpty(charStyleId))
            {
                font = GetFontFromStyle(charStyleId);

                if (!string.IsNullOrWhiteSpace(font))
                    return font;
            }

            // 3- Paragraph Style
            var paragraph =
                run.Ancestors<Paragraph>()
                   .FirstOrDefault();

            var paraStyleId =
                paragraph?
                   .ParagraphProperties?
                   .ParagraphStyleId?
                   .Val?
                   .Value;

            if (!string.IsNullOrEmpty(paraStyleId))
            {
                font = GetFontFromStyle(paraStyleId);

                if (!string.IsNullOrWhiteSpace(font))
                    return font;
            }

            // 4- Document Defaults
            font = GetFontFromDocumentDefaults();

            if (!string.IsNullOrWhiteSpace(font))
                return font;

            // 5- آخرین fallback
            return "Calibri";
        }

        private string? GetFontFromStyle(string styleId)
        {
            var styles =
                _doc.MainDocumentPart?
                    .StyleDefinitionsPart?
                    .Styles;

            if (styles == null)
                return null;

            var style =
                styles.Elements<Style>()
                      .FirstOrDefault(x => x.StyleId == styleId);

            while (style != null)
            {
                var font =
                    GetFontFromRunFonts(
                        style.StyleRunProperties?.RunFonts);

                if (!string.IsNullOrWhiteSpace(font))
                    return font;

                var basedOn =
                    style.BasedOn?
                         .Val?
                         .Value;

                if (string.IsNullOrEmpty(basedOn))
                    break;

                style =
                    styles.Elements<Style>()
                          .FirstOrDefault(x => x.StyleId == basedOn);
            }

            return null;
        }

        private string? GetFontFromDocumentDefaults()
        {
            var runDefaults =
                _doc.MainDocumentPart?
                    .StyleDefinitionsPart?
                    .Styles?
                    .DocDefaults?
                    .RunPropertiesDefault?
                    .RunPropertiesBaseStyle;

            return GetFontFromRunFonts(
                runDefaults?.RunFonts);
        }

        private string? GetFontFromRunFonts(
            RunFonts? fonts)
        {
            if (fonts == null)
                return null;

            // فونت صریح
            if (!string.IsNullOrWhiteSpace(fonts.Ascii?.Value))
                return fonts.Ascii.Value;

            if (!string.IsNullOrWhiteSpace(fonts.HighAnsi?.Value))
                return fonts.HighAnsi.Value;

            // Theme Font
            if (fonts.AsciiTheme != null)
                return ResolveThemeFont(
                    fonts.AsciiTheme.InnerText);

            if (fonts.HighAnsiTheme != null)
                return ResolveThemeFont(
                    fonts.HighAnsiTheme.InnerText);

            return null;
        }

        private string? ResolveThemeFont(
            string themeValue)
        {
            var fontScheme =
                _doc.MainDocumentPart?
                    .ThemePart?
                    .Theme?
                    .ThemeElements?
                    .FontScheme;

            if (fontScheme == null)
                return null;

            switch (themeValue)
            {
                case "majorAscii":
                case "majorHAnsi":
                case "majorBidi":
                    return fontScheme
                        .MajorFont?
                        .LatinFont?
                        .Typeface?
                        .Value;

                case "minorAscii":
                case "minorHAnsi":
                case "minorBidi":
                    return fontScheme
                        .MinorFont?
                        .LatinFont?
                        .Typeface?
                        .Value;
            }

            return null;
        }
    }
}
// استفاده
/*
string? fontName =
    FontResolver.GetEffectiveFontName(
        wordDocument,
        run);

Console.WriteLine(fontName);
*/