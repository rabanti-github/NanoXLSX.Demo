using NanoXLSX;
using NanoXLSX.Extensions;
using NanoXLSX.Styles;

namespace PicoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows inline formatting with phonetic runs (important for East Asian languages like Japanese)
    /// </summary>
    public class InlineFormattingPhonetic
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("InlineFormattingPhonetic.xlsx", "Sheet1");            // Create new workbook

            // Example 1: Japanese text with Hiragana phonetics
            Font phoneticFont = new Font { Size = 8 };                                              // Smaller font for phonetic text

            FormattedTextBuilder builder1 = new FormattedTextBuilder();
            builder1.AddRun("東京");                                                                 // Add base text (Tokyo in Kanji)
            builder1.AddPhoneticRun("とうきょう", 0, 2);                                              // Add phonetic guide (Hiragana)
            builder1.SetPhoneticProperties(phoneticFont,
                PhoneticRun.PhoneticType.Hiragana,
                PhoneticRun.PhoneticAlignment.Center);
            workbook.CurrentWorksheet.AddFormattedTextCell(builder1.Build(), 0, 0);                 // Add to cell A1

            // Example 2: Japanese text with Katakana phonetics
            FormattedTextBuilder builder2 = new FormattedTextBuilder();
            builder2.AddRun("日本語");                                                               // Japanese language (in Kanji)
            builder2.AddPhoneticRun("ニホンゴ", 0, 3);                                               // Katakana phonetic guide
            builder2.SetPhoneticProperties(phoneticFont,
                PhoneticRun.PhoneticType.FullwidthKatakana,
                PhoneticRun.PhoneticAlignment.Left);
            workbook.CurrentWorksheet.AddFormattedTextCell(builder2.Build(), 0, 1);                 // Add to cell A2

            // Example 3: Mixed text with multiple phonetic runs
            FormattedTextBuilder builder3 = new FormattedTextBuilder();
            builder3.AddRun("私");                                                                   // "I/me" in Kanji
            builder3.AddPhoneticRun("わたし", 0, 1);                                                  // Phonetic for first character
            builder3.AddRun("は");                                                                   // Topic particle (Hiragana, no phonetic needed)
            builder3.AddRun("学生");                                                                 // "student" in Kanji
            builder3.AddPhoneticRun("がくせい", 2, 2);                                                // Phonetic starting at position 2, length 2
            builder3.SetPhoneticProperties(phoneticFont,
                PhoneticRun.PhoneticType.Hiragana,
                PhoneticRun.PhoneticAlignment.Center);
            workbook.CurrentWorksheet.AddFormattedTextCell(builder3.Build(), 0, 2);                 // Add to cell A3

            workbook.Save();                                                                        // Save the workbook
        }

        private InlineFormattingPhonetic()
        {
            // Do not instantiate
        }
    }
}
