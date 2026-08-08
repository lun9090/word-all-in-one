using Microsoft.Office.Interop.Word;
// 添加此行以引入WdTexture枚举
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Text;
using Document = Microsoft.Office.Interop.Word.Document;

namespace 李艇的办公助手
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) { }

        // ==================== 工具方法 ====================
        private float ConvertMillimetersToPoints(double millimeters)
        {
            return (float)(millimeters * 2.83465);
        }

        // 屏幕更新辅助
        private void WithScreenUpdatingDisabled(Action action)
        {
            Application app = Globals.ThisAddIn.Application;
            bool original = app.ScreenUpdating;
            try
            {
                app.ScreenUpdating = false;
                action();
            }
            finally
            {
                app.ScreenUpdating = original;
            }
        }

        // 同步 Selection（插入点）格式使后续输入继承 Range 的格式
        private void SyncSelectionToRange(Selection sel)
        {
            if (sel == null) return;

            Range r = sel.Range;
            Font rf = null;
            ParagraphFormat rpf = null;
            Font sf = null;
            ParagraphFormat spf = null;
            try
            {
                rf = r.Font;
                rpf = r.ParagraphFormat;

                sf = sel.Font;
                spf = sel.ParagraphFormat;

                // 字体：同步所有字体槽与常用属性，确保英文/标点/西文也使用相同字体
                if (rf != null && sf != null)
                {
                    try { sf.Name = rf.Name; } catch { }
                    try { sf.NameFarEast = rf.NameFarEast; } catch { }
                    try { sf.NameAscii = rf.NameAscii; } catch { }
                    try { sf.NameOther = rf.NameOther; } catch { }
                    try { sf.NameBi = rf.NameBi; } catch { }
                    try { sf.Size = rf.Size; } catch { }
                    try { sf.Bold = rf.Bold; } catch { }
                    try { sf.Italic = rf.Italic; } catch { }
                    try { sf.Color = rf.Color; } catch { }
                }

                // 段落格式（常用项）
                if (rpf != null && spf != null)
                {
                    try { spf.LineSpacingRule = rpf.LineSpacingRule; } catch { }
                    try { spf.LineSpacing = rpf.LineSpacing; } catch { }
                    try { spf.Alignment = rpf.Alignment; } catch { }
                    try { spf.SpaceBefore = rpf.SpaceBefore; } catch { }
                    try { spf.SpaceAfter = rpf.SpaceAfter; } catch { }
                    try { spf.CharacterUnitFirstLineIndent = rpf.CharacterUnitFirstLineIndent; } catch { }
                    try { spf.FirstLineIndent = rpf.FirstLineIndent; } catch { }
                    try { spf.LeftIndent = rpf.LeftIndent; } catch { }
                    try { spf.CharacterUnitLeftIndent = rpf.CharacterUnitLeftIndent; } catch { }
                    try { spf.OutlineLevel = rpf.OutlineLevel; } catch { }
                }
            }
            finally
            {
                if (spf != null) Marshal.ReleaseComObject(spf);
                if (sf != null) Marshal.ReleaseComObject(sf);
                if (rpf != null) Marshal.ReleaseComObject(rpf);
                if (rf != null) Marshal.ReleaseComObject(rf);
            }
        }

        // ==================== 公用段落样式 ====================
        // 新增：接受 Range 的高性能实现（缓存 COM 对象、批量设置）
        private void ApplyBasicParagraphStyle(Range range, string fontName, float fontSize, float lineSpacing)
        {
            if (range == null) return;

            // 缓存 COM 对象，减少往返
            Font f = null;
            ParagraphFormat pf = null;
            try
            {
                // 将 range.ClearFormatting(); 替换为段落和字体格式的分别清除
                // 原代码：range.ClearFormatting();
                range.Font.Reset();
                range.ParagraphFormat.Reset();

                f = range.Font;
                pf = range.ParagraphFormat;

                // 同步所有字体槽（中文/东亚/西文/其他/双字库）
                try { f.Name = fontName; } catch { }
                try { f.NameFarEast = fontName; } catch { }
                try { f.NameAscii = fontName; } catch { }
                try { f.NameOther = fontName; } catch { }
                try { f.NameBi = fontName; } catch { }

                try { f.Size = fontSize; } catch { }

                pf.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                pf.LineSpacing = lineSpacing;
                pf.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                pf.SpaceBefore = 0;
                pf.SpaceAfter = 0;

                // 确保无左/右缩进，特殊缩进（首行/悬挂）为“无”
                try { pf.CharacterUnitFirstLineIndent = 0f; } catch { }
                try { pf.FirstLineIndent = 0f; } catch { }
                try { pf.LeftIndent = 0f; } catch { }
                try { pf.RightIndent = 0f; } catch { }
                try { pf.CharacterUnitLeftIndent = 0f; } catch { }
                try { pf.CharacterUnitRightIndent = 0f; } catch { }
            }
            finally
            {
                if (pf != null) Marshal.ReleaseComObject(pf);
                if (f != null) Marshal.ReleaseComObject(f);
            }

            // 使用更可靠的查找替换方式修复引号/ASCII 标点字体（替代逐字符遍历）
            try
            {
                ForceQuotesFontViaFindReplace(range, fontName, fontSize);
            }
            catch { }
        }

        // 新增：使用 Find.Replacement 在范围内批量设置引号等 ASCII 标点的字体与字号（性能优，兼容性好）
        private void ForceQuotesFontViaFindReplace(Range range, string fontName, float fontSize)
        {
            if (range == null) return;

            Find find = range.Find;

            // 保存原有设置（以 object 形式保存，恢复时尽量容错）
            object origMatchWildcards = null;
            object origWrap = null;
            object origText = null;
            try
            {
                try { origMatchWildcards = find.MatchWildcards; } catch { }
                try { origWrap = find.Wrap; } catch { }
                try { origText = find.Text; } catch { }

                find.ClearFormatting();
                find.Replacement.ClearFormatting();

                // 匹配直引号/弯引号等
                find.Text = "[\"\"''“”‘’]";
                find.Replacement.Text = "^&"; // 保留原字符，仅替换格式

                // 设置替换格式
                Font replFont = find.Replacement.Font;
                try { replFont.Name = fontName; } catch { }
                try { replFont.NameFarEast = fontName; } catch { }
                try { replFont.NameAscii = fontName; } catch { }
                try { replFont.NameOther = fontName; } catch { }
                try { replFont.NameBi = fontName; } catch { }
                try { replFont.Size = fontSize; } catch { }

                find.MatchWildcards = true;
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;

                object replaceAll = WdReplace.wdReplaceAll;
                // 执行替换（在指定范围内）
                find.Execute(Replace: ref replaceAll);
            }
            finally
            {
                // 尝试恢复原有设置，容错处理
                try { if (origMatchWildcards != null) find.MatchWildcards = (bool)origMatchWildcards; } catch { }
                try { if (origWrap != null) find.Wrap = (WdFindWrap)origWrap; } catch { }
                try { if (origText != null) find.Text = (string)origText; } catch { }
            }
        }

        // 兼容旧调用：Selection 版本仅委托给 Range 版本，并保证在光标处（折叠选区）后续输入继承样式
        private void ApplyBasicParagraphStyle(Selection sel, string fontName, float fontSize, float lineSpacing)
        {
            if (sel == null) return;

            // 先对 Range 做统一处理（覆盖选区或当前段落）
            ApplyBasicParagraphStyle(sel.Range, fontName, fontSize, lineSpacing);

            // 额外设置 Selection（确保光标处后后续输入继承样式）
            Font sf = null;
            ParagraphFormat spf = null;
            try
            {
                sf = sel.Font;
                spf = sel.ParagraphFormat;

                // 设置 Selection 的所有字体槽与大小
                try { sf.Name = fontName; } catch { }
                try { sf.NameFarEast = fontName; } catch { }
                try { sf.NameAscii = fontName; } catch { }
                try { sf.NameOther = fontName; } catch { }
                try { sf.NameBi = fontName; } catch { }
                try { sf.Size = fontSize; } catch { }

                spf.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                spf.LineSpacing = lineSpacing;
                spf.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                spf.SpaceBefore = 0;
                spf.SpaceAfter = 0;

                // 确保无左/右缩进，特殊缩进（首行/悬挂）为“无”
                try { spf.CharacterUnitFirstLineIndent = 0f; } catch { }
                try { spf.FirstLineIndent = 0f; } catch { }
                try { spf.LeftIndent = 0f; } catch { }
                try { spf.RightIndent = 0f; } catch { }
                try { spf.CharacterUnitLeftIndent = 0f; } catch { }
                try { spf.CharacterUnitRightIndent = 0f; } catch { }
            }
            finally
            {
                if (spf != null) Marshal.ReleaseComObject(spf);
                if (sf != null) Marshal.ReleaseComObject(sf);
            }

            // 同样对 Selection 所在范围逐字符修复 ASCII 标点/引号（若需要）
            try
            {
                ForceAsciiPunctuationFont(sel.Range, fontName, fontSize);
            }
            catch { }
        }

        // ==================== 编号辅助方法 ====================
        // Range 重载，避免依赖 Selection
        private void ApplyNumbering(Range range, string numberFormat, int startAt)
        {
            ApplyNumbering(range, numberFormat, startAt, (WdListNumberStyle)39);
        }

        private void ApplyNumbering(Range range, string numberFormat, int startAt, WdListNumberStyle numberStyle)
        {
            if (range == null) return;

            ListTemplate lt = null;
            dynamic level = null;
            try
            {
                lt = Globals.ThisAddIn.Application.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
                // 缓存一级 ListLevel 对象以减少 repeated COM 调用
                level = lt.ListLevels[1];

                level.NumberFormat = numberFormat;
                // 使用有效枚举成员：wdTrailingNone
                level.TrailingCharacter = WdTrailingCharacter.wdTrailingNone;
                level.NumberStyle = numberStyle;
                level.NumberPosition = 0;
                level.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                level.TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
                level.TabPosition = (float)WdConstants.wdUndefined;
                level.ResetOnHigher = 0;
                level.StartAt = startAt;

                object bContinuePrevList = false;
                object applyTo = WdListApplyTo.wdListApplyToWholeList;
                object defBehavior = WdDefaultListBehavior.wdWord9ListBehavior;
                range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
            }
            finally
            {
                if (level != null) Marshal.ReleaseComObject(level);
                if (lt != null) Marshal.ReleaseComObject(lt);
            }
        }

        // ==================== 按钮与其它逻辑（确保光标处继承样式） ====================

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 使用 Selection 版本，使光标处也能继承样式
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正小标宋_GBK", 22f, 29f);

            // 特殊对齐
            Range r = sel.Range;
            ParagraphFormat pf = null;
            try
            {
                pf = r.ParagraphFormat;
                pf.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }
            finally
            {
                if (pf != null) Marshal.ReleaseComObject(pf);
            }

            SyncSelectionToRange(sel);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            SyncSelectionToRange(sel);
        }

        // 黑体一级编号（起始1~10）
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 1);
            SyncSelectionToRange(sel);
        }
        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 2);
            SyncSelectionToRange(sel);
        }
        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 3);
            SyncSelectionToRange(sel);
        }
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 4);
            SyncSelectionToRange(sel);
        }
        private void button14_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 5);
            SyncSelectionToRange(sel);
        }
        private void button15_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 6);
            SyncSelectionToRange(sel);
        }
        private void button16_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 7);
            SyncSelectionToRange(sel);
        }
        private void button17_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 8);
            SyncSelectionToRange(sel);
        }
        private void button18_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 9);
            SyncSelectionToRange(sel);
        }
        private void button19_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 10);
            SyncSelectionToRange(sel);
        }

        // 楷体二级编号（起始1~10）
        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 1);
            SyncSelectionToRange(sel);
        }
        private void button20_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 2);
            SyncSelectionToRange(sel);
        }
        private void button21_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 3);
            SyncSelectionToRange(sel);
        }
        private void button22_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 4);
            SyncSelectionToRange(sel);
        }
        private void button23_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 5);
            SyncSelectionToRange(sel);
        }
        private void button24_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 6);
            SyncSelectionToRange(sel);
        }
        private void button25_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 7);
            SyncSelectionToRange(sel);
        }
        private void button26_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正楱体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 8);
            SyncSelectionToRange(sel);
        }
        private void button27_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 9);
            SyncSelectionToRange(sel);
        }
        private void button28_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 10);
            SyncSelectionToRange(sel);
        }

        // 仿宋三级编号（阿拉伯数字，起始1~10）
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 1, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button29_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 2, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button30_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 3, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button31_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 4, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button32_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 5, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button33_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 6, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button34_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 7, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button35_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 8, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button36_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 9, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button37_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 10, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }

        // 仿宋四级编号（阿拉伯数字，起始1~10）
        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 1, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button38_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 2, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button39_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 3, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button40_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 4, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button41_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 5, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button42_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 6, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button43_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 7, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button44_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 8, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button45_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 9, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button46_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 10, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }

        // ==================== 表格处理 ====================
        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            WithScreenUpdatingDisabled(() =>
            {
                Application app = Globals.ThisAddIn.Application;
                Document doc = app.ActiveDocument;
                Selection sel = app.Selection;
                const float FallbackMinLineSpacing = 0.7f;
                // 先清除所有格式（按新约定）
                ClearFormatting(sel);
                bool selHasTables = sel != null && sel.Tables != null && sel.Tables.Count >= 1;
                bool selInTable = false;
                try
                {
                    if (sel != null)
                        selInTable = sel.get_Information(Microsoft.Office.Interop.Word.WdInformation.wdWithInTable);
                }
                catch { selInTable = false; }

                // 优先：已选中一个或多个表格；其次：光标位于表格内（定位该表）；否则回退到文档第一个表格
                if (selHasTables || selInTable)
                {
                    // 如果光标在表格内但没有选中完整表格，处理包含光标的表
                    if (!selHasTables && selInTable)
                    {
                        Table tbl = null;
                        Range tr = null;
                        try
                        {
                            // 获取包含光标的表（若光标在单元格内）
                            tbl = sel.Range.Tables[1];
                            tr = tbl.Range;

                            // 先设为“网格型”样式，再做具体格式设置
                            tr.set_Style("网格型");

                            tr.Font.Reset();
                            tr.ParagraphFormat.Reset();

                            tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                            // 字体：方正仿宋_GBK（覆盖各字体槽）
                            tr.Font.Name = "方正仿宋_GBK";
                            try { tr.Font.NameFarEast = "方正仿宋_GBK"; } catch { }
                            try { tr.Font.NameAscii = "方正仿宋_GBK"; } catch { }
                            try { tr.Font.NameOther = "方正仿宋_GBK"; } catch { }

                            tr.Font.Size = 10.5f;
                            // 水平居中（表格内文字）
                            tr.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            tbl.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                            tr.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast;
                            try
                            {
                                tr.ParagraphFormat.LineSpacing = 0f; // 尝试“最小值”
                            }
                            catch (System.Runtime.InteropServices.COMException)
                            {
                                try
                                {
                                    tr.ParagraphFormat.LineSpacing = FallbackMinLineSpacing; // 合法回落
                                }
                                catch (System.Runtime.InteropServices.COMException)
                                {
                                    tr.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                                }
                            }

                            tr.ParagraphFormat.SpaceBefore = 0;
                            tr.ParagraphFormat.SpaceAfter = 0;

                            tr.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
                            tr.ParagraphFormat.FirstLineIndent = 0f;
                            tr.ParagraphFormat.LeftIndent = 0f;
                            tr.ParagraphFormat.CharacterUnitLeftIndent = 0f;

                            tbl.Borders.Enable = 1;
                            tbl.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                            tbl.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;

                            // 逐单元格强制设置段落对齐，兼容合并/样式问题
                            Cells cells = null;
                            try
                            {
                                cells = tbl.Range.Cells;
                                for (int ci = 1; ci <= cells.Count; ++ci)
                                {
                                    Cell c = null;
                                    Range cr = null;
                                    try
                                    {
                                        c = cells[ci];
                                        cr = c.Range;
                                        cr.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                    }
                                    finally
                                    {
                                        if (cr != null) Marshal.ReleaseComObject(cr);
                                        if (c != null) Marshal.ReleaseComObject(c);
                                    }
                                }
                            }
                            finally
                            {
                                if (cells != null) Marshal.ReleaseComObject(cells);
                            }

                            if (tbl.Rows.Count >= 1)
                            {
                                SafeApplyFirstLogicalRowHeader(tbl, "方正仿宋_GBK", 10.5f);
                            }
                        }
                        finally
                        {
                            if (tr != null) Marshal.ReleaseComObject(tr);
                            if (tbl != null) Marshal.ReleaseComObject(tbl);
                        }

                        return;
                    }

                    // 已选中一个或多个表格（处理每个选中表）
                    for (int ti = 1; ti <= sel.Tables.Count; ++ti)
                    {
                        Table tbl = null;
                        Range tr = null;
                        try
                        {
                            tbl = sel.Tables[ti];
                            tr = tbl.Range;

                            tr.Font.Reset();
                            tr.ParagraphFormat.Reset();

                            tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                            // 字体：方正仿宋_GBK（覆盖各字体槽）
                            tr.Font.Name = "方正仿宋_GBK";
                            try { tr.Font.NameFarEast = "方正仿宋_GBK"; } catch { }
                            try { tr.Font.NameAscii = "方正仿宋_GBK"; } catch { }
                            try { tr.Font.NameOther = "方正仿宋_GBK"; } catch { }

                            tr.Font.Size = 10.5f;
                            tr.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            tbl.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                            tr.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast;
                            try
                            {
                                tr.ParagraphFormat.LineSpacing = 0f; // 尝试“最小值”
                            }
                            catch (System.Runtime.InteropServices.COMException)
                            {
                                try
                                {
                                    tr.ParagraphFormat.LineSpacing = FallbackMinLineSpacing; // 合法回落
                                }
                                catch (System.Runtime.InteropServices.COMException)
                                {
                                    tr.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                                }
                            }

                            tr.ParagraphFormat.SpaceBefore = 0;
                            tr.ParagraphFormat.SpaceAfter = 0;

                            tr.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
                            tr.ParagraphFormat.FirstLineIndent = 0f;
                            tr.ParagraphFormat.LeftIndent = 0f;
                            tr.ParagraphFormat.CharacterUnitLeftIndent = 0f;

                            tbl.Borders.Enable = 1;
                            tbl.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                            tbl.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;

                            // 逐单元格强制设置段落对齐
                            Cells cells = null;
                            try
                            {
                                cells = tbl.Range.Cells;
                                for (int ci = 1; ci <= cells.Count; ++ci)
                                {
                                    Cell c = null;
                                    Range cr = null;
                                    try
                                    {
                                        c = cells[ci];
                                        cr = c.Range;
                                        cr.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                    }
                                    finally
                                    {
                                        if (cr != null) Marshal.ReleaseComObject(cr);
                                        if (c != null) Marshal.ReleaseComObject(c);
                                    }
                                }
                            }
                            finally
                            {
                                if (cells != null) Marshal.ReleaseComObject(cells);
                            }

                            if (tbl.Rows.Count >= 1)
                            {
                                SafeApplyFirstLogicalRowHeader(tbl, "方正仿宋_GBK", 10.5f);
                            }
                        }
                        finally
                        {
                            if (tr != null) Marshal.ReleaseComObject(tr);
                            if (tbl != null) Marshal.ReleaseComObject(tbl);
                        }
                    }

                    return;
                }

                // 回退：文档第一个表格（原样保留）
                if (doc == null || doc.Tables.Count < 1) return;

                Table firstTbl = null;
                Range firstTr = null;
                try
                {
                    firstTbl = doc.Tables[1];
                    firstTr = firstTbl.Range;

                    firstTr.Font.Reset();
                    firstTr.ParagraphFormat.Reset();

                    firstTbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                    firstTr.Font.Name = "方正仿宋_GBK";
                    try { firstTr.Font.NameFarEast = "方正仿宋_GBK"; } catch { }
                    try { firstTr.Font.NameAscii = "方正仿宋_GBK"; } catch { }
                    try { firstTr.Font.NameOther = "方正仿宋_GBK"; } catch { }

                    firstTr.Font.Size = 10.5f;
                    firstTr.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    firstTbl.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    firstTr.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast;
                    try
                    {
                        firstTr.ParagraphFormat.LineSpacing = 0f;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        firstTr.ParagraphFormat.LineSpacing = FallbackMinLineSpacing;
                    }

                    firstTr.ParagraphFormat.SpaceBefore = 0;
                    firstTr.ParagraphFormat.SpaceAfter = 0;
                    firstTr.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
                    firstTr.ParagraphFormat.FirstLineIndent = 0f;
                    firstTr.ParagraphFormat.LeftIndent = 0f;
                    firstTr.ParagraphFormat.CharacterUnitLeftIndent = 0f;

                    firstTbl.Borders.Enable = 1;
                    firstTbl.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                    firstTbl.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;

                    if (firstTbl.Rows.Count >= 1)
                    {
                        SafeApplyFirstLogicalRowHeader(firstTbl, "方正仿宋_GBK", 10.5f);
                    }
                }
                finally
                {
                    if (firstTr != null) Marshal.ReleaseComObject(firstTr);
                    if (firstTbl != null) Marshal.ReleaseComObject(firstTbl);
                }
            });
        }

        private void button58_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            
            WithScreenUpdatingDisabled(() =>
            {
                Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                const float FallbackMinLineSpacing = 0.7f;

                for (int i = 1; i <= doc.Tables.Count; ++i)
                {
                    Table tbl = null;
                    Range tr = null;
                    try
                    {
                        tbl = doc.Tables[i];
                        tr = tbl.Range;

                        // 先设为“网格型”样式，再做具体格式设置
                        tr.set_Style("网格型");

                        tr.Font.Reset();
                        tr.ParagraphFormat.Reset();

                        tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                        // 字体：方正仿宋_GBK
                        tr.Font.Name = "方正仿宋_GBK";
                        try { tr.Font.NameFarEast = "方正仿宋_GBK"; } catch { }
                        try { tr.Font.NameAscii = "方正仿宋_GBK"; } catch { }
                        try { tr.Font.NameOther = "方正仿宋_GBK"; } catch { }

                        tr.Font.Size = 10.5f;
                        // 改为两端对齐（仅此行被修改）
                        tr.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        tbl.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                        tr.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast;
                        try
                        {
                            tr.ParagraphFormat.LineSpacing = 0f; // 尝试“最小值”
                            tr.ParagraphFormat.LineSpacing = FallbackMinLineSpacing; // 合法回落
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            tr.ParagraphFormat.LineSpacing = FallbackMinLineSpacing;
                        }

                        tr.ParagraphFormat.SpaceBefore = 0;
                        tr.ParagraphFormat.SpaceAfter = 0;
                        tr.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
                        tr.ParagraphFormat.FirstLineIndent = 0f;
                        tr.ParagraphFormat.LeftIndent = 0f;
                        tr.ParagraphFormat.CharacterUnitLeftIndent = 0f;

                        tbl.Borders.Enable = 1;
                        tbl.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                        tbl.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
                        tbl.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                        tbl.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                        tbl.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                        tbl.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                        tbl.Borders[WdBorderType.wdBorderHorizontal].LineStyle = WdLineStyle.wdLineStyleSingle;
                        tbl.Borders[WdBorderType.wdBorderVertical].LineStyle = WdLineStyle.wdLineStyleSingle;

                        if (tbl.Rows.Count >= 1)
                        {
                            SafeApplyFirstLogicalRowHeader(tbl, "方正仿宋_GBK", 10.5f);
                        }
                    }
                    finally
                    {
                        if (tr != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(tr);
                        if (tbl != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(tbl);
                    }
                }
            });
        }

        // ==================== 标记和颜色 ====================
        private void button11_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.HighlightColorIndex = WdColorIndex.wdYellow;
            sel.Range.Font.Color = WdColor.wdColorRed;

            // 额外设置 Selection 本身（使光标后续输入为红色、带高亮）
            Font sf = null;
            ParagraphFormat spf = null;
            try
            {
                sf = sel.Font;
                spf = sel.ParagraphFormat;

                sf.Color = WdColor.wdColorRed;
                spf.LineSpacingRule = spf.LineSpacingRule; // 保持原有行距规则
            }
            finally
            {
                if (spf != null) Marshal.ReleaseComObject(spf);
                if (sf != null) Marshal.ReleaseComObject(sf);
            }

            SyncSelectionToRange(sel);
        }

        private void button13_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
            sel.Range.Font.Color = WdColor.wdColorAutomatic;

            // 额外设置 Selection 本身（清除光标处的颜色设置）
            Font sf = null;
            try
            {
                sf = sel.Font;
                sf.Color = WdColor.wdColorAutomatic;
            }
            finally
            {
                if (sf != null) Marshal.ReleaseComObject(sf);
            }

            SyncSelectionToRange(sel);
        }

        // ==================== 页面设置（仅页面项 + 仅修改“正文”样式） ====================
        private void button12_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;
            Document doc = app.ActiveDocument;
            bool originalScreenUpdating = app.ScreenUpdating;
            try
            {
                app.ScreenUpdating = false;

                doc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
                doc.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
                doc.PageSetup.TopMargin = ConvertMillimetersToPoints(37);
                doc.PageSetup.BottomMargin = ConvertMillimetersToPoints(35);
                doc.PageSetup.LeftMargin = ConvertMillimetersToPoints(28);
                doc.PageSetup.RightMargin = ConvertMillimetersToPoints(26);
                doc.PageSetup.FooterDistance = ConvertMillimetersToPoints(24.7);
                doc.PageSetup.LayoutMode = (WdLayoutMode)1; // wdLayoutModeGrid
                doc.PageSetup.LinesPage = 22;

                Style normalStyle = null;
                try
                {
                    normalStyle = doc.Styles[WdBuiltinStyle.wdStyleNormal];
                    if (normalStyle != null && normalStyle.ParagraphFormat != null)
                    {
                        normalStyle.ParagraphFormat.FarEastLineBreakControl = 0;
                    }
                }
                catch
                {
                }
                finally
                {
                    if (normalStyle != null) Marshal.ReleaseComObject(normalStyle);
                }
            }
            finally
            {
                app.ScreenUpdating = originalScreenUpdating;
            }
        }

        // ==================== 查找替换 ====================
        /// <summary>
        /// 按钮：使用正则通配符查找多余的段落标记并替换为单个段落（针对中文标点后的多余换行）。
        /// </summary>
        private void button47_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 查找模式：捕获句末标点后跟随的多个回车，替换为单个回车
            sel.Find.Text = "([!。！？……])^13{1,}";
            sel.Find.Replacement.Text = @"\1";
            sel.Find.Forward = true;
            sel.Find.Wrap = WdFindWrap.wdFindContinue;
            sel.Find.Format = false;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = true;
            sel.Find.MatchAllWordForms = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchWildcards = true;
            object replaceAll = WdReplace.wdReplaceAll;
            object oMissing = Type.Missing;
            // 执行替换（使用 ref 参数签名）
            sel.Find.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                             ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                             ref oMissing, ref oMissing, ref replaceAll, ref oMissing,
                             ref oMissing, ref oMissing, ref oMissing);

            SyncSelectionToRange(sel);
        }

        // ==================== 大纲级别单独设置 ====================
        private void button48_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            SyncSelectionToRange(sel);
        }
        private void button49_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            SyncSelectionToRange(sel);
        }
        private void button50_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            SyncSelectionToRange(sel);
        }
        private void button51_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            SyncSelectionToRange(sel);
        }
        private void button52_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel5;
            SyncSelectionToRange(sel);
        }
        private void button53_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel6;
            SyncSelectionToRange(sel);
        }
        private void button54_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Paragraphs.OutlinePromote();
            SyncSelectionToRange(sel);
        }
        private void button55_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Paragraphs.OutlineDemote();
            SyncSelectionToRange(sel);
        }

        // ==================== 缩进控制 ====================
        /// <summary>设置首行缩进为 2 个字符单位（用于正文）。</summary>
        private void button56_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 2f;

            // 同步 Selection（确保光标处的后续输入继承）
            ParagraphFormat spf = null;
            try
            {
                spf = sel.ParagraphFormat;
                spf.CharacterUnitFirstLineIndent = 2f;
            }
            finally
            {
                if (spf != null) Marshal.ReleaseComObject(spf);
            }

            SyncSelectionToRange(sel);
        }

        /// <summary>取消首行与左缩进，恢复为无缩进状态。</summary>
        private void button57_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
            sel.Range.ParagraphFormat.FirstLineIndent = 0f;
            sel.Range.ParagraphFormat.LeftIndent = 0f;
            sel.Range.ParagraphFormat.CharacterUnitLeftIndent = 0f;

            // 同步 Selection（确保光标处后续输入继承）
            ParagraphFormat spf = null;
            try
            {
                spf = sel.ParagraphFormat;
                spf.CharacterUnitFirstLineIndent = 0f;
                spf.FirstLineIndent = 0f;
                spf.LeftIndent = 0f;
                spf.CharacterUnitLeftIndent = 0f;
            }
            finally
            {
                if (spf != null) Marshal.ReleaseComObject(spf);
            }

            SyncSelectionToRange(sel);
        }

        private void button59_Click(object sender, RibbonControlEventArgs e)
        {
            // 小标题：方正楷体 三号 居中 正文文本 缩进左右均为0 缩进特殊无 间距段前后均为0 行距29磅
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 先清除所有格式（按新约定）
            ClearFormatting(sel);
            const string fontName = "方正楷体_GBK";
            const float fontSize = 16f; // 三号
            const float lineSpacing = 29f;

            // 使用 Selection 版本，保证光标处后续输入继承样式
            ApplyBasicParagraphStyle(sel, fontName, fontSize, lineSpacing);

            // 覆盖对齐与缩进为题述要求
            ParagraphFormat pf = null;
            try
            {
                pf = sel.ParagraphFormat;
                pf.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                pf.CharacterUnitFirstLineIndent = 0f;
                pf.FirstLineIndent = 0f;
                pf.LeftIndent = 0f;
                pf.CharacterUnitLeftIndent = 0f;
                pf.SpaceBefore = 0;
                pf.SpaceAfter = 0;
            }
            finally
            {
                if (pf != null) Marshal.ReleaseComObject(pf);
            }

            // 确保 Selection.Font 也一致（光标处后续输入）
            Font sf = null;
            try
            {
                sf = sel.Font;
                sf.Name = fontName;
                sf.Size = fontSize;
            }
            finally
            {
                if (sf != null) Marshal.ReleaseComObject(sf);
            }

            SyncSelectionToRange(sel);
        }

        // C#
        private void SafeApplyFirstLogicalRowHeader(Table tbl, string fontName, float fontSize)
        {
            if (tbl == null) return;

            Row hdrRow = null;
            Range hdrRange = null;
            try
            {
                try
                {
                    hdrRow = tbl.Rows[1];
                    hdrRange = hdrRow.Range;
                    hdrRange.Font.Name = fontName;
                    hdrRange.Font.Size = fontSize;
                    hdrRange.Font.Bold = 1;

                    // 首行底纹：前景白色，背景1 深色15%（近似为 RGB(217,217,217)），无纹理
                    try
                    {
                        hdrRange.Shading.ForegroundPatternColor = WdColor.wdColorWhite;
                        int bgOle = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(217, 217, 217));
                        hdrRange.Shading.BackgroundPatternColor = (WdColor)bgOle;
                        hdrRange.Shading.Texture = Microsoft.Office.Interop.Word.WdTextureIndex.wdTextureNone;
                    }
                    catch { /* 忽略无法设置底纹的环境差异 */ }

                    try { hdrRow.HeadingFormat = -1; }
                    catch (ArgumentException) { ApplyHeaderByCells(tbl, fontName, fontSize); }
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // 无法按 Rows 访问 -> 回退到按单元格处理
                    ApplyHeaderByCells(tbl, fontName, fontSize);
                }
            }
            finally
            {
                if (hdrRange != null) Marshal.ReleaseComObject(hdrRange);
                if (hdrRow != null) Marshal.ReleaseComObject(hdrRow);
            }
        }

        private void ApplyHeaderByCells(Table tbl, string fontName, float fontSize)
        {
            if (tbl == null) return;

            Cells cells = null;
            try
            {
                cells = tbl.Range.Cells;
                int n = cells.Count;
                for (int i = 1; i <= n; ++i)
                {
                    Cell cell = null;
                    Range cellRange = null;
                    Font cellFont = null;
                    try
                    {
                        cell = cells[i];
                        // 只处理逻辑上位于第一行的单元格（适用于有纵向合并的表）
                        if (cell.RowIndex == 1)
                        {
                            cellRange = cell.Range;
                            cellFont = cellRange.Font;
                            cellFont.Name = fontName;
                            cellFont.Size = fontSize;

                            // 单元格底纹设置（与 SafeApplyFirstLogicalRowHeader 保持一致）
                            try
                            {
                                cellRange.Shading.ForegroundPatternColor = WdColor.wdColorWhite;
                                int bgOle = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(217, 217, 217));
                                cellRange.Shading.BackgroundPatternColor = (WdColor)bgOle;
                                cellRange.Shading.Texture = Microsoft.Office.Interop.Word.WdTextureIndex.wdTextureNone;
                            }
                            catch { }
                        }
                    }
                    finally
                    {
                        if (cellFont != null) Marshal.ReleaseComObject(cellFont);
                        if (cellRange != null) Marshal.ReleaseComObject(cellRange);
                        if (cell != null) Marshal.ReleaseComObject(cell);
                    }
                }
            }
            finally
            {
                if (cells != null) Marshal.ReleaseComObject(cells);
            }
        }

        private void ApplyHeaderByCellsSafe(Table tbl, string fontName, float fontSize)
        {
            if (tbl == null) return;

            Cells allCells = null;
            try
            {
                allCells = tbl.Range.Cells;
                for (int i = 1; i <= allCells.Count; ++i)
                {
                    Cell c = null;
                    Range r = null;
                    Font f = null;
                    try
                    {
                        c = allCells[i];
                        // RowIndex 是单元格在表格中的起始行（合并单元格的顶端也返回顶行索引）
                        if (c.RowIndex != 1) continue;

                        r = c.Range;
                        f = r.Font;
                        f.Name = fontName;
                        f.Size = fontSize;
                        f.Bold = 1;

                        // 设置首行单元格底纹
                        try
                        {
                            r.Shading.ForegroundPatternColor = WdColor.wdColorWhite;
                            int bgOle = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(217, 217, 217));
                            r.Shading.BackgroundPatternColor = (WdColor)bgOle;
                            r.Shading.Texture = Microsoft.Office.Interop.Word.WdTextureIndex.wdTextureNone;
                        }
                        catch { }
                    }
                    finally
                    {
                        if (f != null) Marshal.ReleaseComObject(f);
                        if (r != null) Marshal.ReleaseComObject(r);
                        if (c != null) Marshal.ReleaseComObject(c);
                    }
                }
            }
            finally
            {
                if (allCells != null) Marshal.ReleaseComObject(allCells);
            }
        }

        private void button60_Click(object sender, RibbonControlEventArgs e)
        {
            WithScreenUpdatingDisabled(() =>
            {
                var app = Globals.ThisAddIn.Application;
                Microsoft.Office.Interop.Word.Selection sel = null;
                Microsoft.Office.Interop.Word.Range rng = null;
                try
                {
                    sel = app.Selection;
                    if (sel == null) return;

                    rng = sel.Range;
                    if (rng == null) return;

                    // 使用高亮颜色：鲜绿
                    rng.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBrightGreen;
                }
                finally
                {
                    if (rng != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rng);
                    if (sel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(sel);
                }
            });
        }

        // 工具：清除 Range/Selection/Table 的常见格式（统一入口）
        private void ClearFormatting(Range range)
        {
            if (range == null) return;
            try { range.Font.Reset(); } catch { }
            try { range.ParagraphFormat.Reset(); } catch { }
            try { range.HighlightColorIndex = WdColorIndex.wdNoHighlight; } catch { }
        }

        private void ClearFormatting(Selection sel)
        {
            if (sel == null) return;
            ClearFormatting(sel.Range);
        }

        private void ClearFormatting(Table tbl)
        {
            if (tbl == null) return;
            ClearFormatting(tbl.Range);
        }

        /// <summary>
        /// 强制将 Range 中的 ASCII 标点和引号字符的字体和字号设置为指定值
        /// </summary>
        private void ForceAsciiPunctuationFont(Range range, string fontName, float fontSize)
        {
            if (range == null) return;

            // 只处理文本内容
            string text = range.Text;
            if (string.IsNullOrEmpty(text)) return;

            // ASCII 标点和引号的 Unicode 范围
            char[] asciiPunctuations = { '!', '"', '#', '$', '%', '&', '\'', '(', ')', '*', '+', ',', '-', '.', '/', ':', ';', '<', '=', '>', '?', '@', '[', '\\', ']', '^', '_', '`', '{', '|', '}', '~' };

            for (int i = 1; i <= text.Length; i++)
            {
                char c = text[i - 1];
                if (c <= 0x7F && (char.IsPunctuation(c) || c == '"' || c == '\'' || Array.IndexOf(asciiPunctuations, c) >= 0))
                {
                    Range charRange = range.Duplicate;
                    charRange.SetRange(range.Start + i - 1, range.Start + i);
                    try
                    {
                        charRange.Font.Name = fontName;
                        charRange.Font.Size = fontSize;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(charRange);
                    }
                }
            }
        }
    }
}