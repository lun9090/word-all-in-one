using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Runtime.InteropServices;
using System.Threading;
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

                // 字体
                if (rf != null && sf != null)
                {
                    try { sf.Name = rf.Name; } catch { }
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

                f.Name = fontName;
                f.Size = fontSize;

                pf.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                pf.LineSpacing = lineSpacing;
                pf.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                pf.SpaceBefore = 0;
                pf.SpaceAfter = 0;
            }
            finally
            {
                if (pf != null) Marshal.ReleaseComObject(pf);
                if (f != null) Marshal.ReleaseComObject(f);
            }
        }

        // 兼容旧调用：Selection 版本仅委托给 Range 版本，并保证在光标处（折叠选区）后续输入继承样式
        private void ApplyBasicParagraphStyle(Selection sel, string fontName, float fontSize, float lineSpacing)
        {
            if (sel == null) return;

            // 先对 Range 做统一处理（覆盖选区或当前段落）
            ApplyBasicParagraphStyle(sel.Range, fontName, fontSize, lineSpacing);

            // 额外设置 Selection（确保光标处后续输入继承样式）
            Font sf = null;
            ParagraphFormat spf = null;
            try
            {
                sf = sel.Font;
                spf = sel.ParagraphFormat;

                // 设置 Selection 的字体与段落属性，使插入点处的后续输入采用相同格式
                sf.Name = fontName;
                sf.Size = fontSize;

                spf.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                spf.LineSpacing = lineSpacing;
                // 保持 Range 版本的一致默认（两端对齐）；调用方可再覆盖 Alignment
                spf.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                spf.SpaceBefore = 0;
                spf.SpaceAfter = 0;
            }
            finally
            {
                if (spf != null) Marshal.ReleaseComObject(spf);
                if (sf != null) Marshal.ReleaseComObject(sf);
            }
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
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            SyncSelectionToRange(sel);
        }

        // 黑体一级编号（起始1~10）
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 1);
            SyncSelectionToRange(sel);
        }
        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 2);
            SyncSelectionToRange(sel);
        }
        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 3);
            SyncSelectionToRange(sel);
        }
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 4);
            SyncSelectionToRange(sel);
        }
        private void button14_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 5);
            SyncSelectionToRange(sel);
        }
        private void button15_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 6);
            SyncSelectionToRange(sel);
        }
        private void button16_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 7);
            SyncSelectionToRange(sel);
        }
        private void button17_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 8);
            SyncSelectionToRange(sel);
        }
        private void button18_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 9);
            SyncSelectionToRange(sel);
        }
        private void button19_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正黑体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 10);
            SyncSelectionToRange(sel);
        }

        // 楷体二级编号（起始1~10）
        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 1);
            SyncSelectionToRange(sel);
        }
        private void button20_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 2);
            SyncSelectionToRange(sel);
        }
        private void button21_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 3);
            SyncSelectionToRange(sel);
        }
        private void button22_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 4);
            SyncSelectionToRange(sel);
        }
        private void button23_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 5);
            SyncSelectionToRange(sel);
        }
        private void button24_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 6);
            SyncSelectionToRange(sel);
        }
        private void button25_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 7);
            SyncSelectionToRange(sel);
        }
        private void button26_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正楱体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 8);
            SyncSelectionToRange(sel);
        }
        private void button27_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 9);
            SyncSelectionToRange(sel);
        }
        private void button28_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正楷体_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            ApplyNumbering(sel.Range, "（%1）", 10);
            SyncSelectionToRange(sel);
        }

        // 仿宋三级编号（阿拉伯数字，起始1~10）
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 1, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button29_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 2, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button30_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 3, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button31_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 4, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button32_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 5, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button33_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 6, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button34_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 7, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button35_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 8, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button36_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 9, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button37_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            ApplyNumbering(sel.Range, "%1．", 10, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }

        // 仿宋四级编号（阿拉伯数字，起始1~10）
        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 1, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button38_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 2, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button39_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 3, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button40_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 4, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button41_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 5, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button42_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 6, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button43_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 7, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button44_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 8, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button45_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel, "方正仿宋_GBK", 16f, 29f);
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
            ApplyNumbering(sel.Range, "（%1）", 9, WdListNumberStyle.wdListNumberStyleArabic);
            SyncSelectionToRange(sel);
        }
        private void button46_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
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

                // 优先处理选中的表格；无则回退到文档第一个表格
                if (sel != null && sel.Tables != null && sel.Tables.Count >= 1)
                {
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

                            // 行距：选择“最小值”对应用户期望的 0 磅，若被拒绝则回退到 0.7 磅
                            tr.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast;
                            try
                            {
                                tr.ParagraphFormat.LineSpacing = 0f; // 尝试设置为 0（表示“最小值”）
                            }
                            catch (System.Runtime.InteropServices.COMException)
                            {
                                tr.ParagraphFormat.LineSpacing = FallbackMinLineSpacing;
                            }

                            // 段前/段后为 0
                            tr.ParagraphFormat.SpaceBefore = 0;
                            tr.ParagraphFormat.SpaceAfter = 0;

                            tr.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
                            tr.ParagraphFormat.FirstLineIndent = 0f;
                            tr.ParagraphFormat.LeftIndent = 0f;
                            tr.ParagraphFormat.CharacterUnitLeftIndent = 0f;

                            tbl.Borders.Enable = 1;
                            tbl.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                            tbl.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;

                            if (tbl.Rows.Count >= 1)
                            {
                                SafeApplyFirstLogicalRowHeader(tbl, "方正仿宋_GBK", 10.5f);
                                // 不要直接: Range hdrRange = tbl.Rows[1].Range;
                                // SafeApplyFirstLogicalRowHeader 内会在可行时设置 HeadingFormat；在回退情况下按单元格处理样式
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

                // 回退：文档第一个表格
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
                    firstTr.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
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
                        // 不要直接: Range hdrRange = firstTbl.Rows[1].Range;
                        // SafeApplyFirstLogicalRowHeader 内会在可行时设置 HeadingFormat；在回退情况下按单元格处理样式
                    }
                }
                finally
                {
                    if (firstTr != null) Marshal.ReleaseComObject(firstTr);
                    if (firstTbl != null) Marshal.ReleaseComObject(firstTbl);
                }
            });
        }

        private void button58_Click(object sender, RibbonControlEventArgs e)
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
                        tr.Font.Reset();
                        tr.ParagraphFormat.Reset();

                        tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                        // 字体：方正仿宋_GBK
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
                            tr.ParagraphFormat.LineSpacing = 0f;
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
                            // 不要直接: Range hdrRange = tbl.Rows[1].Range;
                            // SafeApplyFirstLogicalRowHeader 内会在可行时设置 HeadingFormat；在回退情况下按单元格处理样式
                        }
                    }
                    finally
                    {
                        if (tr != null) Marshal.ReleaseComObject(tr);
                        if (tbl != null) Marshal.ReleaseComObject(tbl);
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

            // 优先尝试直接按行访问（性能最好）
            try
            {
                Range hdrRange = tbl.Rows[1].Range;
                try
                {
                    hdrRange.Font.Name = fontName;
                    hdrRange.Font.Size = fontSize;
                    hdrRange.Font.Bold = 1;
                }
                finally
                {
                    if (hdrRange != null) Marshal.ReleaseComObject(hdrRange);
                }

                // 如果可以访问 Rows[1]，也设置 HeadingFormat
                tbl.Rows[1].HeadingFormat = 1;
                return;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // 回退：按单元格处理属于“逻辑第一行”的单元格
                int cellCount = tbl.Range.Cells.Count;
                for (int ci = 1; ci <= cellCount; ++ci)
                {
                    Cell c = null;
                    try
                    {
                        c = tbl.Range.Cells[ci];
                        if (c.RowIndex == 1)
                        {
                            Range cr = c.Range;
                            try
                            {
                                cr.Font.Name = fontName;
                                cr.Font.Size = fontSize;
                                cr.Font.Bold = 1;
                                cr.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                cr.ParagraphFormat.SpaceBefore = 0;
                                cr.ParagraphFormat.SpaceAfter = 0;
                            }
                            finally
                            {
                                if (cr != null) Marshal.ReleaseComObject(cr);
                            }
                        }
                    }
                    finally
                    {
                        if (c != null) Marshal.ReleaseComObject(c);
                    }
                }

                // 不能安全设置 HeadingFormat —— 可在这里记录或通知用户
            }
        }
    }
}