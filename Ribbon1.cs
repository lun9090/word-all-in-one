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

        // 兼容旧调用：Selection 版本仅委托给 Range 版本
        private void ApplyBasicParagraphStyle(Selection sel, string fontName, float fontSize, float lineSpacing)
        {
            if (sel == null) return;
            ApplyBasicParagraphStyle(sel.Range, fontName, fontSize, lineSpacing);
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

        // ==================== 按钮与其它逻辑（保持原样，但改为 Range 操作） ====================
        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            // 直接使用 Range 实现
            ApplyBasicParagraphStyle(sel.Range, "方正小标宋_GBK", 22f, 29f);
            // 特殊对齐与字号（如果额外设置需要）
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
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f);
        }

        // 黑体一级编号（起始1~10）
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            ApplyBasicParagraphStyle(sel.Range, "方正黑体_GBK", 16f, 29f);
            // 设置大纲级别并应用编号
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
            ApplyNumbering(sel.Range, "%1、", 1);
        }
        private void button5_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正黑体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1; ApplyNumbering(sel.Range, "%1、", 2); }
        private void button6_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正黑体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1; ApplyNumbering(sel.Range, "%1、", 3); }
        private void button7_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正黑体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1; ApplyNumbering(sel.Range, "%1、", 4); }
        private void button14_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正黑体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1; ApplyNumbering(sel.Range, "%1、", 5); }
        private void button15_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正黑体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1; ApplyNumbering(sel.Range, "%1、", 6); }
        private void button16_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正黑体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1; ApplyNumbering(sel.Range, "%1、", 7); }
        private void button17_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正黑体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1; ApplyNumbering(sel.Range, "%1、", 8); }
        private void button18_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正黑体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1; ApplyNumbering(sel.Range, "%1、", 9); }
        private void button19_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正黑体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1; ApplyNumbering(sel.Range, "%1、", 10); }

        // 楷体二级编号（起始1~10）
        private void button8_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正楷体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2; ApplyNumbering(sel.Range, "（%1）", 1); }
        private void button20_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正楷体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2; ApplyNumbering(sel.Range, "（%1）", 2); }
        private void button21_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正楷体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2; ApplyNumbering(sel.Range, "（%1）", 3); }
        private void button22_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正楷体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2; ApplyNumbering(sel.Range, "（%1）", 4); }
        private void button23_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正楷体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2; ApplyNumbering(sel.Range, "（%1）", 5); }
        private void button24_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正楷体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2; ApplyNumbering(sel.Range, "（%1）", 6); }
        private void button25_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正楷体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2; ApplyNumbering(sel.Range, "（%1）", 7); }
        private void button26_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正楱体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2; ApplyNumbering(sel.Range, "（%1）", 8); }
        private void button27_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正楷体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2; ApplyNumbering(sel.Range, "（%1）", 9); }
        private void button28_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正楷体_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2; ApplyNumbering(sel.Range, "（%1）", 10); }

        // 仿宋三级编号（阿拉伯数字，起始1~10）
        private void button4_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3; ApplyNumbering(sel.Range, "%1．", 1, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button29_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3; ApplyNumbering(sel.Range, "%1．", 2, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button30_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3; ApplyNumbering(sel.Range, "%1．", 3, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button31_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3; ApplyNumbering(sel.Range, "%1．", 4, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button32_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3; ApplyNumbering(sel.Range, "%1．", 5, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button33_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3; ApplyNumbering(sel.Range, "%1．", 6, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button34_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3; ApplyNumbering(sel.Range, "%1．", 7, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button35_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3; ApplyNumbering(sel.Range, "%1．", 8, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button36_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3; ApplyNumbering(sel.Range, "%1．", 9, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button37_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3; ApplyNumbering(sel.Range, "%1．", 10, WdListNumberStyle.wdListNumberStyleArabic); }

        // 仿宋四级编号（阿拉伯数字，起始1~10）
        private void button9_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4; ApplyNumbering(sel.Range, "（%1）", 1, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button38_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4; ApplyNumbering(sel.Range, "（%1）", 2, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button39_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4; ApplyNumbering(sel.Range, "（%1）", 3, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button40_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4; ApplyNumbering(sel.Range, "（%1）", 4, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button41_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4; ApplyNumbering(sel.Range, "（%1）", 5, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button42_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4; ApplyNumbering(sel.Range, "（%1）", 6, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button43_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4; ApplyNumbering(sel.Range, "（%1）", 7, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button44_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4; ApplyNumbering(sel.Range, "（%1）", 8, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button45_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4; ApplyNumbering(sel.Range, "（%1）", 9, WdListNumberStyle.wdListNumberStyleArabic); }
        private void button46_Click(object sender, RibbonControlEventArgs e) { Selection sel = Globals.ThisAddIn.Application.Selection; ApplyBasicParagraphStyle(sel.Range, "方正仿宋_GBK", 16f, 29f); sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4; ApplyNumbering(sel.Range, "（%1）", 10, WdListNumberStyle.wdListNumberStyleArabic); }

        // ==================== 表格处理 ====================
        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            WithScreenUpdatingDisabled(() =>
            {
                Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                if (doc.Tables.Count < 1) return;

                Table tbl = null;
                Range tr = null;
                try
                {
                    tbl = doc.Tables[1];
                    tr = tbl.Range;
                    tr.Font.Reset();
                    tr.ParagraphFormat.Reset();

                    tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                    tr.Font.Name = "宋体";
                    tr.Font.Size = 10.5f;
                    tr.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    tbl.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    tr.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast;
                    tr.ParagraphFormat.LineSpacing = 0;
                    tr.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
                    tr.ParagraphFormat.FirstLineIndent = 0f;
                    tr.ParagraphFormat.LeftIndent = 0f;
                    tr.ParagraphFormat.CharacterUnitLeftIndent = 0f;

                    tbl.Borders.Enable = 1;
                    tbl.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                    tbl.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;

                    if (tbl.Rows.Count >= 1)
                    {
                        tbl.Rows[1].Range.Font.Bold = (int)WdConstants.wdToggle;
                        tbl.Rows[1].HeadingFormat = (int)WdConstants.wdToggle;
                    }
                }
                finally
                {
                    if (tr != null) Marshal.ReleaseComObject(tr);
                    if (tbl != null) Marshal.ReleaseComObject(tbl);
                }
            });
        }

        private void button58_Click(object sender, RibbonControlEventArgs e)
        {
            WithScreenUpdatingDisabled(() =>
            {
                Document doc = Globals.ThisAddIn.Application.ActiveDocument;
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
                        tr.Font.Name = "宋体";
                        tr.Font.Size = 10.5f;
                        tr.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        tbl.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                        tr.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast;
                        tr.ParagraphFormat.LineSpacing = 0;
                        tr.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
                        tr.ParagraphFormat.FirstLineIndent = 0f;
                        tr.ParagraphFormat.LeftIndent = 0f;
                        tr.ParagraphFormat.CharacterUnitLeftIndent = 0f;

                        tbl.Borders.Enable = 1;
                        tbl.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                        tbl.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;

                        if (tbl.Rows.Count >= 1)
                        {
                            tbl.Rows[1].Range.Font.Bold = (int)WdConstants.wdToggle;
                            tbl.Rows[1].HeadingFormat = (int)WdConstants.wdToggle;
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
        }

        private void button13_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
            sel.Range.Font.Color = WdColor.wdColorAutomatic;
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
        }

        // ==================== 大纲级别单独设置 ====================
        private void button48_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
        }
        private void button49_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
        }
        private void button50_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
        }
        private void button51_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel4;
        }
        private void button52_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel5;
        }
        private void button53_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel6;
        }
        private void button54_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Paragraphs.OutlinePromote();
        }
        private void button55_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Paragraphs.OutlineDemote();
        }

        // ==================== 缩进控制 ====================
        /// <summary>设置首行缩进为 2 个字符单位（用于正文）。</summary>
        private void button56_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 2f;
        }

        /// <summary>取消首行与左缩进，恢复为无缩进状态。</summary>
        private void button57_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
            sel.Range.ParagraphFormat.FirstLineIndent = 0f;
            sel.Range.ParagraphFormat.LeftIndent = 0f;
            sel.Range.ParagraphFormat.CharacterUnitLeftIndent = 0f;
        }
    }
}