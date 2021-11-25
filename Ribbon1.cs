using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Word;
using Document = Microsoft.Office.Interop.Word.Document;


namespace 李艇的办公助手
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //居中
            sel.ParagraphFormat.Alignment=WdParagraphAlignment.wdAlignParagraphCenter;
            //标题字体
            sel.Font.Name = "方正小标宋_GBK";
            //字体大小二号
            sel.Font.Size = 22;
            //1.5倍间距
            //sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpace1pt5;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
        }

        private static ThisAddIn GetThisAddIn()
        {
            return Globals.ThisAddIn;
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //首行2字符
            sel.ParagraphFormat.CharacterUnitFirstLineIndent = float.Parse("2");
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正黑体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别1
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];            
            lt.ListLevels[1].NumberFormat = "%1、";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].ResetOnHigher=0;
            lt.ListLevels[1].StartAt = 1;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }
        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正黑体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别1
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1、";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 2;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正黑体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别1
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1、";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 3;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正黑体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别1
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1、";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 4;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正楷体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别2
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 1;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别3
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1．";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 1;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别4
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 1;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //居中
            sel.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            //标题字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小四号
            sel.Font.Size = 14;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //文字与单元格中心对齐
            sel.Cells.VerticalAlignment = (WdCellVerticalAlignment)1;

        }

        private void button11_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //设置文字背景为标黄
            sel.Range.HighlightColorIndex = WdColorIndex.wdYellow;
            //设置文字颜色为标红
            sel.Range.Font.Color= WdColor.wdColorRed;
        }

        private void button12_Click(object sender, RibbonControlEventArgs e)
        {
            //初始化页面
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            //设置纸张样式为A4纸
            doc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
            //排列方式为垂直方向
            doc.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
            //边距—上35mm、下32mm、左28mm、右26mm
            doc.PageSetup.TopMargin = Globals.ThisAddIn.Application.CentimetersToPoints(float.Parse("3.5"));
            doc.PageSetup.BottomMargin = Globals.ThisAddIn.Application.CentimetersToPoints(float.Parse("3.2"));
            doc.PageSetup.LeftMargin = Globals.ThisAddIn.Application.CentimetersToPoints(float.Parse("2.8"));
            doc.PageSetup.RightMargin = Globals.ThisAddIn.Application.CentimetersToPoints(float.Parse("2.6"));
        }

        private void button13_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //设置文字背景为默认
            sel.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
            //设置文字颜色为默认
            sel.Range.Font.Color = WdColor.wdColorAutomatic;

        }

        private void button14_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正黑体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别1
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1、";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].ResetOnHigher = 0;
            lt.ListLevels[1].StartAt = 5;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button15_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正黑体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别1
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1、";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].ResetOnHigher = 0;
            lt.ListLevels[1].StartAt = 6;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button16_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正黑体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别1
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1、";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].ResetOnHigher = 0;
            lt.ListLevels[1].StartAt = 7;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button17_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正黑体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别1
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1、";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].ResetOnHigher = 0;
            lt.ListLevels[1].StartAt = 8;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button18_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正黑体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别1
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1、";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].ResetOnHigher = 0;
            lt.ListLevels[1].StartAt = 9;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button19_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正黑体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别1
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1、";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].ResetOnHigher = 0;
            lt.ListLevels[1].StartAt = 10;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button20_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正楷体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别2
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 2;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button21_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正楷体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别2
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 3;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button22_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正楷体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别2
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 4;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button23_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正楷体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别2
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 5;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button24_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正楷体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别2
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 6;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button25_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正楷体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别2
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 7;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button26_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正楷体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别2
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 8;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button27_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正楷体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别2
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 9;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button28_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正楷体_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别2
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)39;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 10;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button29_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别3
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1．";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 2;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button30_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别3
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1．";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 3;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button31_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别3
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1．";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 4;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button32_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别3
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1．";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 5;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button33_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别3
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1．";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 6;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button34_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别3
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1．";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 7;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button35_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别3
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1．";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 8;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button36_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别3
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1．";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 9;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button37_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别3
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "%1．";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 10;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button38_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别4
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 2;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button39_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别4
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 3;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button40_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别4
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 4;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button41_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别4
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 5;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button42_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别4
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 6;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button43_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别4
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 7;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button45_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别4
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 9;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button44_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别4
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 8;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button46_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //正文字体
            sel.Font.Name = "方正仿宋_GBK";
            //字体大小三号
            sel.Font.Size = 16;
            //行距—固定值29磅
            sel.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
            sel.ParagraphFormat.LineSpacing = 29;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别4
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4;
            //设置编号
            Microsoft.Office.Interop.Word.ListTemplate lt = Globals.ThisAddIn.Application.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[1];
            lt.ListLevels[1].NumberFormat = "（%1）";
            lt.ListLevels[1].TrailingCharacter = (WdTrailingCharacter)2;
            lt.ListLevels[1].NumberStyle = (WdListNumberStyle)0;
            lt.ListLevels[1].NumberPosition = 0;
            lt.ListLevels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            lt.ListLevels[1].TextPosition = Globals.ThisAddIn.Application.CentimetersToPoints(0);
            lt.ListLevels[1].TabPosition = (float)Microsoft.Office.Interop.Word.WdConstants.wdUndefined;
            lt.ListLevels[1].StartAt = 10;
            object bContinuePrevList = false;
            object applyTo = Microsoft.Office.Interop.Word.WdListApplyTo.wdListApplyToWholeList;
            object defBehavior = Microsoft.Office.Interop.Word.WdDefaultListBehavior.wdWord9ListBehavior;
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplateWithLevel(lt, bContinuePrevList, applyTo, defBehavior);
        }

        private void button47_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
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
            sel.Find.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);


        }
    }
}
