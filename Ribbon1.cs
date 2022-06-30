using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Word;
using Document = Microsoft.Office.Interop.Word.Document;
using System.Windows.Forms;


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
           
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清理所有格式
            sel.ClearFormatting();
            //根据窗口调整单元格
            sel.Tables[1].AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
            //宋体，五号，行距固定值18磅，垂直对齐：居中。
            sel.Tables[1].Range.Font.Name= "宋体";
            sel.Tables[1].Range.ParagraphFormat.Alignment= WdParagraphAlignment.wdAlignParagraphCenter;
            sel.Tables[1].Range.Font.Size = 10.5f;
            sel.Cells.VerticalAlignment = (WdCellVerticalAlignment)1;
            sel.Tables[1].Range.ParagraphFormat.LineSpacing = 18;
            
            //表格属性：段落：无缩进，段前0行，段后0行
            sel.Tables[1].Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
            sel.Tables[1].Range.ParagraphFormat.FirstLineIndent = 0f;
            sel.Tables[1].Range.ParagraphFormat.LeftIndent = 0f;
            sel.Tables[1].Range.ParagraphFormat.CharacterUnitLeftIndent = 0f;

            //默认创建的表格没有边框，这里修改其属性，使得创建的表格带有边框 
            sel.Tables[1].Borders.Enable = 1;//这个值可以设置得很大，例如5、13等等

            // 设置table边框样式
            sel.Tables[1].Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;//表格外框是单线
            sel.Tables[1].Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;//表格内框是单线
            //表格内容加粗
            sel.Tables[1].Cell(1, 1).Select();
            sel.SelectRow();
            sel.Range.Font.Bold = (int)WdConstants.wdToggle;
            //表格重复标题行
            //sel.Tables[1].Cell(1, 1).Select();
            //sel.MoveRight();
            //sel.Rows.HeadingFormat = (int)WdConstants.wdToggle;
            sel.Tables[1].Cell(1, 1).Select();
            sel.MoveRight();
            sel.MoveLeft();
            sel.Range.Rows.HeadingFormat = (int)WdConstants.wdToggle;
            sel.Range.Rows.HeadingFormat = (int)WdConstants.wdToggle;
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

        private void button48_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别1
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1;
        }

        private void button49_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别2
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2;
        }

        private void button50_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别3
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3;
        }

        private void button51_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别4
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4;
        }

        private void button52_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别5
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel5;
        }

        private void button53_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //清除大纲级别
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //设置大纲级别6
            sel.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel6;
        }

        private void button54_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //大纲升级
            sel.Paragraphs.OutlinePromote();
        }

        private void button55_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //大纲升降级
            sel.Paragraphs.OutlineDemote();
        }

        private void button56_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //首行2字符
            sel.ParagraphFormat.CharacterUnitFirstLineIndent = float.Parse("2");
        }

        private void button57_Click(object sender, RibbonControlEventArgs e)
        {
            //选择内容
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection sel = Globals.ThisAddIn.Application.Selection;
            //段落顶格
            sel.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
            sel.ParagraphFormat.FirstLineIndent = 0f;
            sel.ParagraphFormat.LeftIndent = 0f;
            sel.ParagraphFormat.CharacterUnitLeftIndent = 0f;
        }

        private void button58_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            
            for (int i = 1;  i<= doc.Tables.Count; ++i)
            {

                doc.Tables[i].Select();
                Selection sel = Globals.ThisAddIn.Application.Selection;
                //清理所有格式
                sel.ClearFormatting();
                //根据窗口调整单元格
                sel.Tables[1].AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                //宋体，五号，行距固定值18磅，垂直对齐：居中。
                sel.Tables[1].Range.Font.Name = "宋体";
                sel.Tables[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                sel.Tables[1].Range.Font.Size = 10.5f;
                sel.Cells.VerticalAlignment = (WdCellVerticalAlignment)1;
                sel.Tables[1].Range.ParagraphFormat.LineSpacing = 18;

                //表格属性：段落：无缩进，段前0行，段后0行
                sel.Tables[1].Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
                sel.Tables[1].Range.ParagraphFormat.FirstLineIndent = 0f;
                sel.Tables[1].Range.ParagraphFormat.LeftIndent = 0f;
                sel.Tables[1].Range.ParagraphFormat.CharacterUnitLeftIndent = 0f;

                //默认创建的表格没有边框，这里修改其属性，使得创建的表格带有边框 
                sel.Tables[1].Borders.Enable = 1;//这个值可以设置得很大，例如5、13等等

                // 设置table边框样式
                sel.Tables[1].Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;//表格外框是单线
                sel.Tables[1].Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;//表格内框是单线
                                                                                      //表格内容加粗
                sel.Tables[1].Cell(1, 1).Select();
                sel.SelectRow();
                sel.Range.Font.Bold = (int)WdConstants.wdToggle;
                //表格重复标题行
                //sel.Tables[1].Cell(1, 1).Select();
                //sel.MoveRight();
                //sel.Rows.HeadingFormat = (int)WdConstants.wdToggle;
                sel.Tables[1].Cell(1, 1).Select();
                sel.MoveRight();
                sel.MoveLeft();
                sel.Range.Rows.HeadingFormat = (int)WdConstants.wdToggle;
                sel.Range.Rows.HeadingFormat = (int)WdConstants.wdToggle;
            }
        }
    }
}
