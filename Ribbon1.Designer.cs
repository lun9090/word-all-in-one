
namespace 李艇的办公助手
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.button12 = this.Factory.CreateRibbonButton();
            this.menu5 = this.Factory.CreateRibbonMenu();
            this.button47 = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button56 = this.Factory.CreateRibbonButton();
            this.button57 = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button14 = this.Factory.CreateRibbonButton();
            this.button15 = this.Factory.CreateRibbonButton();
            this.button16 = this.Factory.CreateRibbonButton();
            this.button17 = this.Factory.CreateRibbonButton();
            this.button18 = this.Factory.CreateRibbonButton();
            this.button19 = this.Factory.CreateRibbonButton();
            this.menu3 = this.Factory.CreateRibbonMenu();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button20 = this.Factory.CreateRibbonButton();
            this.button21 = this.Factory.CreateRibbonButton();
            this.button22 = this.Factory.CreateRibbonButton();
            this.button23 = this.Factory.CreateRibbonButton();
            this.button24 = this.Factory.CreateRibbonButton();
            this.button25 = this.Factory.CreateRibbonButton();
            this.button26 = this.Factory.CreateRibbonButton();
            this.button27 = this.Factory.CreateRibbonButton();
            this.button28 = this.Factory.CreateRibbonButton();
            this.menu2 = this.Factory.CreateRibbonMenu();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button29 = this.Factory.CreateRibbonButton();
            this.button30 = this.Factory.CreateRibbonButton();
            this.button31 = this.Factory.CreateRibbonButton();
            this.button32 = this.Factory.CreateRibbonButton();
            this.button33 = this.Factory.CreateRibbonButton();
            this.button34 = this.Factory.CreateRibbonButton();
            this.button35 = this.Factory.CreateRibbonButton();
            this.button36 = this.Factory.CreateRibbonButton();
            this.button37 = this.Factory.CreateRibbonButton();
            this.menu4 = this.Factory.CreateRibbonMenu();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button38 = this.Factory.CreateRibbonButton();
            this.button39 = this.Factory.CreateRibbonButton();
            this.button40 = this.Factory.CreateRibbonButton();
            this.button41 = this.Factory.CreateRibbonButton();
            this.button42 = this.Factory.CreateRibbonButton();
            this.button43 = this.Factory.CreateRibbonButton();
            this.button44 = this.Factory.CreateRibbonButton();
            this.button45 = this.Factory.CreateRibbonButton();
            this.button46 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button11 = this.Factory.CreateRibbonButton();
            this.button13 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.menu6 = this.Factory.CreateRibbonMenu();
            this.button48 = this.Factory.CreateRibbonButton();
            this.button49 = this.Factory.CreateRibbonButton();
            this.button50 = this.Factory.CreateRibbonButton();
            this.button51 = this.Factory.CreateRibbonButton();
            this.button52 = this.Factory.CreateRibbonButton();
            this.button53 = this.Factory.CreateRibbonButton();
            this.button54 = this.Factory.CreateRibbonButton();
            this.button55 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group5.SuspendLayout();
            this.group1.SuspendLayout();
            this.group6.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "办公助手";
            this.tab1.Name = "tab1";
            // 
            // group5
            // 
            this.group5.Items.Add(this.button12);
            this.group5.Items.Add(this.menu5);
            this.group5.Name = "group5";
            // 
            // button12
            // 
            this.button12.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button12.Image = ((System.Drawing.Image)(resources.GetObject("button12.Image")));
            this.button12.Label = "页面设置";
            this.button12.Name = "button12";
            this.button12.ShowImage = true;
            this.button12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button12_Click);
            // 
            // menu5
            // 
            this.menu5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu5.Image = ((System.Drawing.Image)(resources.GetObject("menu5.Image")));
            this.menu5.Items.Add(this.button47);
            this.menu5.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu5.Label = "排版";
            this.menu5.Name = "menu5";
            this.menu5.ShowImage = true;
            // 
            // button47
            // 
            this.button47.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button47.Label = "断行重排";
            this.button47.Name = "button47";
            this.button47.ShowImage = true;
            this.button47.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button47_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.button56);
            this.group1.Items.Add(this.button57);
            this.group1.Name = "group1";
            // 
            // button56
            // 
            this.button56.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button56.Image = global::李艇的办公助手.Properties.Resources.右缩进;
            this.button56.Label = "缩进";
            this.button56.Name = "button56";
            this.button56.ShowImage = true;
            this.button56.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button56_Click);
            // 
            // button57
            // 
            this.button57.Image = global::李艇的办公助手.Properties.Resources.左缩进;
            this.button57.Label = "顶格";
            this.button57.Name = "button57";
            this.button57.ShowImage = true;
            this.button57.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button57_Click);
            // 
            // group6
            // 
            this.group6.Items.Add(this.button1);
            this.group6.Items.Add(this.button10);
            this.group6.Items.Add(this.button2);
            this.group6.Name = "group6";
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "标题";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);
            // 
            // button10
            // 
            this.button10.Image = ((System.Drawing.Image)(resources.GetObject("button10.Image")));
            this.button10.Label = "表格";
            this.button10.Name = "button10";
            this.button10.ShowImage = true;
            this.button10.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button10_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "正文";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.menu1);
            this.group2.Items.Add(this.menu3);
            this.group2.Items.Add(this.menu2);
            this.group2.Items.Add(this.menu4);
            this.group2.Name = "group2";
            // 
            // menu1
            // 
            this.menu1.Image = ((System.Drawing.Image)(resources.GetObject("menu1.Image")));
            this.menu1.Items.Add(this.button3);
            this.menu1.Items.Add(this.button5);
            this.menu1.Items.Add(this.button6);
            this.menu1.Items.Add(this.button7);
            this.menu1.Items.Add(this.button14);
            this.menu1.Items.Add(this.button15);
            this.menu1.Items.Add(this.button16);
            this.menu1.Items.Add(this.button17);
            this.menu1.Items.Add(this.button18);
            this.menu1.Items.Add(this.button19);
            this.menu1.Label = "标题一、";
            this.menu1.Name = "menu1";
            this.menu1.ShowImage = true;
            // 
            // button3
            // 
            this.button3.Label = "标题一、";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button5
            // 
            this.button5.Label = "标题二、";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Label = "标题三、";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Label = "标题四、";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // button14
            // 
            this.button14.Label = "标题五、";
            this.button14.Name = "button14";
            this.button14.ShowImage = true;
            this.button14.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button14_Click);
            // 
            // button15
            // 
            this.button15.Label = "标题六、";
            this.button15.Name = "button15";
            this.button15.ShowImage = true;
            this.button15.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button15_Click);
            // 
            // button16
            // 
            this.button16.Label = "标题七、";
            this.button16.Name = "button16";
            this.button16.ShowImage = true;
            this.button16.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button16_Click);
            // 
            // button17
            // 
            this.button17.Label = "标题八、";
            this.button17.Name = "button17";
            this.button17.ShowImage = true;
            this.button17.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button17_Click);
            // 
            // button18
            // 
            this.button18.Label = "标题九、";
            this.button18.Name = "button18";
            this.button18.ShowImage = true;
            this.button18.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button18_Click);
            // 
            // button19
            // 
            this.button19.Label = "标题十、";
            this.button19.Name = "button19";
            this.button19.ShowImage = true;
            this.button19.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button19_Click);
            // 
            // menu3
            // 
            this.menu3.Image = ((System.Drawing.Image)(resources.GetObject("menu3.Image")));
            this.menu3.Items.Add(this.button8);
            this.menu3.Items.Add(this.button20);
            this.menu3.Items.Add(this.button21);
            this.menu3.Items.Add(this.button22);
            this.menu3.Items.Add(this.button23);
            this.menu3.Items.Add(this.button24);
            this.menu3.Items.Add(this.button25);
            this.menu3.Items.Add(this.button26);
            this.menu3.Items.Add(this.button27);
            this.menu3.Items.Add(this.button28);
            this.menu3.Label = "标题（一）";
            this.menu3.Name = "menu3";
            this.menu3.ShowImage = true;
            // 
            // button8
            // 
            this.button8.Label = "标题（一）";
            this.button8.Name = "button8";
            this.button8.ShowImage = true;
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button8_Click);
            // 
            // button20
            // 
            this.button20.Label = "标题（二）";
            this.button20.Name = "button20";
            this.button20.ShowImage = true;
            this.button20.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button20_Click);
            // 
            // button21
            // 
            this.button21.Label = "标题（三）";
            this.button21.Name = "button21";
            this.button21.ShowImage = true;
            this.button21.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button21_Click);
            // 
            // button22
            // 
            this.button22.Label = "标题（四）";
            this.button22.Name = "button22";
            this.button22.ShowImage = true;
            this.button22.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button22_Click);
            // 
            // button23
            // 
            this.button23.Label = "标题（五）";
            this.button23.Name = "button23";
            this.button23.ShowImage = true;
            this.button23.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button23_Click);
            // 
            // button24
            // 
            this.button24.Label = "标题（六）";
            this.button24.Name = "button24";
            this.button24.ShowImage = true;
            this.button24.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button24_Click);
            // 
            // button25
            // 
            this.button25.Label = "标题（七）";
            this.button25.Name = "button25";
            this.button25.ShowImage = true;
            this.button25.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button25_Click);
            // 
            // button26
            // 
            this.button26.Label = "标题（八）";
            this.button26.Name = "button26";
            this.button26.ShowImage = true;
            this.button26.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button26_Click);
            // 
            // button27
            // 
            this.button27.Label = "标题（九）";
            this.button27.Name = "button27";
            this.button27.ShowImage = true;
            this.button27.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button27_Click);
            // 
            // button28
            // 
            this.button28.Label = "标题（十）";
            this.button28.Name = "button28";
            this.button28.ShowImage = true;
            this.button28.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button28_Click);
            // 
            // menu2
            // 
            this.menu2.Image = ((System.Drawing.Image)(resources.GetObject("menu2.Image")));
            this.menu2.Items.Add(this.button4);
            this.menu2.Items.Add(this.button29);
            this.menu2.Items.Add(this.button30);
            this.menu2.Items.Add(this.button31);
            this.menu2.Items.Add(this.button32);
            this.menu2.Items.Add(this.button33);
            this.menu2.Items.Add(this.button34);
            this.menu2.Items.Add(this.button35);
            this.menu2.Items.Add(this.button36);
            this.menu2.Items.Add(this.button37);
            this.menu2.Label = "标题1.";
            this.menu2.Name = "menu2";
            this.menu2.ShowImage = true;
            // 
            // button4
            // 
            this.button4.Label = "标题1.";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // button29
            // 
            this.button29.Label = "标题2.";
            this.button29.Name = "button29";
            this.button29.ShowImage = true;
            this.button29.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button29_Click);
            // 
            // button30
            // 
            this.button30.Label = "标题3.";
            this.button30.Name = "button30";
            this.button30.ShowImage = true;
            this.button30.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button30_Click);
            // 
            // button31
            // 
            this.button31.Label = "标题4.";
            this.button31.Name = "button31";
            this.button31.ShowImage = true;
            this.button31.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button31_Click);
            // 
            // button32
            // 
            this.button32.Label = "标题5.";
            this.button32.Name = "button32";
            this.button32.ShowImage = true;
            this.button32.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button32_Click);
            // 
            // button33
            // 
            this.button33.Label = "标题6.";
            this.button33.Name = "button33";
            this.button33.ShowImage = true;
            this.button33.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button33_Click);
            // 
            // button34
            // 
            this.button34.Label = "标题7.";
            this.button34.Name = "button34";
            this.button34.ShowImage = true;
            this.button34.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button34_Click);
            // 
            // button35
            // 
            this.button35.Label = "标题8.";
            this.button35.Name = "button35";
            this.button35.ShowImage = true;
            this.button35.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button35_Click);
            // 
            // button36
            // 
            this.button36.Label = "标题9.";
            this.button36.Name = "button36";
            this.button36.ShowImage = true;
            this.button36.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button36_Click);
            // 
            // button37
            // 
            this.button37.Label = "标题10.";
            this.button37.Name = "button37";
            this.button37.ShowImage = true;
            this.button37.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button37_Click);
            // 
            // menu4
            // 
            this.menu4.Image = ((System.Drawing.Image)(resources.GetObject("menu4.Image")));
            this.menu4.Items.Add(this.button9);
            this.menu4.Items.Add(this.button38);
            this.menu4.Items.Add(this.button39);
            this.menu4.Items.Add(this.button40);
            this.menu4.Items.Add(this.button41);
            this.menu4.Items.Add(this.button42);
            this.menu4.Items.Add(this.button43);
            this.menu4.Items.Add(this.button44);
            this.menu4.Items.Add(this.button45);
            this.menu4.Items.Add(this.button46);
            this.menu4.Label = "标题（1）";
            this.menu4.Name = "menu4";
            this.menu4.ShowImage = true;
            // 
            // button9
            // 
            this.button9.Label = "标题（1）";
            this.button9.Name = "button9";
            this.button9.ShowImage = true;
            this.button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button9_Click);
            // 
            // button38
            // 
            this.button38.Label = "标题（2）";
            this.button38.Name = "button38";
            this.button38.ShowImage = true;
            this.button38.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button38_Click);
            // 
            // button39
            // 
            this.button39.Label = "标题（3）";
            this.button39.Name = "button39";
            this.button39.ShowImage = true;
            this.button39.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button39_Click);
            // 
            // button40
            // 
            this.button40.Label = "标题（4）";
            this.button40.Name = "button40";
            this.button40.ShowImage = true;
            this.button40.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button40_Click);
            // 
            // button41
            // 
            this.button41.Label = "标题（5）";
            this.button41.Name = "button41";
            this.button41.ShowImage = true;
            this.button41.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button41_Click);
            // 
            // button42
            // 
            this.button42.Label = "标题（6）";
            this.button42.Name = "button42";
            this.button42.ShowImage = true;
            this.button42.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button42_Click);
            // 
            // button43
            // 
            this.button43.Label = "标题（7）";
            this.button43.Name = "button43";
            this.button43.ShowImage = true;
            this.button43.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button43_Click);
            // 
            // button44
            // 
            this.button44.Label = "标题（8）";
            this.button44.Name = "button44";
            this.button44.ShowImage = true;
            this.button44.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button44_Click);
            // 
            // button45
            // 
            this.button45.Label = "标题（9）";
            this.button45.Name = "button45";
            this.button45.ShowImage = true;
            this.button45.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button45_Click);
            // 
            // button46
            // 
            this.button46.Label = "标题（10）";
            this.button46.Name = "button46";
            this.button46.ShowImage = true;
            this.button46.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button46_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button11);
            this.group3.Items.Add(this.button13);
            this.group3.Name = "group3";
            // 
            // button11
            // 
            this.button11.Image = ((System.Drawing.Image)(resources.GetObject("button11.Image")));
            this.button11.Label = "标黄标红";
            this.button11.Name = "button11";
            this.button11.ShowImage = true;
            this.button11.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button11_Click);
            // 
            // button13
            // 
            this.button13.Image = ((System.Drawing.Image)(resources.GetObject("button13.Image")));
            this.button13.Label = "恢复默认";
            this.button13.Name = "button13";
            this.button13.ShowImage = true;
            this.button13.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button13_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.menu6);
            this.group4.Items.Add(this.button54);
            this.group4.Items.Add(this.button55);
            this.group4.Name = "group4";
            // 
            // menu6
            // 
            this.menu6.Image = ((System.Drawing.Image)(resources.GetObject("menu6.Image")));
            this.menu6.Items.Add(this.button48);
            this.menu6.Items.Add(this.button49);
            this.menu6.Items.Add(this.button50);
            this.menu6.Items.Add(this.button51);
            this.menu6.Items.Add(this.button52);
            this.menu6.Items.Add(this.button53);
            this.menu6.Label = "大纲级别一、";
            this.menu6.Name = "menu6";
            this.menu6.ShowImage = true;
            // 
            // button48
            // 
            this.button48.Label = "大纲级别一、";
            this.button48.Name = "button48";
            this.button48.ShowImage = true;
            this.button48.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button48_Click);
            // 
            // button49
            // 
            this.button49.Label = "大纲级别二、";
            this.button49.Name = "button49";
            this.button49.ShowImage = true;
            this.button49.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button49_Click);
            // 
            // button50
            // 
            this.button50.Label = "大纲级别三、";
            this.button50.Name = "button50";
            this.button50.ShowImage = true;
            this.button50.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button50_Click);
            // 
            // button51
            // 
            this.button51.Label = "大纲级别四、";
            this.button51.Name = "button51";
            this.button51.ShowImage = true;
            this.button51.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button51_Click);
            // 
            // button52
            // 
            this.button52.Label = "大纲级别五、";
            this.button52.Name = "button52";
            this.button52.ShowImage = true;
            this.button52.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button52_Click);
            // 
            // button53
            // 
            this.button53.Label = "大纲级别六、";
            this.button53.Name = "button53";
            this.button53.ShowImage = true;
            this.button53.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button53_Click);
            // 
            // button54
            // 
            this.button54.Image = global::李艇的办公助手.Properties.Resources.上升;
            this.button54.Label = "大纲升级";
            this.button54.Name = "button54";
            this.button54.ShowImage = true;
            this.button54.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button54_Click);
            // 
            // button55
            // 
            this.button55.Image = global::李艇的办公助手.Properties.Resources.下降;
            this.button55.Label = "大纲降级";
            this.button55.Name = "button55";
            this.button55.ShowImage = true;
            this.button55.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button55_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button13;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button14;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button15;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button16;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button17;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button18;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button19;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button20;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button21;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button22;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button23;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button24;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button25;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button26;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button27;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button28;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button29;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button30;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button31;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button32;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button33;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button34;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button35;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button36;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button37;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button38;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button39;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button40;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button41;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button42;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button43;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button44;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button45;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button46;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button47;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button48;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button49;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button50;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button51;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button52;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button53;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button54;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button55;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button56;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button57;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
