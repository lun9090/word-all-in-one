# 📄 办公助手

**Word 公文排版 VSTO 插件** —— 一键实现报告格式标准，大幅提升文档处理效率。

---

## ✨ 主要功能

| 功能类别 | 具体操作 | 快捷键/按钮 |
|----------|----------|-------------|
| **标题样式** | 方正小标宋、二号、居中，固定行距29磅 | `button1` |
| **正文样式** | 方正仿宋、三号、两端对齐，段前段后0磅，固定行距29磅 | `button2` |
| **一级标题（黑体）** | 方正黑体 + 自动编号（一、二、…） | `button3` ~ `button19`（起始1~10） |
| **二级标题（楷体）** | 方正楷体 + 自动编号（（一）、（二）、…） | `button8`、`button20~28` |
| **三级标题（仿宋）** | 方正仿宋 + 阿拉伯数字编号（1.、2.、…） | `button4`、`button29~37` |
| **四级标题（仿宋）** | 方正仿宋 + 括号编号（（1）、（2）、…） | `button9`、`button38~46` |
| **表格快速排版** | 自动调整列宽、宋体五号、居中、边框、标题行加粗并重复 | `button10`（当前表）<br>`button58`（全部表） |
| **页面精细设置** | A4、边距（上37下35左28右26mm）、页脚24.7mm、每页22行网格、禁止标点溢出、修改正文样式使后续段落自动继承 | `button12` |
| **标记与颜色** | 黄色高亮 + 红色字体 / 清除标记 | `button11` / `button13` |
| **查找替换** | 删除多余换行符（保留中文标点后一个回车） | `button47` |
| **大纲级别** | 一键设置为1~6级，或升降级 | `button48~55` |
| **缩进控制** | 首行缩进2字符 / 取消所有缩进 | `button56` / `button57` |

---

## 🚀 快速开始

### 环境要求
- **操作系统**：Windows 7 及以上
- **Office**：Microsoft Office 2010 及以上（支持 Word）
- **.NET Framework**：4.6.1 或更高
- **开发工具**：Visual Studio 2019/2022（含 VSTO 工作负载）

### 安装与使用
1. 下载本仓库源码或 Release 安装包。
2. 使用 Visual Studio 打开解决方案，编译生成。
3. 运行安装程序（.vsto 或 .msi），插件会自动注册到 Word。
4. 启动 Word，在 **“加载项”** 选项卡中即可看到 **“办公助手”** 功能区，点击相应按钮即可应用样式。

> 💡 **提示**：首次使用建议先点击 `button12`（页面设置），让后续所有段落自动继承标点控制和首行缩进。

---

## 📐 设计理念

- **高效批处理**：支持批量处理整个文档的表格、段落和样式。
- **性能优化**：关闭屏幕更新、显式释放 COM 对象，处理百页文档仍保持流畅。
- **易扩展**：代码分层清晰，新增样式或编号只需调用核心方法。

---

## 🛠️ 技术栈

- **语言**：C#
- **框架**：.NET Framework 4.6.1
- **Office 互操作**：Microsoft.Office.Interop.Word
- **开发模型**：VSTO（Visual Studio Tools for Office）
- **设计模式**：辅助方法封装 COM 操作，资源管理使用 `try-finally` + `Marshal.ReleaseComObject`

---

## 📁 项目结构

```
办公助手/
├── Ribbon1.cs              # 功能区按钮逻辑（核心）
├── ThisAddIn.cs            # 插件启动/关闭
├── Ribbon1.Designer.cs     # 功能区设计器（自动生成）
├── Properties/             # 程序集信息
└── Resources/              # 图标等资源
```

---

## 🤝 贡献指南

欢迎提交 Issue 或 Pull Request 来帮助改进！

1. Fork 本项目。
2. 创建您的特性分支 (`git checkout -b feature/AmazingFeature`)。
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)。
4. 推送到分支 (`git push origin feature/AmazingFeature`)。
5. 打开一个 Pull Request。

---

## 📄 许可证

本项目基于 **Apache-2.0 license** 开源，详情请见 [LICENSE](LICENSE) 文件。

---

## 📞 联系与支持

- 作者：lun9090
- 如有问题，请在此仓库提交 [Issues](https://github.com/lun9090/word-all-in-one/issues)。

---

**让公文排版变得如此简单，助力高效办公！** 🎉
