# Factorio Mods Locale Tool
## 异星工厂模组汉化工具
- [异星工厂吧](http://tieba.baidu.com/f?kw=factorio)出品
- 用搜集而来的文本库批量翻译异星工厂的模组，方便地共享模组汉化。

### 如何下载本工具？
在本项目的附件中下载本工具的最新发布版本。

### 如何使用本工具？
首先把脚本`[FMLT] Script.vbs`和文本库`[FMLT] Library for zh-CN`解压到异星工厂[用户数据目录](https://wiki.factorio.com/index.php?title=Application_directory/zh)下。

- 直接双击运行脚本`[FMLT] Script.vbs`即可自动翻译同目录下`\mods`内的所有模组。
- 可能包含模组的zip压缩包或文件夹也可以直接拖到脚本`[FMLT] Script.vbs`上面，识别后即可翻译。
- 把文本库拖到脚本`[FMLT] Script.vbs`上可根据文本库生成一个带翻译的模组列表。

### 如何自己汉化模组？
对于某个模组，用本工具汉化后发现汉化率过低，就需要你自己来为其增添汉化了。

1. 用文本编辑器（不要用记事本，推荐Notepad++）打开翻译文件，编辑后以`UTF-8 without BOM`编码保存。
2. 每个模组可翻译的部分包括：
- `\info.json` 其中的 "title"(MOD名) 和 "description"(简介) 
- `\locale\zh-CN\*.cfg` 

### 如何分享自己的汉化？
辛辛苦苦自己汉化的MOD当然不希望独享，这个工具就是帮你干这个事的。

- 联系文本库管理员 Quiet95scholar(QQ-211398520, linshuboy@qq.com)，他会把你提交的汉化加入文本库。
- 如果你会使用Git，也可以本项目的 develop 分支为基础，向其发起 Pull Request 。

### 其它说明
- 如有疑问，可在本项目中发起 Issue 或直接联系管理员。
- 分支说明：`master`分支为阶段性发布分支，`develop`分支为汉化搜集分支。
