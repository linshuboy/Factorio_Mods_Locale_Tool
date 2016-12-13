## 异星工厂模组汉化工具
- [异星工厂吧](http://tieba.baidu.com/f?kw=factorio)出品
- 用搜集而来的文本库批量翻译异星工厂的模组，方便地共享模组汉化。

### 如何下载本工具？
- 正式发布版本可在[百度网盘](http://pan.baidu.com/s/1pJ1MEVL)下载。
- 也可以登录后直接下载本项目文件。

### 如何使用本工具？
首先把脚本`[FMLT] Script.vbs`和文本库`[FMLT] Library for zh-CN`解压到异星工厂[用户数据目录](https://wiki.factorio.com/index.php?title=Application_directory/zh)下。

- 直接双击运行脚本`[FMLT] Script.vbs`即可自动翻译同目录下`\mods`内的所有模组。
- 可能包含模组的zip压缩包或文件夹也可以直接拖到脚本`[FMLT] Script.vbs`上面，识别后即可翻译。
- 把文本库拖到脚本`[FMLT] Script.vbs`上可根据文本库生成带翻译的模组列表`mods-list.txt`。

### 如何自己汉化模组？
对于某个模组，用本工具汉化后发现汉化率过低，就需要你自己来为其增添汉化了。

1. 用文本编辑器（不要用记事本，推荐Notepad++）打开翻译文件，编辑后以`UTF-8 without BOM`编码保存。
2. 每个模组可翻译的部分包括：

- `\info.json` 其中的 "title"(MOD名) 和 "description"(简介) 
- `\locale\zh-CN\*.cfg` 

### 如何分享自己的汉化？
辛辛苦苦自己汉化的MOD当然不希望独享，这个工具就是帮你干这个事的。

- 联系文本库管理员 Quiet95scholar(QQ:211398520, linshuboy@qq.com)，他会把你提交的汉化加入文本库。
- 最有效的提交方式，将汉化完成的整个MOD作为附件发给文本库管理员 Quiet95scholar(linshuboy@qq.com)。
- 在联系管理员之前，我倡议你为你汉化的MOD准备一个更新信息，并且在自己的贡献署名。
- 如果你会使用Git，也可以本项目的 develop 分支为基础，向其发起 Pull Request 。

### 其它说明
- 如有疑问，可在本项目中发起 Issue 或直接联系管理员。
- 分支说明：`master`分支为阶段性发布分支，`develop`分支为汉化搜集分支。

#### 伟大的参与者们(排名不分先后)

- Gao
- fubixingwzy
- 好人Ⅲ(DEMO)
- BlueSky234
- 稀饭加糖
- Satan'心
- 寒冰之幽梦
- 长空X
- 让你贱笑了/jy
- 冷月无声
- Quiet95scholar
- Mr.Jos
- tpz
- 教皇
- 普宁老兵诊锁
- 被54的小怀