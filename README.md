# 移除Office Word文件的指定字符小工具

#### 介绍
技术奇客出品的Office系列小工具 - 移除Office Word文件的指定字符。可指定目录和自定义移除字符列表， 可批量处理。ITGeeker技术奇客是奇客罗方智能科技的关联公司。

#### 软件架构
软件采用Python 3.11版本开发，可以运行于Windows 10/11，也可以运行于Linux系统。Windows系统请下载发行版的exe可执行文件即可，Linux理论上只要是Python 3的环境都可直接运行。

#### 运行及使用说明

> Windows版本

    1.  下载可执行文件
    2.  双击文件并执行
    3.  添加想要移除的文本列表
    4.  选择要处理的文档所在的目录
    5.  开始处理

> Linux版本

    1.  确定你又安装Python 3版本，最好3.8以上版本
    2.  安装Python依赖："pip install python-docx tkinter"
    3.  下载本项目到本地目录，并运行："python remove_str_from_word_main.py"

处理过的文件将被保存到子目录“已处理文件”当中，文件名末尾附上“-revised”字样以示区别。

#### 更新日志

> 2023-06-09 v1.0.1.0
 
    1. 调整替换算法，让替换忽略大小写并且更加准确
    2. 如果是doc文件，先转换为docx文件; 转换前检查之前是否已转换
    3. Word文件跳过处理office产生的临时文件（~$*.*）
    4. 文件获取不包含子目录，只处理当前根目录的文件
    5. 添加需处理的字符时，自动去除前后空格
    6. 修复文件名处理时的包含了扩展名错误
    7. 修复首次运行没有默认字符的错误，若把之前的staff_mobile_email.txt放到文件的同一目录，则会自动导入
    8. 任务成功与否都会弹窗提醒

> 2023-06-08 v1.0.0.0
 
    1. 文本列表可自动保存，并在启动时自动加载
    2. 运行移除后将在文件目录创建“已处理文件”子目录
    3. 已处理文件名附件“-Revised”字样，以示区别


#### 参与贡献

    1.  Fork 本仓库
    2.  新建 Feat_xxx 分支
    3.  提交代码
    4.  新建 Pull Request


#### 其他

    1.  阅读README.en.md可以查看英文指导
    2.  ITGeeker 官方博客 [www.itgeeker.net](https://www.itgeeker.net)
    3.  Gitee开源项目地址 [https://gitee.com/itgeeker/itgeeker_remove_str_from_word](https://gitee.com/itgeeker/itgeeker_remove_str_from_word) 
    4.  Github开源项目地址 [https://github.com/alanljj/itgeeker_remove_str_from_word](https://github.com/alanljj/itgeeker_remove_str_from_word) 
    5.  GeekerCloud奇客罗方智能科技 [https://www.geekercloud.com](https://www.geekercloud.com)
