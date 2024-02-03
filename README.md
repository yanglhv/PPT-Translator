# PPT-Translator
## 简介

一个将PPT用LLM能力实现的小翻译工具

本项目有前置要求:

1. 调用LLM的翻译能力
2. 调用PPT的COM接口获取文本信息
3. 输出结构化, 翻译后的文本要返回到PPT中, 进行替换

## 安装

```
$ pip install -r requirements.txt
建议使用venv或conda起一个新环境进行安装
```

## 使用

```
$ python winMain.py
1. 执行命令后会弹出主窗口, 先填入LLM的API配置信息(当前默认给了一个配置,但token有限,仅可用于demo)
2. 填完信息, 点击保存, 下次会自动填充本次提交的信息;
3. 单选要翻译为的语言;
4. 将PPT拖拽到下方执行框, 翻译子程序即开始运行(因为是调用LLM能力,输出会较慢);
5. 执行完成后会弹框, 显示翻译后的文件路径(原路径就地保存了)
```

API 的配置信息在讯飞官网获取: https://xinghuo.xfyun.cn/spark

## 实现逻辑

1. 翻译部分: 通过调用大模型(目前使用讯飞星火,在翻译方面在国内LLM做的较好, 官方提供了web接口, 需要鉴权等操作比较繁琐, 实现可参考SparkAPI项目), 给以调教好特定的提示词(比较费功夫), 让其将Text文本翻译为指定的文本; 

2. PPT的操作部分, 直接用pptx库, 逐一取出段落文本, 输入到上一步的翻译类, 得到输出结果, 回填到段落, 另存为新文件;

   (PS: pptx库中需要仔细理解paragraphs, 成体系的资料不多, 打断点调试看变量能更快了解原理)

3. 主界面部分:  直接使用Tkinter来做一个简单的界面, 将操作和通知简单串接起来



## 界面效果

![](https://github.com/yanglhv/PPT-Translator/blob/main/translator_showdemo.gif)