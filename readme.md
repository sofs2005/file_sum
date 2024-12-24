# FileSum Plugin

## 简介
本项目是一个文件总结插件，可以配合[dify-on-wechat](https://github.com/hanfangyuan4396/dify-on-wechat)项目使用，支持多种文件格式的内容总结，理论上cow也支持，不过我没测试。发送文件后，通过"总结"命令触发，可以对文件内容进行智能总结，并支持后续追问。

## 功能特点
- 支持多种文件格式的内容总结：
  - PDF文件 (.pdf)
  - Word文档 (.doc, .docx)
  - Markdown文件 (.md)
  - 文本文件 (.txt)
  - Excel表格 (.xls, .xlsx)
  - CSV文件 (.csv)
  - HTML文件 (.html, .htm)
  - PPT文件 (.ppt, .pptx)
- 手动触发总结，避免无用调用
- 支持文件内容追问，加深理解
- 支持群聊和私聊场景
- 支持自定义提示词

## 安装
使用管理员口令在线安装，管理员认证方法见：[管理员认证](https://github.com/hanfangyuan4396/dify-on-wechat/tree/master/plugins/godcmd)
```bash
#installp https://github.com/sofs2005/file_sum.git
```
安装成功后，根据提示使用`#scanp` 命令来扫描新插件

## 配置
复制插件目录的`config.json.template`文件，重命名为`config.json`，配置参数即可。

配置文件参数说明：
```json
{
  "open_ai_api_key": "your-key",      # OpenAI API密钥
  "open_ai_api_base": "https://api.openai.com/v1",  # OpenAI API地址
  "model": "gpt-3.5-turbo",           # 使用的模型
  "enabled": true,                     # 是否启用插件
  "service": "openai",                 # 使用的服务（目前仅支持openai）
  "max_file_size": 15000,             # 支持的最大文件大小（KB）
  "max_token_size": 4000,             # API调用时的最大字符数
  "group": true,                       # 是否在群聊中启用
  "qa_prefix": "问",                   # 追问前缀
  "prompt": "请总结这个文件的主要内容", # 默认的总结提示词
  "file_cache_time": 60,              # 文件路径缓存时间（秒）
  "content_cache_time": 300           # 文件内容缓存时间（秒）
}
```

## 使用方法
1. 发送文件到聊天窗口
2. 处理方式：
   - 单聊：自动进行总结
   - 群聊：发送包含"总结"的消息触发总结功能
3. 在总结完成后的5分钟内，可以使用追问前缀（默认为"问"）进行追问
   例如：`问 这个文件的主要观点是什么？`

## 使用流程
```
用户: [发送文件]
用户: 总结一下
Bot: [文件总结内容]
用户: 问 文件中提到的第一个观点是什么？
Bot: [基于文件内容的回答]
```

## 注意事项
- 文件大小限制：默认最大支持15MB的文件
- 文件格式：请确保文件格式正确，且文件未加密
- API限制：根据OpenAI的token限制，过长的文件内容会被截断
- 缓存时间：
  - 发送文件后需要在1分钟内触发总结
  - 总结完成后可以在5分钟内进行追问
- 文件处理：总结完成后文件会被自动删除，请注意保管原文件

## 更新日志
- V1.0.0 (2024-12-24)
  - 初始版本发布
  - 支持多种文件格式的内容总结
  - 支持群聊手动触发总结功能
  - 支持单聊自动触发总结功能
  - 支持文件内容追问功能
  - 优化缓存机制

本插件基于fatwang的sum4all修改，感谢原作者。由于改动较大，我就重开了一个项目。
