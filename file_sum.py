import json
import os
import csv
import re
import requests
import plugins
from bridge.reply import Reply, ReplyType
from bridge.context import ContextType
from channel.chat_message import ChatMessage
from common.log import logger
from common.expired_dict import ExpiredDict
from docx import Document
from bs4 import BeautifulSoup
from pptx import Presentation
from openpyxl import load_workbook
import fitz  # PyMuPDF

# 文件类型映射
EXTENSION_TO_TYPE = {
    'pdf': 'pdf',
    'doc': 'docx', 'docx': 'docx',
    'md': 'md',
    'txt': 'txt',
    'xls': 'excel', 'xlsx': 'excel',
    'csv': 'csv',
    'html': 'html', 'htm': 'html',
    'ppt': 'ppt', 'pptx': 'ppt'
}

@plugins.register(
    name="filesum",
    desire_priority=2,
    desc="A plugin for summarizing files",
    version="1.0.0",
    author="sofs2005",
)

class filesum(Plugin):
    def __init__(self):
        super().__init__()
        try:
            curdir = os.path.dirname(__file__)
            config_path = os.path.join(curdir, "config.json")
            if os.path.exists(config_path):
                with open(config_path, "r", encoding="utf-8") as f:
                    self.config = json.load(f)
            else:
                self.config = super().load_config()
                if not self.config:
                    raise Exception("config.json not found")
                    
            self.handlers[Event.ON_HANDLE_CONTEXT] = self.on_handle_context
            
            # API配置
            self.api_config = {
                'open_ai_api_key': self.config.get("open_ai_api_key", ""),
                'open_ai_api_base': self.config.get("open_ai_api_base", "https://api.openai.com/v1"),
                'model': self.config.get("model", "gpt-3.5-turbo"),
                'max_token_size': self.config.get("max_token_size", 4000)
            }
            
            # 功能配置
            self.feature_config = {
                'enabled': self.config.get("enabled", False),
                'service': self.config.get("service", ""),
                'max_file_size': self.config.get("max_file_size", 15000),
                'group': self.config.get("group", True),
                'qa_prefix': self.config.get("qa_prefix", "问"),
                'prompt': self.config.get("prompt", ""),
                'file_cache_time': self.config.get("file_cache_time", 60),
                'content_cache_time': self.config.get("content_cache_time", 300)
            }

            # 将配置直接映射到类属性
            for key, value in {**self.api_config, **self.feature_config}.items():
                setattr(self, key, value)

            # 必要的配置检查
            if not self.open_ai_api_key:
                logger.error("[filesum] open_ai_api_key not configured")
                raise Exception("open_ai_api_key not configured")
                
            if not self.enabled:
                logger.info("[filesum] plugin is disabled")

            # 创建两个不同过期时间的缓存
            self.file_cache = ExpiredDict(self.file_cache_time)      # 文件路径缓存
            self.content_cache = ExpiredDict(self.content_cache_time) # 文件内容缓存

            logger.info("[filesum] inited.")
        except Exception as e:
            logger.warn(f"filesum init failed: {e}")

    def on_handle_context(self, e_context: EventContext):
        context = e_context["context"]
        msg: ChatMessage = e_context["context"]["msg"]
        
        # 获取会话ID，如果没有则使用默认值
        chat_id = context.get("session_id", "default")
        user_id = msg.from_user_id
        isgroup = e_context["context"].get("isgroup", False)
        
        # 生成缓存key
        cache_key = f"{chat_id}_{user_id}"

        if isgroup and not self.group:
            logger.info("[filesum] 群聊消息，文件处理功能已禁用")
            return

        # 处理文件消息
        if context.type == ContextType.FILE and self.enabled:
            logger.info("[filesum] 收到文件，存入缓存")
            context.get("msg").prepare()
            file_path = context.content
            
            # 使用组合key存储文件路径
            self.file_cache[cache_key] = {
                'file_path': file_path,
                'processed': False
            }
            logger.info(f"[filesum] 文件路径已缓存: {file_path}")

            # 如果是单聊，直接触发总结
            if not isgroup:
                logger.info("[filesum] 单聊消息，自动触发总结")
                return self._process_file_summary(cache_key, e_context)
            return

        # 处理文本消息
        if context.type == ContextType.TEXT and self.enabled:
            text = context.content
            
            # 处理总结请求（仅群聊需要手动触发）
            if "总结" in text and cache_key in self.file_cache and isgroup:
                return self._process_file_summary(cache_key, e_context)

            # 处理追问
            elif text.startswith(self.qa_prefix) and cache_key in self.content_cache:
                cache_data = self.content_cache.get(cache_key)
                if not cache_data:
                    logger.info("[filesum] 未找到缓存的文件内容")
                    reply = Reply(ReplyType.ERROR, "文件内容已过期，请重新发送文件")
                    e_context["reply"] = reply
                    return

                file_content = cache_data.get('file_content')
                if not file_content:
                    logger.info("[filesum] 缓存中没有文件内容")
                    reply = Reply(ReplyType.ERROR, "文件内容已过期，请重新发送文件")
                    e_context["reply"] = reply
                    return

                # 处理追问
                question = text[len(self.qa_prefix):].strip()
                self.handle_question(file_content, question, e_context)

    def _process_file_summary(self, cache_key: str, e_context: EventContext):
        """处理文件总结的核心逻辑"""
        cache_data = self.file_cache.get(cache_key)
        if not cache_data:
            logger.info("[filesum] 未找到缓存的文件")
            return
        
        if cache_data.get('processed', False):
            logger.info("[filesum] 该文件已经处理过")
            return

        file_path = cache_data.get('file_path')
        if not file_path or not os.path.exists(file_path):
            logger.info("[filesum] 缓存的文件不存在")
            reply = Reply(ReplyType.ERROR, "文件已过期，请重新发送")
            e_context["reply"] = reply
            return

        # 读取文件内容
        logger.info("[filesum] 开始读取文件内容")
        file_content = self.extract_content(file_path)
        if file_content is None:
            logger.info("[filesum] 文件内容无法提取")
            reply = Reply(ReplyType.ERROR, "无法读取文件内容")
            e_context["reply"] = reply
            return

        # 将文件内容存入内容缓存
        self.content_cache[cache_key] = {
            'file_content': file_content,
            'processed': True
        }
        
        # 处理文件内容
        self.handle_file(file_content, e_context)
        
        # 处理完成后删除文件
        try:
            os.remove(file_path)
            logger.info(f"[filesum] 文件 {file_path} 已删除")
            # 删除文件路径缓存
            del self.file_cache[cache_key]
        except Exception as e:
            logger.error(f"[filesum] 删除文件失败: {str(e)}")

    def extract_content(self, file_path):
        """提取文件内容"""
        try:
            # 添加文件大小检查
            file_size = os.path.getsize(file_path) / 1024  # 转换为KB
            if file_size > self.max_file_size:
                logger.error(f"文件大小 ({file_size}KB) 超过限制 ({self.max_file_size}KB)")
                return None
            
            file_extension = os.path.splitext(file_path)[1].lower().replace('.', '')
            file_type = EXTENSION_TO_TYPE.get(file_extension)
            
            if file_type == 'pdf':
                return self.read_pdf(file_path)
            elif file_type == 'docx':
                return self.read_docx(file_path)
            elif file_type == 'md':
                return self.read_markdown(file_path)
            elif file_type == 'txt':
                return self.read_txt(file_path)
            elif file_type == 'excel':
                return self.read_excel(file_path)
            elif file_type == 'csv':
                return self.read_csv(file_path)
            elif file_type == 'html':
                return self.read_html(file_path)
            elif file_type == 'ppt':
                return self.read_ppt(file_path)
            else:
                logger.error(f"不支持的文件类型: {file_extension}")
                return None
        except Exception as e:
            logger.error(f"提取文件内容时出错: {str(e)}")
            return None

    def read_pdf(self, file_path):
        """读取PDF文件"""
        try:
            doc = fitz.open(file_path)
            content = ' '.join([page.get_text() for page in doc])
            doc.close()
            return content
        except Exception as e:
            logger.error(f"读取PDF文件失败: {str(e)}")
            return None

    def read_docx(self, file_path):
        """读取Word文档"""
        try:
            doc = Document(file_path)
            content = '\n'.join([paragraph.text for paragraph in doc.paragraphs])
            return content
        except Exception as e:
            logger.error(f"读取Word文档失败: {str(e)}")
            return None

    def read_markdown(self, file_path):
        """读取Markdown文件"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            return remove_markdown(content)
        except Exception as e:
            logger.error(f"读取Markdown文件失败: {str(e)}")
            return None

    def read_txt(self, file_path):
        """读取文本文件"""
        encodings = ['utf-8', 'gbk', 'gb2312', 'ascii']
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    return f.read()
            except UnicodeDecodeError:
                continue
            except Exception as e:
                logger.error(f"读取文本文件失败: {str(e)}")
                return None
        return None

    def read_excel(self, file_path):
        """读取Excel文件"""
        try:
            wb = load_workbook(file_path)
            content = []
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    content.append('\t'.join([str(cell) if cell is not None else '' for cell in row]))
            return '\n'.join(content)
        except Exception as e:
            logger.error(f"读取Excel文件失败: {str(e)}")
            return None

    def read_csv(self, file_path):
        """读取CSV文件"""
        try:
            content = []
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                for row in reader:
                    content.append('\t'.join(row))
            return '\n'.join(content)
        except Exception as e:
            logger.error(f"读取CSV文件失败: {str(e)}")
            return None

    def read_html(self, file_path):
        """读取HTML文件"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                soup = BeautifulSoup(f.read(), 'html.parser')
                return soup.get_text()
        except Exception as e:
            logger.error(f"读取HTML文件失败: {str(e)}")
            return None

    def read_ppt(self, file_path):
        """读取PPT文件"""
        try:
            prs = Presentation(file_path)
            content = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        content.append(shape.text)
            return '\n'.join(content)
        except Exception as e:
            logger.error(f"读取PPT文件失败: {str(e)}")
            return None

    def handle_file(self, content, e_context):
        """处理文件内容"""
        try:
            if not content:
                reply = Reply(ReplyType.ERROR, "无法读取文件内容")
                e_context["reply"] = reply
                return

            # 使用配置中的token限制
            if len(content) > self.max_token_size:
                content = content[:self.max_token_size] + "..."
                logger.warning(f"文件内容已截断到 {self.max_token_size} 个字符")

            user_id = e_context["context"]["msg"].from_user_id
            prompt = self.prompt
            if user_id in self.params_cache and 'prompt' in self.params_cache[user_id]:
                prompt = self.params_cache[user_id]['prompt']

            # 构建提示词
            messages = [
                {"role": "system", "content": "你是一个文件总结助手。"},
                {"role": "user", "content": f"{prompt}\n\n{content}"}
            ]

            # 调用OpenAI API
            response = requests.post(
                f"{self.open_ai_api_base}/chat/completions",
                headers={
                    "Authorization": f"Bearer {self.open_ai_api_key}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": self.model,
                    "messages": messages
                }
            )

            if response.status_code == 200:
                result = response.json()
                summary = result['choices'][0]['message']['content']
                reply = Reply(ReplyType.TEXT, summary)
            else:
                reply = Reply(ReplyType.ERROR, "调用API失败")

            e_context["reply"] = reply

        except Exception as e:
            logger.error(f"处理文件内容时出错: {str(e)}")
            reply = Reply(ReplyType.ERROR, f"处理文件时出错: {str(e)}")
            e_context["reply"] = reply

    def handle_question(self, content, question, e_context):
        """处理追问"""
        try:
            # 构建提示词
            messages = [
                {"role": "system", "content": "你是一个文件问答助手。请基于给定的文件内容回答问题。"},
                {"role": "user", "content": f"文件内容如下：\n\n{content}\n\n问题：{question}"}
            ]

            # 调用OpenAI API
            response = requests.post(
                f"{self.open_ai_api_base}/chat/completions",
                headers={
                    "Authorization": f"Bearer {self.open_ai_api_key}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": self.model,
                    "messages": messages
                }
            )

            if response.status_code == 200:
                result = response.json()
                answer = result['choices'][0]['message']['content']
                reply = Reply(ReplyType.TEXT, answer)
            else:
                reply = Reply(ReplyType.ERROR, "调用API失败")

            e_context["reply"] = reply

        except Exception as e:
            logger.error(f"处理追问时出错: {str(e)}")
            reply = Reply(ReplyType.ERROR, f"处理追问时出错: {str(e)}")
            e_context["reply"] = reply

def remove_markdown(text):
    """移除Markdown格式"""
    # 移除标题
    text = re.sub(r'#{1,6}\s+', '', text)
    # 移除加粗和斜体
    text = re.sub(r'\*{1,2}(.*?)\*{1,2}', r'\1', text)
    # 移除链接
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)
    # 移除代码块
    text = re.sub(r'```[\s\S]*?```', '', text)
    # 移除行内代码
    text = re.sub(r'`([^`]+)`', r'\1', text)
    return text.strip()