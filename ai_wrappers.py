import base64
import json
import requests
from abc import ABC, abstractmethod

class AIWrapper(ABC):
    """AI 服务统一接口"""
    @abstractmethod
    def ocr_and_classify(self, img_bytes: bytes, filename: str):
        """返回 (识别出的文字, 证件类型)"""
        pass

    @abstractmethod
    def chat(self, messages, system_prompt=None, response_format=None, **kwargs):
        """发送对话请求，返回模型输出的字符串或解析后的 JSON"""
        pass


class ZhipuWrapper(AIWrapper):
    """智谱 AI 实现（保留原功能）"""
    def __init__(self, api_key):
        self.api_key = api_key
        from zhipuai import ZhipuAI
        self.client = ZhipuAI(api_key=api_key)

    def ocr_and_classify(self, img_bytes: bytes, filename: str):
        url = "https://open.bigmodel.cn/api/paas/v4/files/ocr"
        headers = {"Authorization": f"Bearer {self.api_key}"}
        files = {'file': (filename, img_bytes, 'image/jpeg')}
        data = {'tool_type': 'hand_write'}
        try:
            resp = requests.post(url, headers=headers, files=files, data=data, timeout=30)
            resp.raise_for_status()
            result = resp.json()
            words_list = result.get('words_result', [])
            text = "\n".join([item.get('words', '') for item in words_list])
        except Exception as e:
            return "", None

        # 关键词匹配判断证件类型
        img_type = None
        if text:
            if "居民身份证" in text and "签发机关" in text and "有效期限" in text:
                img_type = "身份证正面照片"
            elif "姓名" in text and "性别" in text and "公民身份号码" in text:
                img_type = "身份证反面照片"
            elif "教育部学历证书电子注册备案表" in text:
                img_type = "学信网学历证书电子备案截图"
            elif "中国高等教育学位在线验证报告" in text:
                img_type = "学信网学位证书电子备案截图"
            elif "毕业证书" in text:
                img_type = "毕业证照片"
            elif "学位证书" in text:
                img_type = "学位证照片"
        return text, img_type

    def chat(self, messages, system_prompt=None, response_format=None, **kwargs):
        full_messages = []
        if system_prompt:
            full_messages.append({"role": "system", "content": system_prompt})
        full_messages.extend(messages)
        params = {
            "model": kwargs.get("model", "glm-4-plus"),
            "messages": full_messages,
            "temperature": kwargs.get("temperature", 0.1),
        }
        if response_format == "json_object":
            params["response_format"] = {"type": "json_object"}
        resp = self.client.chat.completions.create(**params)
        content = resp.choices[0].message.content
        if response_format == "json_object":
            try:
                return json.loads(content)
            except:
                return content
        return content


class AgnesWrapper(AIWrapper):
    """Agnes AI 实现（兼容 OpenAI 格式）"""
    def __init__(self, api_key, base_url="https://apihub.agnes-ai.com/v1"):
        self.api_key = api_key
        self.base_url = base_url

    def _chat_completion(self, messages, model="agnes-2.0-flash", temperature=0.1):
        url = f"{self.base_url}/chat/completions"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        payload = {
            "model": model,
            "messages": messages,
            "temperature": temperature,
        }
        resp = requests.post(url, headers=headers, json=payload, timeout=60)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]

    def ocr_and_classify(self, img_bytes: bytes, filename: str):
        # 将图片转为 base64
        b64 = base64.b64encode(img_bytes).decode('utf-8')
        mime = "image/png" if filename.lower().endswith('.png') else "image/jpeg"
        image_url = f"data:{mime};base64,{b64}"

        prompt = (
            "请识别这张图片中的所有文字，并判断它属于以下哪一类证件（若无法判断则返回 null）："
            "身份证正面照片、身份证反面照片、毕业证照片、学位证照片、学信网学历证书电子备案截图、学信网学位证书电子备案截图。"
            "请以 JSON 格式返回：{\"text\": \"识别出的全部文字\", \"type\": \"证件类型或 null\"}"
        )
        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": image_url}}
                ]
            }
        ]
        try:
            # 使用图像模型（假设 agnes-image-2.1-flash 支持 chat 多模态）
            response_text = self._chat_completion(messages, model="agnes-image-2.1-flash", temperature=0.1)
            result = json.loads(response_text)
            return result.get("text", ""), result.get("type")
        except Exception as e:
            return "", None

    def chat(self, messages, system_prompt=None, response_format=None, **kwargs):
        full_messages = []
        if system_prompt:
            full_messages.append({"role": "system", "content": system_prompt})
        full_messages.extend(messages)
        model = kwargs.get("model", "agnes-2.0-flash")
        temperature = kwargs.get("temperature", 0.1)
        content = self._chat_completion(full_messages, model=model, temperature=temperature)
        if response_format == "json_object":
            try:
                return json.loads(content)
            except:
                return content
        return content
