import requests
import random
import json
import os
import yaml
from hashlib import md5

class BaiduTranslator:
    def __init__(self, appid=None, appkey=None):
        """
        初始化翻译器
        :param appid: 用户的appid，如果为None则从配置文件读取
        :param appkey: 用户的appkey，如果为None则从配置文件读取
        """
        # 如果未提供appid和appkey，则从配置文件读取
        if appid is None or appkey is None:
            config = self._load_config()
            self.appid = appid or config.get('baidu', {}).get('appid', '')
            self.appkey = appkey or config.get('baidu', {}).get('appkey', '')
        else:
            self.appid = appid
            self.appkey = appkey
            
        self.endpoint = 'http://api.fanyi.baidu.com'
        self.path = '/api/trans/vip/translate'
        self.url = self.endpoint + self.path
        self.headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        
    def _load_config(self):
        """
        从yaml配置文件加载翻译API配置
        :return: 配置字典
        """
        config_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                                 "yaml", "translate_config.yaml")
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    return yaml.safe_load(f) or {}
        except Exception as e:
            print(f"加载翻译配置文件失败: {e}")
        return {}

    def _make_md5(self, s, encoding='utf-8'):
        """
        生成MD5签名
        :param s: 原始字符串
        :param encoding: 编码方式
        :return: MD5签名
        """
        return md5(s.encode(encoding)).hexdigest()

    def translate(self, query, from_lang='zh', to_lang='en'):
        """
        翻译文本
        :param query: 需要翻译的文本
        :param from_lang: 源语言，默认中文
        :param to_lang: 目标语言，默认英文
        :return: 翻译结果或错误信息
        """
        salt = random.randint(32768, 65536)
        sign = self._make_md5(self.appid + query + str(salt) + self.appkey)
        payload = {
            'appid': self.appid,
            'q': query,
            'from': from_lang,
            'to': to_lang,
            'salt': salt,
            'sign': sign
        }

        try:
            response = requests.post(self.url, params=payload, headers=self.headers)
            result = response.json()
            if "error_code" in result:
                return f"Error {result['error_code']}: {result.get('error_msg', 'Unknown error')}"
            return result['trans_result'][0]['dst']
        except requests.exceptions.RequestException as e:
            return f"Request failed: {e}"

# 使用示例
if __name__ == "__main__":
    # 从配置文件加载API密钥
    translator = BaiduTranslator()
    text_to_translate = "你好"
    result = translator.translate(text_to_translate)
    print("翻译结果:", result)
