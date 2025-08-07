import hashlib
import hmac
import json
import time
import os
import yaml
from datetime import datetime
from http.client import HTTPSConnection


class TencentTranslator:
    def __init__(self, secret_id=None, secret_key=None, source="zh", target="en", region=None):
        # 从配置文件加载API密钥和区域设置
        if secret_id is None or secret_key is None or region is None:
            config = self._load_config()
            self.secret_id = secret_id or config.get('tencent', {}).get('secret_id', '')
            self.secret_key = secret_key or config.get('tencent', {}).get('secret_key', '')
            self.region = region or config.get('tencent', {}).get('region', 'ap-guangzhou')
        else:
            self.secret_id = secret_id
            self.secret_key = secret_key
            self.region = region
            
        self.source = source
        self.target = target
        self.service = "tmt"
        self.host = f"{self.service}.tencentcloudapi.com"
        self.algorithm = "TC3-HMAC-SHA256"
        self.version = "2018-03-21"
        self.action = "TextTranslate"
        
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

    def sign(self, key, msg):
        return hmac.new(key, msg.encode("utf-8"), hashlib.sha256).digest()

    def build_canonical_request(self, payload):
        http_request_method = "POST"
        canonical_uri = "/"
        canonical_querystring = ""
        ct = "application/json; charset=utf-8"
        canonical_headers = f"content-type:{ct}\nhost:{self.host}\nx-tc-action:{self.action.lower()}\n"
        signed_headers = "content-type;host;x-tc-action"
        hashed_request_payload = hashlib.sha256(payload.encode("utf-8")).hexdigest()
        canonical_request = (
            f"{http_request_method}\n"
            f"{canonical_uri}\n"
            f"{canonical_querystring}\n"
            f"{canonical_headers}\n"
            f"{signed_headers}\n"
            f"{hashed_request_payload}"
        )
        return canonical_request

    def build_string_to_sign(self, canonical_request, timestamp, date):
        credential_scope = f"{date}/{self.service}/tc3_request"
        hashed_canonical_request = hashlib.sha256(canonical_request.encode("utf-8")).hexdigest()
        string_to_sign = (
            f"{self.algorithm}\n"
            f"{timestamp}\n"
            f"{credential_scope}\n"
            f"{hashed_canonical_request}"
        )
        return string_to_sign

    def calculate_signature(self, string_to_sign, date):
        secret_date = self.sign(("TC3" + self.secret_key).encode("utf-8"), date)
        secret_service = self.sign(secret_date, self.service)
        secret_signing = self.sign(secret_service, "tc3_request")
        signature = hmac.new(secret_signing, string_to_sign.encode("utf-8"), hashlib.sha256).hexdigest()
        return signature

    def build_authorization(self, signature, date):
        credential_scope = f"{date}/{self.service}/tc3_request"
        signed_headers = "content-type;host;x-tc-action"
        authorization = (
            f"{self.algorithm} "
            f"Credential={self.secret_id}/{credential_scope}, "
            f"SignedHeaders={signed_headers}, "
            f"Signature={signature}"
        )
        return authorization

    def translate(self, text):
        timestamp = int(time.time())
        date = datetime.utcfromtimestamp(timestamp).strftime("%Y-%m-%d")

        # 构造请求 payload
        payload = json.dumps({
            "SourceText": text,
            "Source": self.source,
            "Target": self.target,
            "ProjectId": 0  # 默认 ProjectId，可根据需求调整
        })

        # 构造请求
        canonical_request = self.build_canonical_request(payload)
        string_to_sign = self.build_string_to_sign(canonical_request, timestamp, date)
        signature = self.calculate_signature(string_to_sign, date)
        authorization = self.build_authorization(signature, date)

        headers = {
            "Authorization": authorization,
            "Content-Type": "application/json; charset=utf-8",
            "Host": self.host,
            "X-TC-Action": self.action,
            "X-TC-Timestamp": str(timestamp),
            "X-TC-Version": self.version,
            "X-TC-Region": self.region
        }

        try:
            req = HTTPSConnection(self.host)
            req.request("POST", "/", headers=headers, body=payload.encode("utf-8"))
            resp = req.getresponse()
            return resp.read().decode("utf-8")
        except Exception as err:
            return str(err)


# 使用示例
if __name__ == "__main__":
    # 从配置文件加载API密钥
    text_to_translate = "你好，世界！"  # 需要翻译的文本

    translator = TencentTranslator()
    response = translator.translate(text_to_translate)
    json_response = json.loads(response)
    print(json_response['Response'])
    print(json_response['Response']['TargetText'])
    js = json.loads(json_response['Response'])
    
    # 假设返回结果字典中，翻译后的文本对应的键是'target_text'，根据实际情况修改
    if isinstance(json_response, dict) and 'target_text' in json_response:
        translated_text = json_response['target_text']
        print(translated_text)
    else:
        print("返回结果格式不符合预期")
    print(response)
