import websocket
import uuid
import json
import urllib.request
import urllib.parse
from io import BytesIO
from PIL import Image


class ComfyWebSocketClient:
    def __init__(self, server_address, client_id: str = None):
        """
        初始化类，设置服务器地址，并生成客户端ID。
        
        :param server_address: 服务器地址，格式如 "127.0.0.1:8188"
        """
        self.server_address = server_address
        self.client_id = client_id or str(uuid.uuid4())
        self.ws = None
        self.connect()

    def connect(self):
        """
        建立与服务器的WebSocket连接。
        """
        self.ws = websocket.WebSocket()
        self.ws.connect(f"ws://{self.server_address}/ws?clientId={self.client_id}")

    def queue_prompt(self, prompt):
        """
        向服务器发送提示请求，并返回包含提示ID等信息的响应。
        
        :param prompt: 提示信息的字典格式数据
        :return: 服务器响应解析后的字典，包含提示ID等关键信息
        """
        p = {"prompt": prompt, "client_id": self.client_id}
        data = json.dumps(p).encode('utf-8')
        req = urllib.request.Request(f"http://{self.server_address}/prompt", data=data)
        return json.loads(urllib.request.urlopen(req).read())

    def get_image(self, filename, subfolder, folder_type):
        """
        根据给定的文件名、子文件夹和文件夹类型，从服务器获取图像数据。
        
        :param filename: 图像文件名
        :param subfolder: 子文件夹名称
        :param folder_type: 文件夹类型
        :return: 图像数据（字节形式）
        """
        data = {"filename": filename, "subfolder": subfolder, "type": folder_type}
        url_values = urllib.parse.urlencode(data)
        with urllib.request.urlopen(f"http://{self.server_address}/view?{url_values}") as response:
            return response.read()

    def get_history(self, prompt_id):
        """
        根据提示ID获取历史记录信息。
        
        :param prompt_id: 提示对应的ID
        :return: 解析后的历史记录信息字典
        """
        with urllib.request.urlopen(f"http://{self.server_address}/history/{prompt_id}") as response:
            return json.loads(response.read())

    def get_images(self, prompt):
        """
        通过WebSocket交互以及后续的历史记录查询和图像获取操作，获取与提示相关的所有输出图像。
        
        :param prompt: 提示信息的字典格式数据
        :return: 以节点ID为键，对应输出图像数据列表为值的字典
        """
        prompt_id = self.queue_prompt(prompt)['prompt_id']
        output_images = {}
        while True:
            out = self.ws.recv()
            if isinstance(out, str):
                message = json.loads(out)
                if message['type'] == 'executing':
                    data = message['data']
                    if data['node'] is None and data['prompt_id'] == prompt_id:
                        break  # Execution is done
            else:
                continue  # previews are binary data
            
        history = self.get_history(prompt_id)[prompt_id]
        for node_id in history['outputs']:
            node_output = history['outputs'][node_id]
            images_output = []
            if 'images' in node_output:
                for image in node_output['images']:
                    image_data = self.get_image(image['filename'], image['subfolder'], image['type'])
                    images_output.append(image_data)
            output_images[node_id] = images_output
            
        return output_images

    def close(self):
        """
        关闭WebSocket连接。
        """
        if self.ws:
            self.ws.close()