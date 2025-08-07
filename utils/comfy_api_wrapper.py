import json
import requests
import uuid
import logging
import websockets
import asyncio
from requests.auth import HTTPBasicAuth
from requests.compat import urljoin, urlencode
from comfy_api_simplified.comfy_workflow_wrapper import ComfyWorkflowWrapper
import os

_log = logging.getLogger(__name__)


class ComfyApiWrapper:
    def __init__(
        self, url: str = "http://127.0.0.1:8188", user: str = "", password: str = ""
    ):
        """
        Initializes the ComfyApiWrapper object.
        
        Args:
            url (str): Comfy API 服务器的 URL，默认为 "http://127.0.0.1:8188"。
            user (str): 用于身份认证的用户名，默认为空字符串。
            password (str): 用于身份认证的密码，默认为空字符串。
            
        初始化步骤：
        - 保存服务器的 URL 和认证信息。
        - 根据 URL 决定使用的 WebSocket 协议（`ws://` 或 `wss://`）。
        - 如果提供用户名和密码，则将其加入 WebSocket URL 中用于认证。
        """
        # 保存服务器的 URL
        self.url = url
        self.auth = None  # 如果用户未提供用户名和密码，则不使用认证
        # 去掉 URL 的协议部分（http:// 或 https://），只保留主机名和端口
        url_without_protocol = url.split("//")[-1]
        # 根据 URL 是否包含 "https" 确定使用的 WebSocket 协议（wss 是加密的）
        if "https" in url:
            ws_protocol = "wss"  # 使用加密的 WebSocket 协议
        else:
            ws_protocol = "ws"  # 使用普通的 WebSocket 协议
            
        # 如果提供了用户名和密码，则设置 HTTP Basic Auth 并生成认证 WebSocket URL
        if user:
            # HTTPBasicAuth 用于在后续的 HTTP 请求中添加认证信息
            self.auth = HTTPBasicAuth(user, password)
            # 拼接 WebSocket URL，包含用户名和密码信息
            ws_url_base = f"{ws_protocol}://{user}:{password}@{url_without_protocol}"
        else:
            # 如果未提供认证信息，直接使用普通的 WebSocket URL
            ws_url_base = f"{ws_protocol}://{url_without_protocol}"
            
        # 拼接完整的 WebSocket URL，包含用于识别客户端的 `clientId` 参数
        self.ws_url = urljoin(ws_url_base, "/ws?clientId={}")
        

    def queue_prompt(self, prompt: dict, client_id: str | None = None) -> dict:
        """
        发送一个生成请求（prompt），并返回服务器的响应（通常包含 prompt_id）。
        
        Args:
            prompt (dict): 要发送到服务器的生成请求，通常是一个包含指令的字典。
            client_id (str): 用于标识客户端的 ID。默认为 None，如果不提供，会由服务器自行处理。
            
        Returns:
            dict: 服务器返回的响应，通常是一个包含 prompt_id 的 JSON 对象。
            
        Raises:
            Exception: 如果服务器响应的状态码不是 200（即非成功状态），抛出异常。
        """
        # 将生成请求 (prompt) 包装到一个字典中
        p = {"prompt": prompt}
        # 如果提供了 client_id，则将其添加到请求数据中
        if client_id:
            p["client_id"] = client_id
        # 将字典转换为 JSON 格式，并编码为字节数据以便发送
        data = json.dumps(p).encode("utf-8")
        # 记录请求发送日志（方便调试）
        _log.info(f"Posting prompt to {self.url}/prompt")
        # 使用 POST 方法向服务器发送请求
        #    - urljoin(self.url, "/prompt")：构造完整的请求 URL
        #    - data=data：请求体为 JSON 数据
        #    - auth=self.auth：如果设置了用户名和密码，添加认证信息
        resp = requests.post(urljoin(self.url, "/prompt"), data=data, auth=self.auth)
        # 记录服务器返回的状态码和原因
        _log.info(f"{resp.status_code}: {resp.reason}")
        # 如果请求成功（状态码为 200），将返回的数据转换为 JSON 并返回
        if resp.status_code == 200:
            return resp.json()
        # 如果请求失败（状态码不是 200），抛出异常，提示错误信息
        else:
            raise Exception(
                f"请求失败 with status code {resp.status_code}: {resp.reason}"
            )

    async def queue_prompt_and_wait(self, prompt: dict) -> str:
        """
        异步方法，发送生成任务 (prompt) 请求并等待其执行完成。
        
        功能：
            1. 使用 HTTP 提交 prompt 请求，并获取生成任务的 ID (prompt_id)。
            2. 通过 WebSocket 建立实时连接，监听任务执行状态。
            3. 根据接收到的服务器消息判断任务是否完成或发生错误。
            
        参数：
            prompt (dict): 要提交到服务器的生成任务请求。
            
        返回：
            str: 生成任务的唯一 ID (prompt_id)，用来标识任务。
            
        异常：
            Exception: 如果任务执行过程中出现错误，则抛出异常。
        """
        # 1. 生成一个唯一的客户端 ID（UUID，用于标识客户端连接）
        client_id = str(uuid.uuid4())
        # 2. 发送生成请求到服务器，并获取响应数据（包含 prompt_id）
        resp = self.queue_prompt(prompt, client_id)
        _log.debug(resp)  # 记录响应内容以便调试
        prompt_id = resp["prompt_id"]  # 提取生成任务的唯一 ID
        # 3. 记录即将建立 WebSocket 连接的 URL（屏蔽用户名和密码以保证安全性）
        _log.info(f"Connecting to {self.ws_url.format(client_id).split('@')[-1]}")
        # 4. 使用异步方式与服务器建立 WebSocket 连接
        async with websockets.connect(uri=self.ws_url.format(client_id)) as websocket:
            # 进入循环监听服务器返回的实时消息
            while True:
                # 5. 等待接收服务器发送的消息
                out = await websocket.recv()
                # 6. 如果收到的消息是字符串，将其解析为 JSON 格式
                if isinstance(out, str):
                    message = json.loads(out)
                    # 7. 过滤不需要的监控类型消息
                    if message["type"] == "crystools.monitor":
                        continue
                    _log.debug(message)  # 记录接收到的消息
                    # 8. 如果消息类型是 "execution_error"，检查是否与当前任务相关
                    if message["type"] == "execution_error":
                        data = message["data"]
                        if data["prompt_id"] == prompt_id:
                            raise Exception("Execution error occurred.")  # 抛出执行错误异常
                    # 9. 如果消息类型是 "status"，判断队列是否清空（任务是否完成）
                    if message["type"] == "status":
                        data = message["data"]
                        # 如果任务队列中剩余的任务数量为 0，则认为当前任务已完成
                        if data["status"]["exec_info"]["queue_remaining"] == 0:
                            return prompt_id
                    # 10. 如果消息类型是 "executing"，检查当前任务是否已执行完毕
                    if message["type"] == "executing":
                        data = message["data"]
                        # 如果任务节点为空并且任务 ID 匹配，则任务已完成
                        if data["node"] is None and data["prompt_id"] == prompt_id:
                            return prompt_id

    def queue_and_wait_images(
        self, prompt: ComfyWorkflowWrapper, output_node_title: str, loop: asyncio.BaseEventLoop = asyncio.get_event_loop()
    ) -> dict:
        """
        发送工作流任务（prompt），等待图片生成完成，并返回生成的图片内容。
        
        功能：
            1. 使用 ComfyWorkflowWrapper 封装任务，提交生成请求。
            2. 等待任务完成，获取生成历史记录。
            3. 提取指定输出节点的生成图片信息，下载图片内容。
            
        参数：
            prompt (ComfyWorkflowWrapper): 包含任务信息的 ComfyWorkflowWrapper 对象。
            output_node_title (str): 输出节点的标题，用于定位生成图片的节点。
            loop (asyncio.BaseEventLoop): 事件循环，用于运行异步任务，默认使用当前事件循环。
            
        返回：
            dict: 一个字典，将图片文件名映射到其内容（字节数据）。
            
        异常：
            Exception: 如果请求失败或任务执行失败，会抛出异常。
        """
        # 异步调用 queue_prompt_and_wait 方法，提交任务并等待任务完成
        prompt_id = loop.run_until_complete(self.queue_prompt_and_wait(prompt))
        # 获取任务执行的历史记录
        history = self.get_history(prompt_id)
        # 使用 prompt 提供的方法，获取指定输出节点的唯一 ID
        image_node_id = prompt.get_node_id(output_node_title)
        # 从历史记录中提取指定节点的生成图片信息
        images = history[prompt_id]["outputs"][image_node_id]["images"]
        # 遍历图片列表，下载每张图片的内容，并返回一个字典
        return {
            image["filename"]: self.get_image(
                image["filename"], image["subfolder"], image["type"]
            )
            for image in images
        }


    def get_queue(self) -> dict:
        """
        获取当前任务队列。
        
        功能：
            从服务器获取当前的任务队列状态，包括：
            - queue_running：当前正在执行的任务列表。
            - queue_pending：当前等待执行的任务列表。
            
        返回：
            dict: 包含任务队列状态的 JSON 对象。
            
        异常：
            Exception: 如果请求失败（非 200 状态码），抛出异常。
        """
        # 构造任务队列的请求 URL
        url = urljoin(self.url, f"/queue")
        _log.info(f"Getting queue from {url}")  # 记录获取队列的日志信息
        # 发送 GET 请求，获取任务队列数据
        resp = requests.get(url, auth=self.auth)
        # 检查 HTTP 响应状态码
        if resp.status_code == 200:
            # 如果状态码为 200，返回服务器响应的 JSON 数据
            return resp.json()
        else:
            # 如果状态码非 200，抛出异常并记录错误
            raise Exception(
                f"请求失败 with status code  {resp.status_code}: {resp.reason}"
            )


    def get_queue_size_before(self, prompt_id: str) -> int:
        """
        检索在某个prompt之前队列中prompt的数量。
        
        参数：
            prompt_id (str)：（prompt）的ID。
            
        返回值：
            int：在该prompt之前队列中prompt的数量，0表示该prompt正在运行。
            
        异常抛出：
            Exception：如果请求失败且状态码不是200。
            ValueError：如果prompt_id不在队列中。
        """
        # 调用self对象的get_queue方法获取队列相关信息，返回的结果赋值给resp变量
        resp = self.get_queue()
        
        # 遍历队列中正在运行的部分（假设resp["queue_running"] 存储的是正在运行的相关元素列表）
        for elem in resp["queue_running"]:
            # 如果找到元素中第二个值（索引为1的元素，可能是对应prompt的标识等）与传入的prompt_id相等
            if elem[1] == prompt_id:
                # 说明该prompt正在运行，返回0
                return 0
            
        # 初始化结果为1，表示如果在后续遍历等待队列时发现匹配的prompt，其前面至少有1个元素（此处只是初始化）
        result = 1
        # 遍历队列中等待的部分（假设resp["queue_pending"]存储的是处于等待状态的相关元素列表）
        for elem in resp["queue_pending"]:
            # 如果找到元素中第二个值（索引为1的元素）与传入的prompt_id相等
            if elem[1] == prompt_id:
                # 返回此时记录的前面元素的个数，即result的值
                return result
            # 如果当前元素不是要找的prompt_id对应的元素，则将前面元素个数加1，继续遍历下一个元素
            result = result + 1
        # 如果遍历完等待队列都没找到对应的prompt_id，则抛出值错误异常，
        raise ValueError("prompt_id不在队列中")

    def get_history(self, prompt_id: str) -> dict:
        """
        获取一个提示（prompt）的执行历史记录。
        
        参数：
            prompt_id (str)：提示（prompt）的ID。
            
        返回值：
            dict：响应的JSON对象。
            
        异常抛出：
            Exception：如果请求失败且状态码不是200。
        """
        # 使用urljoin函数将self.url和指定的路径（根据prompt_id拼接的历史记录相关路径）拼接起来，得到完整的请求URL
        url = urljoin(self.url, f"/history/{prompt_id}")
        # 使用日志记录工具记录当前正在从哪个URL获取历史记录信息，方便后续查看和调试
        _log.info(f"从 {url} 获取历史记录")
        # 使用requests库的get方法发送HTTP GET请求到指定的URL，同时传入认证信息（self.auth），并将响应结果赋值给resp变量
        resp = requests.get(url, auth=self.auth)
        # 判断响应的状态码是否为200，如果是，表示请求成功
        if resp.status_code == 200:
            # 将响应内容（JSON格式的字符串）解析为Python的字典对象并返回
            return resp.json()
        else:
            # 如果状态码不是200，说明请求失败，抛出异常，并在异常信息中包含状态码以及对应的原因，方便排查问题
            raise Exception(
                f"请求失败 with status code {resp.status_code}：{resp.reason}"
            )

    def get_image(self, filename: str, subfolder: str, folder_type: str) -> bytes:
        """
        从Comfy API服务器获取一张图片。
        根据文件名、子文件夹和文件类型，从服务器中获取生成的图片内容（以字节形式返回）。
        参数：
            filename (str)：图片的文件名。
            subfolder (str)：图片所在的子文件夹。
            folder_type (str)：文件夹的类型。
        返回值：
            bytes：图片的内容（以字节形式）。
        异常抛出：
            Exception：如果请求失败且状态码不是200。
        """
        # 将文件名、子文件夹和文件夹类型这些参数整理到一个字典中，方便后续构造请求URL时使用
        params = {"filename": filename, "subfolder": subfolder, "type": folder_type}
        # 使用urljoin函数把self.url和构造好的查询参数（通过urlencode将字典形式的参数编码成URL查询字符串格式）拼接起来，形成完整的请求URL
        url = urljoin(self.url, f"/view?{urlencode(params)}")
        # 使用日志记录工具记录当前正在从哪个URL获取图片，方便后续查看和调试操作，知晓图片获取的来源
        _log.info(f"从 {url} 获取图片")
        # 使用requests库的get方法发送HTTP GET请求到刚刚构造好的URL上，同时传入认证信息（self.auth），并将服务器返回的响应对象赋值给resp变量
        resp = requests.get(url, auth=self.auth)
        # 使用日志记录工具以调试级别记录响应的状态码以及对应的原因，方便在调试时详细查看请求的响应情况，排查可能出现的问题
        _log.debug(f"{resp.status_code}: {resp.reason}")
        # 判断响应的状态码是否为200，如果是，则表示图片获取请求成功
        if resp.status_code == 200:
            # 将响应内容（也就是图片的二进制数据内容）直接返回，因为响应的content属性存储的就是字节形式的内容，符合函数返回值要求
            return resp.content
        else:
            # 如果状态码不是200，说明图片获取请求失败，此时抛出异常，并在异常信息中包含具体的状态码以及对应的原因，方便后续定位和解决请求失败的问题
            raise Exception(
                f"请求失败 with status code {resp.status_code}：{resp.reason}"
            )
    
    def upload_image(
            self, filename: str, subfolder: str = "default_upload_folder"
        ) -> dict:
        """
        向Comfy API服务器上传一张图片。
        
        参数：
            filename (str)：图片的文件名。
            subfolder (str)：要将图片上传到的子文件夹。默认值为"default_upload_folder"。
            
        返回值：
            dict：响应的JSON对象。
            
        异常抛出：
            Exception：如果请求失败且状态码不是200。
        """
        # 使用urljoin函数将self.url和"/upload/image"路径拼接起来，得到用于上传图片的完整请求URL
        url = urljoin(self.url, "/upload/image")
        # 获取文件名的基本名称（去除路径部分，只保留文件名本身），通常用于在服务器端标识上传的文件，赋值给serv_file变量
        serv_file = os.path.basename(filename)
        # 创建一个字典，包含要上传到服务器的子文件夹信息，后续会作为请求的数据部分发送给服务器，这里键名为"subfolder"
        data = {"subfolder": subfolder}
        # 创建一个字典，用于指定要上传的文件相关信息。键名为"image"，对应的值是一个元组，元组中第一个元素是处理后的文件名（serv_file），
        # 第二个元素是通过以二进制只读模式（"rb"）打开的文件对象，这样requests库就能识别并读取该文件内容用于上传
        files = {"image": (serv_file, open(filename, "rb"))}
        # 使用日志记录工具记录当前正在将哪个文件（通过文件名标识）上传到哪个URL，以及附带的数据信息（这里的data），方便后续查看和调试上传操作情况
        _log.info(f"正在将 {filename} 发送到 {url}，附带数据 {data}")
        # 使用requests库的post方法发送HTTP POST请求到指定的URL，同时传入要上传的文件信息（files参数）、请求的数据（data参数）以及认证信息（self.auth），
        # 并将服务器返回的响应对象赋值给resp变量
        resp = requests.post(url, files=files, data=data, auth=self.auth)
        # 使用日志记录工具以调试级别记录响应的状态码、对应的原因以及响应的文本内容（resp.text，可能包含服务器返回的一些详细信息等），
        # 方便在调试时详细查看请求的响应情况，排查可能出现的上传问题
        _log.debug(f"{resp.status_code}: {resp.reason}, {resp.text}")
        # 判断响应的状态码是否为200，如果是，则表示图片上传请求成功
        if resp.status_code == 200:
            # 将响应内容（假设是JSON格式的字符串）解析为Python的字典对象并返回，这样调用者可以根据返回的字典获取服务器端关于上传操作的相关反馈信息
            return resp.json()
        else:
            # 如果状态码不是200，说明图片上传请求失败，此时抛出异常，并在异常信息中包含具体的状态码以及对应的原因，方便后续定位和解决上传请求失败的问题
            raise Exception(
                f"请求失败，状态码为 {resp.status_code}：{resp.reason}"
            )
