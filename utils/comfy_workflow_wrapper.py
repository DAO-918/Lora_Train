import json
import logging
from typing import Any, List

_log = logging.getLogger(__name__)

class ComfyWorkflowWrapper(dict):
    def __init__(self, path: str):
        """
        初始化ComfyWorkflowWrapper对象。
        此构造函数的主要目的是从给定的文件路径读取工作流相关的数据，并将其解析后存储为字典形式，以便后续对工作流进行各种操作。
        
        参数：
            path (str)：指向工作流文件的路径，通过该路径找到对应的文件来读取工作流信息。
        """
        # 默认的模式是r（只读模式）
        # 打开指定路径的文件，以文本模式读取文件内容，并将内容存储在workflow_str变量中，这里假设文件能正常打开读取
        with open(path, 'r', encoding='utf-8') as f:
            workflow_str = f.read()
        # 使用json.loads方法将读取到的字符串形式的工作流数据解析为Python的字典结构，
        # 然后调用父类（dict）的构造函数，将解析后的字典作为参数传入，完成对象的初始化，使其具备字典的属性和方法，便于后续操作
        super().__init__(json.loads(workflow_str))

    def list_nodes(self) -> List[str]:
        """
        获取工作流中所有节点的标题信息，并以列表形式返回。
        遍历工作流中存储的各个节点（通过super().values()获取所有节点的字典表示），提取每个节点中"_meta"字段下的"title"值，
        最终将这些标题值组成一个列表返回，方便了解工作流中包含哪些节点。
        
        返回值：
            List[str]：由工作流中各个节点的标题组成的列表。
        """
        return [node["_meta"]["title"] for node in super().values()]

    def set_node_param(self, title: str, param: str, value):
        """
        为指定标题的节点设置特定参数的值。
        需要注意的是，该方法会更改所有具有相同标题的节点对应的参数值，因为它是遍历所有节点进行匹配和设置的。
        
        参数：
            title (str)：要设置参数的节点的标题，用于在工作流的众多节点中定位目标节点。
            param (str)：需要设置的参数的名称，明确具体要修改哪个参数。
            value：要设置给参数的具体值，其类型可以是各种Python支持的数据类型，取决于参数本身的定义。
            
        异常抛出：
            ValueError：如果在遍历完所有节点后，都没有找到标题匹配的节点，则抛出此异常，表示要操作的节点不存在于当前工作流中。
        """
        smth_changed = False
        # 遍历工作流中的所有节点（通过super().values()获取节点字典），查找标题与传入title匹配的节点
        for node in super().values():
            if node["_meta"]["title"] == title:
                # 如果找到匹配的节点，使用日志记录工具记录当前正在为该节点设置参数的操作信息，方便后续查看操作记录和调试
                _log.info(f"Setting parameter '{title}' > '{param}' > '{value}'")
                # 将找到的节点中，对应参数名称（param）的参数值设置为传入的value值，完成参数设置操作
                node["inputs"][param] = value
                smth_changed = True
        # 如果遍历完所有节点后，都没有进行过参数设置操作（即没有找到匹配的节点），则抛出值错误异常
        if not smth_changed:
            raise ValueError(f"Node '{title}' not found.")

    def get_node_param(self, title: str, param: str) -> Any:
        """
        获取指定节点的指定参数的值。
        要注意的是，此方法只会返回首个找到的具有指定标题的节点的对应参数值，若工作流中有多个同名节点，不会返回所有同名节点的该参数值。
        
        参数：
            title (str)：目标节点的标题，用于定位节点。
            param (str)：要获取值的参数的名称，明确具体获取哪个参数的值。
            
        返回值：
            参数的值，其类型取决于参数本身在工作流中定义的数据类型，可以是任意Python支持的数据类型。
            
        异常抛出：
            ValueError：如果遍历完所有节点后，都没有找到标题匹配的节点，则抛出此异常，表示要获取参数值的节点不存在于当前工作流中。
        """
        for node in super().values():
            if node["_meta"]["title"] == title:
                # 当找到标题匹配的节点时，直接返回该节点中对应参数（param）的值，即从节点的"inputs"字段下获取对应参数值返回
                return node["inputs"][param]
        raise ValueError(f"Node '{title}' not found.")

    def get_node_id(self, title: str) -> str:
            """
            获取指定标题的节点的唯一标识符（ID）。
            
            参数：
                title (str)：目标节点的标题，通过该标题在工作流的所有节点中查找对应的节点。
                
            返回值：
                str：找到的节点的ID，用于唯一标识该节点，类型为字符串。
                
            异常抛出：
                ValueError：如果遍历完所有节点后，都没有找到标题匹配的节点，则抛出此异常，表示要获取ID的节点不存在于当前工作流中。
            """
            for id, node in super().items():
                if node["_meta"]["title"] == title:
                    return id
            raise ValueError(f"Node '{title}' not found.")

    def save_to_file(self, path: str):
        """
        将当前工作流对象以格式化后的JSON字符串形式保存到指定路径的文件中。
        
        参数：
            path (str)：要保存工作流文件的目标路径，确定文件保存的位置。
        """
        # 使用json.dumps方法将当前对象（本身继承自字典，存储着工作流相关信息）转换为格式化（缩进为4个空格）的JSON字符串，方便阅读和查看内容
        workflow_str = json.dumps(self, indent=4, ensure_ascii=False)
        # 打开指定路径的文件，以写入模式（如果文件不存在则创建，存在则覆盖）打开，将格式化后的JSON字符串写入文件中，完成工作流的保存操作
        with open(path, "w+", encoding='utf-8') as f:
            f.write(workflow_str)