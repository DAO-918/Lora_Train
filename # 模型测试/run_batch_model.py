import os
import argparse
import logging
import sys
from batch_model_test import BatchModelTester

# 配置日志
logging.basicConfig(stream=sys.stdout, level=logging.INFO)
logger = logging.getLogger(__name__)

def main():
    # 解析命令行参数
    parser = argparse.ArgumentParser(description="批量测试模型")
    parser.add_argument("--model-info", type=str, required=True, 
                        help="模型信息表格路径，包含模型的基本信息")
    parser.add_argument("--test-file", type=str, required=True,
                        help="测试文件路径，Excel格式，包含测试参数和配置")
    parser.add_argument("--comfy-host", type=str, default="127.0.0.1:8188", 
                        help="ComfyUI服务器地址，默认为127.0.0.1:8188")
    parser.add_argument("--insert-image", action="store_true", 
                        help="是否将生成的图片插入到Excel表格中")
    
    args = parser.parse_args()
    
    # 检查文件是否存在
    if not os.path.exists(args.model_info):
        logger.error(f"模型信息文件不存在: {args.model_info}")
        return
    
    if not os.path.exists(args.test_file):
        logger.error(f"测试文件不存在: {args.test_file}")
        return
    
    try:
        # 创建批量测试模型对象
        logger.info(f"初始化批量测试模型，模型信息: {args.model_info}, ComfyUI服务器: {args.comfy_host}")
        tester = BatchModelTester(args.model_info, args.comfy_host)
        
        # 加载模型信息
        logger.info("加载模型信息...")
        tester.load_model_info()
        
        # 处理测试文件
        logger.info(f"开始处理测试文件: {args.test_file}")
        tester.process_test_file(args.test_file)
        
        # 插入图片到Excel表格中
        if args.insert_image:
            logger.info("开始将图片插入到Excel表格中...")
            tester.reinsert_image(args.test_file)
            
        logger.info("批量测试完成！")
    except Exception as e:
        logger.error(f"执行过程中发生错误: {e}")

# 使用示例
if __name__ == "__main__":
    main()
    
'''
使用示例:

1. 基本用法:
   python run_batch_test.py --model-info "E:\models\model_info.xlsx" --test-file "D:\Code\MY_ComfyUI\# 模型测试\测试文件.xlsx"

2. 指定ComfyUI服务器地址:
   python run_batch_test.py --model-info "E:\models\model_info.xlsx" --test-file "D:\Code\MY_ComfyUI\# 模型测试\测试文件.xlsx" --comfy-host "127.0.0.1:8191"

3. 生成图片后插入到Excel:
   python run_batch_test.py --model-info "E:\models\model_info.xlsx" --test-file "D:\Code\MY_ComfyUI\# 模型测试\测试文件.xlsx" --insert-image
'''