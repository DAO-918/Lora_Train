"""
python代码要求：
1. 代码文件的根目录下有256、512、768、896、1024、1152、1280、1408命名的文件夹(不存在的不用处理，不加入区间)，用已存在的文件夹作为总像素点区间(用最大值做分隔)，分别里面存放了目标是放大或缩小至该像素范围内的图片。仅处理**已存在的文件夹**，动态获取有效区间。  例如：若`896`文件夹不存在，代码会跳过该区间。区间从768训练集最大像素点到1024训练集最大像素点
2. 根目录下还有零散的图片，先要计算该图片的总像素点，根据已有的区间，决定缩小后应该存放在那一个文件夹中。
3. 图片处理第一步：进行裁剪，图片的尺寸的长宽能被64整除，向中心均匀裁剪
4. 图片处理第二部：如果图片尺寸过大或者过小，用最长边去计算最佳的放大缩小比例(按原尺寸比例)，使用Lanczos放大缩小图片
5. 图片处理第三步：如果图片的尺寸的长宽都不能被64整除，向中心均匀裁剪
6. 图片处理的每一步，都需要重新获取上一步处理后的图片尺寸。都需要进行一次判断，如果总像素在范围内，且长宽能被64整除，就不用执行下一步的图片处理。
7. 不需要考虑最小像素点不达标的情况
"""

import os
from PIL import Image
import sys
from typing import List, Tuple, Optional


def get_existing_folders(root_dir: str) -> List[str]:
    """
    获取根目录下所有以数字命名的现有文件夹
    
    Args:
        root_dir: 根目录路径
        
    Returns:
        按数值大小排序的文件夹名列表
    """
    print(f"正在扫描目录 {root_dir} 中的数字命名文件夹...")
    existing_folders = [f for f in os.listdir(root_dir) 
                        if os.path.isdir(os.path.join(root_dir, f)) and f.isdigit()]
    sorted_folders = sorted(existing_folders, key=int)  # 按数值大小排序
    if sorted_folders:
        print(f"找到以下数字命名文件夹: {', '.join(sorted_folders)}")
    else:
        print("未找到任何数字命名文件夹")
    return sorted_folders


def initialize_pixel_ranges(existing_folders: List[str]) -> Tuple[List[int], List[int]]:
    """
    初始化像素范围
    
    Args:
        existing_folders: 现有文件夹列表
        
    Returns:
        Ns: 文件夹名转为整数的列表
        max_pixels_list: 每个N对应的最大像素数列表
    """
    Ns = [int(f) for f in existing_folders]  # 文件夹名转为整数
    max_pixels_list = [N * N for N in Ns]    # 计算每个N对应的最大像素数
    print("初始化像素范围:")
    for i, N in enumerate(Ns):
        print(f"  - 文件夹 {N}: 最大像素数 = {max_pixels_list[i]} ({N}x{N})")
    return Ns, max_pixels_list

# 根据总像素数确定目标文件夹的函数
def get_target_N(total_pixels, Ns, max_pixels_list):
    """
    找到满足条件的最小N，使得total_pixels <= N*N。
    如果总像素数超过所有范围，则返回最大的N。
    
    Args:
        total_pixels: 图片总像素数
        Ns: 文件夹名转为整数的列表
        max_pixels_list: 每个N对应的最大像素数列表
        
    Returns:
        目标文件夹对应的N值
    """
    for i, max_pixels in enumerate(max_pixels_list):
        if total_pixels <= max_pixels:
            return Ns[i]
    return Ns[-1]  # 如果超过所有范围，返回最大的N

# 裁剪图片使尺寸为指定值的倍数的函数
def crop_to_multiple(img, multiple=64):
    """
    从中心裁剪图片，使宽度和高度都是'multiple'的倍数。
    返回裁剪后的图片及其新尺寸。
    
    Args:
        img: PIL图像对象
        multiple: 尺寸倍数，默认为64
        
    Returns:
        裁剪后的图片, 新宽度, 新高度
    """
    width, height = img.size
    w_crop = (width // multiple) * multiple  # 计算能被multiple整除的宽度
    h_crop = (height // multiple) * multiple  # 计算能被multiple整除的高度
    start_x = (width - w_crop) // 2  # 计算裁剪起始x坐标（居中）
    start_y = (height - h_crop) // 2  # 计算裁剪起始y坐标（居中）
    return img.crop((start_x, start_y, start_x + w_crop, start_y + h_crop)), w_crop, h_crop

def process_single_image(filepath: str, root_dir: str, Ns: List[int], max_pixels_list: List[int], 
                    delete_original: bool = True, multiple: int = 64) -> Optional[str]:
    """
    处理单张图片
    
    Args:
        filepath: 图片文件路径
        root_dir: 根目录路径
        Ns: 文件夹名转为整数的列表
        max_pixels_list: 每个N对应的最大像素数列表
        delete_original: 是否删除原始图片，默认为True
        multiple: 尺寸倍数，默认为64
        
    Returns:
        处理成功返回目标文件夹名，失败返回None
    """
    filename = os.path.basename(filepath)
    print(f"\n开始处理图片: {filename}")
    try:
        with Image.open(filepath) as img:
            # 步骤1：计算原始总像素数
            original_width, original_height = img.size
            total_pixels = original_width * original_height
            print(f"  步骤1: 原始图片尺寸 = {original_width}x{original_height}, 总像素数 = {total_pixels}")

            # 步骤2：确定目标文件夹
            target_N = get_target_N(total_pixels, Ns, max_pixels_list)
            target_max_pixels = target_N * target_N
            print(f"  步骤2: 确定目标文件夹 = {target_N}, 目标最大像素数 = {target_max_pixels}")

            # 步骤3：首先裁剪，使尺寸为64的倍数
            print(f"  步骤3: 裁剪图片使尺寸为{multiple}的倍数")
            cropped_img, w_crop, h_crop = crop_to_multiple(img, multiple)
            new_total_pixels = w_crop * h_crop
            print(f"    裁剪后尺寸 = {w_crop}x{h_crop}, 总像素数 = {new_total_pixels}")

            # 裁剪后检查
            if new_total_pixels <= target_max_pixels:
                # 如果在范围内且尺寸是64的倍数，使用裁剪后的图片
                final_img = cropped_img
                print(f"    裁剪后像素数在目标范围内，无需进一步处理")
            else:
                # 步骤4：如果总像素数超出目标范围，则调整大小
                print(f"  步骤4: 裁剪后像素数超出目标范围，需要调整大小")
                # 基于总像素数计算缩放比例，保持原始宽高比
                scale = (target_max_pixels / new_total_pixels) ** 0.5  # 计算缩放比例的平方根
                new_width = round(w_crop * scale)  # 计算新宽度
                new_height = round(h_crop * scale)  # 计算新高度
                print(f"    原始总像素数 = {new_total_pixels}, 目标最大像素数 = {target_max_pixels}")
                print(f"    缩放比例 = {scale:.4f}")
                print(f"    调整大小至 = {new_width}x{new_height}")
                resized_img = cropped_img.resize((new_width, new_height), Image.Resampling.LANCZOS)  # 使用Lanczos算法调整大小

                # 步骤5：调整大小后检查尺寸，如果需要再次裁剪
                if new_width % multiple == 0 and new_height % multiple == 0:
                    final_img = resized_img  # 如果尺寸已是64的倍数，直接使用
                    print(f"    调整大小后尺寸已是{multiple}的倍数，无需再次裁剪")
                else:
                    print(f"  步骤5: 调整大小后尺寸不是{multiple}的倍数，需要再次裁剪")
                    final_img, final_w, final_h = crop_to_multiple(resized_img, multiple)  # 否则再次裁剪
                    print(f"    最终尺寸 = {final_w}x{final_h}, 总像素数 = {final_w * final_h}")

            # 步骤6：将处理后的图片保存到目标文件夹
            print(f"  步骤6: 保存处理后的图片")
            target_folder = str(target_N)
            target_dir = os.path.join(root_dir, target_folder)
            if not os.path.exists(target_dir):
                os.makedirs(target_dir)
                print(f"    创建目标文件夹: {target_dir}")
            save_path = os.path.join(target_dir, filename)
            final_img.save(save_path)
            print(f"    图片已保存至: {save_path}")

        # 处理成功后删除源文件（如果需要）
        if delete_original:
            os.remove(filepath)
            print(f"已处理并移动 {filename} 到 {target_folder} 文件夹")
        else:
            print(f"已处理并复制 {filename} 到 {target_folder} 文件夹")
            
        return target_folder

    except Exception as e:
        print(f"处理图片 {filename} 时发生错误: {str(e)}")
        return None


def process_images(root_dir: str, delete_original: bool = True, multiple: int = 64) -> bool:
    """
    处理根目录下的所有图片
    
    Args:
        root_dir: 根目录路径
        delete_original: 是否删除原始图片，默认为True
        multiple: 尺寸倍数，默认为64
        
    Returns:
        处理成功返回True，失败返回False
    """
    print(f"\n开始处理目录 {root_dir} 下的图片...")
    print(f"参数设置: 删除原始图片 = {delete_original}, 尺寸倍数 = {multiple}")
    
    # 获取所有以数字命名的现有文件夹
    existing_folders = get_existing_folders(root_dir)
    
    # 如果没有找到有效文件夹，则退出程序
    if not existing_folders:
        print("未找到有效的数字命名文件夹，处理终止。")
        return False
    
    # 初始化像素范围
    Ns, max_pixels_list = initialize_pixel_ranges(existing_folders)
    
    # 处理根目录下的每张图片
    processed_count = 0
    error_count = 0
    image_count = 0
    
    # 统计需要处理的图片数量
    for filename in os.listdir(root_dir):
        if filename.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.tiff')):
            image_count += 1
    
    if image_count == 0:
        print("目录中没有找到需要处理的图片文件。")
        return True
    
    print(f"找到 {image_count} 张需要处理的图片。")
    current_image = 0
    
    for filename in os.listdir(root_dir):
        # 检查文件是否为图片
        if filename.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.tiff')):
            current_image += 1
            print(f"\n处理图片 {current_image}/{image_count}: {filename}")
            filepath = os.path.join(root_dir, filename)
            result = process_single_image(filepath, root_dir, Ns, max_pixels_list, delete_original, multiple)
            if result:
                processed_count += 1
            else:
                error_count += 1
    
    print(f"\n处理完成，总计: {image_count}，成功: {processed_count}，失败: {error_count}")
    return True


def main():
    """
    主函数入口
    """
    print("\n===== 图片尺寸处理程序开始 =====\n")
    
    # 默认根目录
    default_root_dir = r'E:\Design\Styles\口口AX1的插图・漫画 - pixiv\resize'
    
    # 如果有命令行参数，使用第一个参数作为根目录
    if len(sys.argv) > 1:
        root_dir = sys.argv[1]
        print(f"使用命令行参数指定的目录: {root_dir}")
    else:
        root_dir = default_root_dir
        print(f"使用默认目录: {root_dir}")
    
    # 处理图片
    result = process_images(root_dir)
    
    print("\n===== 图片尺寸处理程序结束 =====\n")
    return result


# 当脚本直接运行时执行main函数
if __name__ == "__main__":
    main()