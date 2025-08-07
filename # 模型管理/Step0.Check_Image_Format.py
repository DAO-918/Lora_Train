from PIL import Image
import os



def check_image_format(folder_path):
    """
    检查多层文件夹结构中是否有非 PNG 编码格式的图片。
    :param folder_path: 顶层文件夹路径
    :return: 返回非 PNG 图片的列表
    """
    non_png_images = []  # 保存非 PNG 图片的路径
    
    for root, _, files in os.walk(folder_path):  # 遍历所有子文件夹
        for filename in files:
            file_path = os.path.join(root, filename)
            if not filename.lower().endswith(('.png', '.jpg', '.jpeg', '.webp', '.bmp', '.gif')):
                continue  # 跳过非图片文件
            
            # 尝试打开文件，检查是否是图片
            try:
                with Image.open(file_path) as img:
                    detected_format = img.format  # 获取图片的实际编码格式
                    if detected_format != "PNG":  # 检测非 PNG 图片
                        print(f"Found non-PNG image: {file_path}, Format: {detected_format}")
                        non_png_images.append((file_path, detected_format))
            except Exception as e:
                # 忽略非图片文件或无法处理的文件
                print(f"Skipping file: {file_path}, Error: {e}")
                
    return non_png_images


def batch_process_images(folder_path):
    """
    批量处理文件夹中的图片，将所有图片转换为 PNG 并覆盖原路径。
    :param folder_path: 图片所在文件夹路径
    """
    for root, _, files in os.walk(folder_path):  # 遍历所有子文件夹
        for filename in files:
            file_path = os.path.join(root, filename)
            if not filename.lower().endswith(('.png', '.jpg', '.jpeg', '.webp', '.bmp', '.gif')):
                continue  # 跳过非图片文件
            # 尝试打开文件，检查是否是图片
            try:
                with Image.open(file_path) as img:
                    detected_format = img.format  # 获取图片的实际编码格式
                    if detected_format != "PNG":  # 检测非 PNG 图片
                        print(f"Found non-PNG image: {file_path}, Format: {detected_format}")
                        # 转换为 PNG 格式
                        
                        # 转换为 RGB 模式（部分图片可能是其他模式，如 "P" 或 "RGBA"）
                        if img.mode != "RGB":
                            img = img.convert("RGB")
                        
                        # 示例 1: 调整图片质量（适用于 JPEG 格式）
                        #if filename.lower().endswith(('.jpg', '.jpeg')):
                        #    output_path = os.path.splitext(file_path)[0] + "_processed.jpg"
                        #    img.save(output_path, "JPEG", quality=85)  # 调整压缩质量为 85
                        
                        # 示例 2: 保存为 PNG 格式并优化
                        png_path = os.path.splitext(file_path)[0] + ".png"  # 将文件扩展名替换为 .png
                        img.save(png_path, "PNG", optimize=True)
                        
                        # 示例 3: 转为灰度模式
                        # gray_img = img.convert("L")
                        # gray_output_path = os.path.splitext(file_path)[0] + "_gray.png"
                        # gray_img.save(gray_output_path, "PNG")
                        if not filename.lower().endswith('.png'):
                            os.remove(file_path)  # 删除原文件
            except Exception as e:
                # 忽略非图片文件或无法处理的文件
                print(f"Skipping file: {file_path}, Error: {e}")


# 示例调用
folder_path = "E:\models"  # 替换为你的图片文件夹路径
non_png_images = check_image_format(folder_path)

if non_png_images:
    print("\nNon-PNG images detected:")
    for image_path, image_format in non_png_images:
        print(f"File: {image_path}, Format: {image_format}")
else:
    print("All images in the folder (including subfolders) are PNG format.")

batch_process_images(folder_path)


