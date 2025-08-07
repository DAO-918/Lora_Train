from PIL import Image as PILImage
import os

def resize_image_square(folder_path):
    for root, _, files in os.walk(folder_path):  # 遍历所有子文件夹
        for filename in files:
            file_path = os.path.join(root, filename)
            if filename.lower().endswith('.png'):
                # 打开原始 png 图片
                image = PILImage.open(file_path)
                # 获取原始图片的宽度和高度
                width, height = image.size
                if width == 512 and height == 512:
                    continue
                # 目标尺寸
                target_width = 512
                target_height = 512
                # 计算缩放比例，保持原始图片的宽高比
                if width > height:
                    scale = target_width / width
                else:
                    scale = target_height / height
                # 计算缩放后的尺寸
                new_width = int(width * scale)
                new_height = int(height * scale)
                # 缩放原始图片
                image = image.resize((new_width, new_height), PILImage.Resampling.LANCZOS)
                # 计算需要填充的边距（水平和垂直方向），以保证图片居中
                x_margin = (target_width - new_width) // 2
                y_margin = (target_height - new_height) // 2
                # 创建一个新的 512*512 的透明背景图片
                new_image = PILImage.new('RGBA', (target_width, target_height), (0, 0, 0, 0))
                # 将缩放后的图片粘贴到新图片的居中位置
                new_image.paste(image, (x_margin, y_margin))
                # 保存处理后的图片，保存为新的文件 new_image.png
                new_image.save(file_path)
                
            
# 示例调用
folder_path = "E:\models"  # 替换为你的图片文件夹路径
resize_image_square(folder_path)