import os
import shutil

# 指定根目录（需要遍历的文件夹）
root_dir = r"E:\Research\IOSS清关_win10\IOSS清关_win10\xml-235-95136635\IM5280002556"

# 指定目标文件夹（提取结果存放的文件夹）
output_dir = os.path.join(root_dir, "zc415_collected")

# 确保目标文件夹存在
os.makedirs(output_dir, exist_ok=True)

# 遍历根目录及其所有子文件夹
for foldername, subfolders, filenames in os.walk(root_dir):
    for filename in filenames:
        if filename.lower().startswith("zc415"):  # 匹配前缀 zc415，不区分大小写
            source_path = os.path.join(foldername, filename)
            target_path = os.path.join(output_dir, filename)

            # 如果重名，可以加上所在子目录名避免覆盖
            if os.path.exists(target_path):
                base, ext = os.path.splitext(filename)
                target_path = os.path.join(output_dir, f"{base}_{os.path.basename(foldername)}{ext}")

            shutil.copy2(source_path, target_path)  # 保留元数据复制

print(f"所有前缀为 'zc415' 的文件已提取到: {output_dir}")
