import os
import sys


def get_base_path():
    if getattr(sys, 'frozen', False):
        # 如果是在 MacOS 打包后的 .app 中，获取 dist 目录
        if sys.platform == "darwin" and ".app" in sys.executable:
            # 找到 .app 目录的路径
            app_path = os.path.dirname(sys.executable)
            # 返回 dist 目录，即跳过三个目录
            return os.path.dirname(os.path.dirname(os.path.dirname(app_path)))
        # 对于 Windows 或 Linux，直接返回执行文件所在目录
        return os.path.dirname(sys.executable)
    else:
        # 开发环境下，直接返回当前文件所在目录
        current_path = os.path.abspath(__file__)
        # while current_path and not current_path.endswith("manifest_package_for_win_10"):
        while current_path and not current_path.endswith("manifest_for_win11"):
            current_path = os.path.dirname(current_path)
        return current_path


# 获取资源文件路径
def get_resource_path(relative_path):
    base_path = get_base_path()
    return os.path.join(base_path, relative_path)
