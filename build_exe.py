"""
打包脚本，用于创建独立可执行的Windows应用程序
"""
import os
import sys
import shutil
import subprocess

print("开始执行打包过程...")

# 获取当前Python解释器路径
python_exe = sys.executable
print(f"使用Python解释器: {python_exe}")

# 安装必要的包 (如果尚未安装)
print("正在检查并安装必要的依赖...")
required_packages = ["pyinstaller", "pillow"]
for package in required_packages:
    print(f"确保已安装 {package}...")
    try:
        subprocess.check_call([python_exe, "-m", "pip", "install", package])
    except subprocess.CalledProcessError:
        print(f"警告: 安装 {package} 失败，可能会影响打包过程")

# 清理之前的构建文件夹（如果存在）
if os.path.exists("dist"):
    print("清理旧的dist文件夹...")
    shutil.rmtree("dist")
if os.path.exists("build"):
    print("清理旧的build文件夹...")
    shutil.rmtree("build")

# 使用PyInstaller进行打包 - 通过Python模块方式运行，而不是直接命令
print("开始打包应用程序...")

# 检查图标文件
icon_path = os.path.join("resources", "icon.ico")
if not os.path.exists(icon_path):
    print(f"警告: 图标文件 {icon_path} 不存在，将不使用自定义图标")
    icon_option = []
else:
    # 验证图标文件格式
    try:
        from PIL import Image
        try:
            # 尝试打开图标文件以验证格式
            test_icon = Image.open(icon_path)
            test_icon.verify()  # 验证图标格式
            print(f"图标文件验证成功: {icon_path}")
            icon_option = [f"--icon={icon_path}"]
        except Exception as e:
            print(f"图标文件无效 ({e})，将不使用自定义图标")
            icon_option = []
    except ImportError:
        print("警告: 未找到PIL模块，无法验证图标，但将继续尝试使用")
        icon_option = [f"--icon={icon_path}"]

# 构建PyInstaller命令
pyinstaller_cmd = [
    python_exe, 
    "-m", "PyInstaller",
    "--clean",
    "--noconfirm",
    "--windowed",
    "--onefile",
    "--name=PPT转视频工具", 
    "app.py"
]

# 添加图标选项(如果有效)
if icon_option:
    pyinstaller_cmd[6:6] = icon_option  # 在适当位置插入图标选项

# 执行打包命令
try:
    print(f"执行命令: {' '.join(pyinstaller_cmd)}")
    subprocess.check_call(pyinstaller_cmd)
    
    # 创建resources文件夹在dist目录中（如果不存在）
    dist_resources = os.path.join("dist", "resources")
    if not os.path.exists(dist_resources):
        os.makedirs(dist_resources)

    # 复制必要的资源文件
    if os.path.exists("XFYUN_TROUBLESHOOTING.md"):
        print("复制故障排除指南...")
        shutil.copy("XFYUN_TROUBLESHOOTING.md", "dist")

    print("\n===============================================")
    print("打包成功！生成的可执行文件位于dist文件夹")
    print("您可以分发dist文件夹中的'PPT转视频工具.exe'文件")
    print("===============================================")
except subprocess.CalledProcessError as e:
    print(f"\n打包失败: {e}")
    print("\n可能的解决方案:")
    print("1. 确保已安装PyInstaller和Pillow (pip install pyinstaller pillow)")
    print("2. 尝试以管理员身份运行此脚本")
    print("3. 尝试不使用图标进行打包 (移除--icon选项)")
    print("4. 手动运行PyInstaller命令进行调试:")
    print(f"   {' '.join(pyinstaller_cmd)}")
    sys.exit(1)
