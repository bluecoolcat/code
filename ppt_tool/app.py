import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
import threading
import os
import sys
import io
import datetime  # 添加datetime导入
from ppt_to_video_converter import ppt_to_video
import ssl
import traceback  # 确保导入了traceback模块
import re  # 添加re模块

class TextRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = ""
        
    def write(self, string):
        self.buffer += string
        self.text_widget.after(0, self.update_text_widget)
        
    def update_text_widget(self):
        self.text_widget.configure(state=tk.NORMAL)
        self.text_widget.insert(tk.END, self.buffer)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state=tk.DISABLED)
        self.buffer = ""
        
    def flush(self):
        pass

class PPTToVideoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("微源PPT转视频工具v1.3")
        self.root.geometry("800x900")  # 减小窗口的默认高度
        self.root.resizable(True, True)
        
        # 科大讯飞参数
        self.xfyun_app_id = tk.StringVar(value="b6537c6a")
        self.xfyun_api_key = tk.StringVar(value="2ced9e083736d0b00cd78bd2be7e5c85")
        self.xfyun_api_secret = tk.StringVar(value="NzlhY2EzYTkyYjlmOGY0ODFkZGQ5OGU0")
        self.xfyun_voice = tk.StringVar(value="x4_yezi")
        self.xfyun_speed = tk.IntVar(value=54)  # 添加语速变量，默认56
        
        # 马克配音参数
        self.ttsmaker_token = tk.StringVar(value="ttsmaker_demo_token")
        self.ttsmaker_voice_id = tk.StringVar(value="1504")
        
        # 添加精准字幕变量
        self.precise_subtitle = tk.BooleanVar(value=True)
        
        # 存储选中的PPT文件列表
        self.ppt_files = []
        
        # 添加文件状态字典 - 用于跟踪每个文件的转换状态
        self.file_statuses = {}  # 格式: {文件索引: "success"/"failed"}
        
        # 添加水印相关变量
        self.use_watermark = tk.BooleanVar(value=False)
        self.watermark_opacity = tk.IntVar(value=100)
        self._full_watermark_path = None  # 存储完整的水印图片路径
        
        # 添加文件计数器变量（用于目录递归扫描进度）
        self.file_scan_count = 0
        
        self.setup_ui()
    
    def setup_ui(self):
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File selection
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        # 创建左右双栏布局 - 左边放标签和按钮，右边放文件列表
        left_panel = ttk.Frame(file_frame)
        left_panel.grid(row=0, column=0, padx=5, pady=5, sticky=tk.N+tk.W)
        
        right_panel = ttk.Frame(file_frame)
        right_panel.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        file_frame.columnconfigure(1, weight=1)  # 让右侧面板可以扩展
        
        # 左侧面板 - 放置标签和按钮
        #ttk.Label(left_panel, text="PPT文件:").pack(anchor=tk.W, pady=0)
        
        buttons_frame = ttk.Frame(left_panel)
        buttons_frame.pack(fill=tk.X)
        
        ttk.Button(left_panel, text="添加文件", command=self.browse_ppt_files, width=15).pack(pady=0)
        ttk.Button(left_panel, text="添加文件夹", command=self.browse_ppt_folder, width=15).pack(pady=0)
        ttk.Button(left_panel, text="删除选中", command=self.delete_selected_files, width=15).pack(pady=0)
        
        # 右侧面板 - 放置文件列表
        files_list_frame = ttk.Frame(right_panel)
        files_list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建滚动条
        scrollbar = ttk.Scrollbar(files_list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 替换原来的Listbox为自定义文件列表 - 使用Text组件以支持颜色显示
        self.files_text = tk.Text(files_list_frame, height=6, width=60, 
                                  yscrollcommand=scrollbar.set,
                                  wrap=tk.NONE, cursor="arrow")
        self.files_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.files_text.yview)
        
        # 配置文本标签用于显示不同状态
        self.files_text.tag_configure("default", foreground="black")
        self.files_text.tag_configure("success", foreground="green")
        self.files_text.tag_configure("failed", foreground="red")
        
        # 绑定鼠标点击事件处理文件选择
        self.files_text.bind("<ButtonRelease-1>", self.on_file_click)
        self.files_text.configure(state=tk.DISABLED)  # 设为只读模式
        
        # 存储已选择的行号
        self.selected_indices = []
        
        # Options
        options_frame = ttk.LabelFrame(main_frame, text="选项", padding="10")
        options_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(options_frame, text="语音引擎:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.tts_engine_var = tk.StringVar(value="xfyun")  # 默认改为科大讯飞TTS
        ttk.Radiobutton(options_frame, text="系统TTS (pyttsx3)", variable=self.tts_engine_var, 
                       value="pyttsx3", command=self.update_tts_options).grid(row=0, column=1, sticky=tk.W, pady=5)
        ttk.Radiobutton(options_frame, text="科大讯飞TTS (需API密钥)", variable=self.tts_engine_var, 
                       value="xfyun", command=self.update_tts_options).grid(row=0, column=2, sticky=tk.W, pady=5)
        ttk.Radiobutton(options_frame, text="马克配音TTS", variable=self.tts_engine_var, 
                       value="ttsmaker", command=self.update_tts_options).grid(row=0, column=3, sticky=tk.W, pady=5)
        
        # 科大讯飞 设置框架
        self.xfyun_frame = ttk.LabelFrame(options_frame, text="科大讯飞设置", padding="5")
        self.xfyun_frame.grid(row=1, column=0, columnspan=4, sticky=tk.W+tk.E, pady=5)
        
        ttk.Label(self.xfyun_frame, text="APPID:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.xfyun_frame, textvariable=self.xfyun_app_id, width=40).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(self.xfyun_frame, text="API Key:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.xfyun_frame, textvariable=self.xfyun_api_key, width=40).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(self.xfyun_frame, text="API Secret:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.xfyun_frame, textvariable=self.xfyun_api_secret, width=40, show="*").grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # 添加测试按钮来验证API信息
        ttk.Button(self.xfyun_frame, text="测试API连接", command=self.test_xfyun_api).grid(row=2, column=2, padx=5, pady=5)
        
        ttk.Label(self.xfyun_frame, text="发音人:").grid(row=3, column=0, sticky=tk.W, pady=5)
        voice_combo = ttk.Combobox(self.xfyun_frame, textvariable=self.xfyun_voice, width=15)
        voice_combo['values'] = ( 'x4_yezi', 'x4_xiaoyuan', 'x4_lingxiaolu_en', 'aisbabyxu')
        voice_combo.grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # 添加语速调节滑块
        ttk.Label(self.xfyun_frame, text="语速:").grid(row=5, column=0, sticky=tk.W, pady=5)
        speed_scale = ttk.Scale(self.xfyun_frame, from_=0, to=100, 
                               variable=self.xfyun_speed, orient=tk.HORIZONTAL, length=200)
        speed_scale.grid(row=5, column=1, sticky=tk.W, pady=5)
        
        # 添加显示当前数值的标签
        self.speed_label = ttk.Label(self.xfyun_frame, text="54")
        self.speed_label.grid(row=5, column=2, sticky=tk.W, pady=5)
        
        # 更新滑块数值显示的函数
        def update_speed_label(*args):
            self.speed_label.config(text=str(self.xfyun_speed.get()))
        
        # 绑定滑块值变化事件
        self.xfyun_speed.trace_add("write", update_speed_label)
        
        # 添加关于科大讯飞发音人的说明
        ttk.Label(self.xfyun_frame, text="发音人说明: x4_yezi(讯飞小露)、x4_xiaoyuan(讯飞小媛)、x4_lingxiaolu_en(聆小璐)、aisbabyxu(童声)",
                 foreground="blue").grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # 马克配音TTS设置框架
        self.ttsmaker_frame = ttk.LabelFrame(options_frame, text="马克配音TTS设置", padding="5")
        self.ttsmaker_frame.grid(row=3, column=0, columnspan=4, sticky=tk.W+tk.E, pady=5)
        
        ttk.Label(self.ttsmaker_frame, text="Token:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.ttsmaker_frame, textvariable=self.ttsmaker_token, width=40).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(self.ttsmaker_frame, text="语音ID:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.ttsmaker_frame, textvariable=self.ttsmaker_voice_id, width=10).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # 添加测试按钮来验证API
        ttk.Button(self.ttsmaker_frame, text="测试API连接", command=self.test_ttsmaker_api).grid(row=1, column=2, padx=5, pady=5)
        
        # 添加语速输入框 (替换滑块)
        ttk.Label(self.ttsmaker_frame, text="语速:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.ttsmaker_speed = tk.DoubleVar(value=1.0)  # 默认语速为1.0

        # 创建验证函数，限制输入值在0.5-2.0之间，且第二位小数只能是0或5
        def validate_speed(action, value_if_allowed):
            if action == '1':  # 插入操作
                if value_if_allowed == "" or value_if_allowed == ".":
                    return True
                try:
                    # 检查是否超过允许的最大长度(如1.95为4个字符)
                    if len(value_if_allowed) > 4:
                        return False
                    
                    # 如果包含小数点且小数部分长度为2，检查第二位是否为0或5
                    if '.' in value_if_allowed:
                        parts = value_if_allowed.split('.')
                        # 如果小数部分长度为2，检查第二位是否为0或5
                        if len(parts[1]) == 2 and parts[1][1] not in ('0', '5'):
                            return False
                    
                    # 检查数值范围
                    val = float(value_if_allowed)
                    return 0.5 <= val <= 2.0
                except ValueError:
                    return False
            return True
            
        validate_cmd = self.root.register(validate_speed)
        speed_entry = ttk.Entry(self.ttsmaker_frame, textvariable=self.ttsmaker_speed, 
                                width=6, validate="key", 
                                validatecommand=(validate_cmd, '%d', '%P'))
        speed_entry.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # 添加格式化函数，确保值符合X.X0或X.X5格式
        def format_speed_value(*args):
            try:
                current_value = self.ttsmaker_speed.get()
                
                # 计算最接近的合法值(增量为0.05)
                # 先将值乘以100，四舍五入到整数，再除以100
                rounded_value = round(current_value * 20) / 20
                
                # 确保值在范围内
                if rounded_value < 0.5:
                    rounded_value = 0.5
                elif rounded_value > 2.0:
                    rounded_value = 2.0
                
                # 只有当值改变时才设置，避免无限循环
                if abs(current_value - rounded_value) > 0.001:
                    self.ttsmaker_speed.set(rounded_value)
                    
                # 始终格式化显示，确保有两位小数
                formatted_text = f"{rounded_value:.2f}"
                speed_entry.delete(0, tk.END)
                speed_entry.insert(0, formatted_text)
                
            except (ValueError, tk.TclError):
                # 恢复默认值
                self.ttsmaker_speed.set(1.0)
                speed_entry.delete(0, tk.END)
                speed_entry.insert(0, "1.00")
                
        # 绑定格式化函数到失去焦点和回车事件
        speed_entry.bind("<FocusOut>", format_speed_value)
        speed_entry.bind("<Return>", format_speed_value)
        
        # 添加范围说明标签 (更新说明文本)
        ttk.Label(self.ttsmaker_frame, text="（取值范围：0.50-2.00，增量0.05）").grid(row=2, column=2, sticky=tk.W, pady=5)
        
        # 马克配音说明
        ttk.Label(self.ttsmaker_frame, text="说明: 请输入正确的语音ID，默认1504为潇潇。",
                 foreground="blue").grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # 初始更新UI元素状态
        self.update_tts_options()
        
        # 添加字幕设置框架
        subtitle_frame = ttk.LabelFrame(options_frame, text="字幕设置", padding="5")
        subtitle_frame.grid(row=4, column=0, columnspan=4, sticky=tk.W+tk.E, pady=5)
        
        # 字幕背景颜色选择
        ttk.Label(subtitle_frame, text="背景颜色:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.subtitle_bg_color = tk.StringVar(value="无背景")
        bg_color_combo = ttk.Combobox(subtitle_frame, textvariable=self.subtitle_bg_color, width=15, state="readonly")
        bg_color_combo['values'] = ("无背景", "白色半透明", "黑色半透明", "蓝色半透明", "灰色半透明")
        bg_color_combo.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 添加字体颜色RGB设置 - 向右移动40像素
        ttk.Label(subtitle_frame, text="字体颜色(RGB):").grid(row=0, column=2, sticky=tk.W, pady=5, padx=(10, 0))  # 原来是(10, 0)，增加到(50, 0)
        self.font_color_rgb = tk.StringVar(value="44, 84, 162")  # 默认深蓝色
        font_color_entry = ttk.Entry(subtitle_frame, textvariable=self.font_color_rgb, width=12)
        font_color_entry.grid(row=0, column=3, sticky=tk.W, pady=5)
        
        # 字体大小设置
        ttk.Label(subtitle_frame, text="字体大小:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.font_size = tk.IntVar(value=60)  # 默认字体大小
        
        # 创建包含滑块和标签的框架
        font_size_container = ttk.Frame(subtitle_frame)
        font_size_container.grid(row=1, column=1, sticky=tk.W+tk.E, pady=5)
        
        # 滑块
        font_scale = ttk.Scale(font_size_container, from_=28, to=70, 
                              variable=self.font_size, orient=tk.HORIZONTAL, length=120)
        font_scale.pack(side=tk.LEFT)
        
        # 字体大小数值标签
        self.font_size_label = ttk.Label(font_size_container, text="60", width=3)
        self.font_size_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # 添加精准字幕复选框 - 移动到右侧与字体颜色标签水平对齐
        ttk.Checkbutton(subtitle_frame, text="精准字幕", variable=self.precise_subtitle).grid(
            row=1, column=2, sticky=tk.W, pady=5, padx=(10, 0))
        
        # 添加图片水印设置区域
        #ttk.Separator(subtitle_frame, orient=tk.HORIZONTAL).grid(row=2, column=0, columnspan=4, sticky=tk.EW, pady=8)
        
        # 添加图片水印复选框
        self.use_watermark = tk.BooleanVar(value=False)
        ttk.Checkbutton(subtitle_frame, text="添加水印", variable=self.use_watermark).grid(
            row=0, column=4, sticky=tk.W, pady=5, padx=(20, 0))
        
        # 添加选择图片按钮
        ttk.Button(subtitle_frame, text="选择图片", command=self.select_watermark_image).grid(
            row=0, column=5, sticky=tk.W, pady=5)
        
        # 添加图片路径显示标签
        self.watermark_path = tk.StringVar(value="未选择图片")
        ttk.Label(subtitle_frame, textvariable=self.watermark_path, width=30, 
                 font=("Arial", 8), foreground="gray").grid(
            row=0, column=6, sticky=tk.W, pady=5)
        
        # 添加水印透明度设置
        ttk.Label(subtitle_frame, text="水印透明度:").grid(row=1, column=4, sticky=tk.W, pady=5, padx=(20, 0))
        
        # 创建包含滑块和标签的框架
        watermark_opacity_container = ttk.Frame(subtitle_frame)
        watermark_opacity_container.grid(row=1, column=5, sticky=tk.W+tk.E, pady=5)
        
        # 透明度变量和滑块
        self.watermark_opacity = tk.IntVar(value=100)  # 默认30%透明度
        opacity_scale = ttk.Scale(watermark_opacity_container, from_=5, to=100, 
                                 variable=self.watermark_opacity, orient=tk.HORIZONTAL, length=80)
        opacity_scale.pack(side=tk.LEFT)
        
        # 透明度数值标签
        self.watermark_opacity_label = ttk.Label(watermark_opacity_container, text="100%", width=4)
        self.watermark_opacity_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # 更新滑块数值显示的函数
        def update_opacity_label(*args):
            self.watermark_opacity_label.config(text=f"{self.watermark_opacity.get()}%")
        
        # 绑定滑块值变化事件
        self.watermark_opacity.trace_add("write", update_opacity_label)
        
        # 更新滑块数值显示的函数
        def update_font_size_label(*args):
            self.font_size_label.config(text=str(self.font_size.get()))
        
        # 绑定滑块值变化事件
        self.font_size.trace_add("write", update_font_size_label)
        
        # 添加字幕设置说明
        #ttk.Label(subtitle_frame, text="注意: 字幕将显示在视频底部，背景色可以提高可读性",
        #         foreground="blue").grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # 添加多音字替换设置框架
        pronunciation_frame = ttk.LabelFrame(options_frame, text="多音字优化", padding="5")
        pronunciation_frame.grid(row=5, column=0, columnspan=4, sticky=tk.W+tk.E, pady=5)
        
        # 添加多音字替换输入框
        ttk.Label(pronunciation_frame, text="替换规则:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.pronunciation_replacements = tk.StringVar(value="压=鸭")  # 默认替换规则
        replacement_entry = ttk.Entry(pronunciation_frame, textvariable=self.pronunciation_replacements, width=50)
        replacement_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        
        # 添加说明文本
        ttk.Label(pronunciation_frame, 
                 text="格式: 原字=替换字, 例如: 压=鸭,重=虫 使用逗号或分号分隔多个规则。字幕中仍显示原始文字", 
                 foreground="blue").grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=2)
        #ttk.Label(pronunciation_frame, 
        #         text="注意: 字幕中仍显示原始文字，替换仅影响语音发音", 
        #         foreground="blue").grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        # Control buttons - 移到日志框上方
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_frame, text="转换", command=self.start_conversion).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="退出", command=self.on_close).pack(side=tk.RIGHT, padx=5)
        
        # Log output
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=80, height=12)  # 减小日志框高度
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state=tk.DISABLED)
        
        # Redirect stdout to the log text widget
        self.stdout_redirector = TextRedirector(self.log_text)
        sys.stdout = self.stdout_redirector
        
        # Progress
        progress_frame = ttk.LabelFrame(main_frame, text="进度", padding="10")
        progress_frame.pack(fill=tk.X, pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, length=400)
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        status_label.pack(anchor=tk.W, pady=5)
    
    def update_tts_options(self):
        """根据选择的TTS引擎更新UI显示"""
        engine = self.tts_engine_var.get()
        
        # 隐藏所有TTS选项框架
        self.xfyun_frame.grid_remove()
        self.ttsmaker_frame.grid_remove()
        
        # 根据选择显示相应选项
        if (engine == "xfyun"):
            self.xfyun_frame.grid()
        elif (engine == "ttsmaker"):
            self.ttsmaker_frame.grid()
    
    def find_ppt_files(self, directory):
        """递归查找目录及其子目录中的所有PPT文件"""
        ppt_files = []
        
        for root, dirs, files in os.walk(directory):
            for file in files:
                self.file_scan_count += 1
                # 每扫描100个文件更新一次状态
                if self.file_scan_count % 100 == 0:
                    self.status_var.set(f"正在扫描文件夹: 已检查 {self.file_scan_count} 个文件...")
                    self.root.update_idletasks()
                
                if file.lower().endswith(('.pptx', '.ppt')):
                    ppt_path = os.path.join(root, file)
                    ppt_files.append(ppt_path)
        
        return ppt_files
    
    def browse_ppt_files(self):
        """选择和添加PPT文件"""
        filenames = filedialog.askopenfilenames(
            title="选择PPT文件",
            filetypes=[("PowerPoint文件", "*.pptx *.ppt"), ("所有文件", "*.*")]
        )
        
        self._add_files_to_list(list(filenames) if filenames else [])
    
    def browse_ppt_folder(self):
        """选择和添加包含PPT的文件夹"""
        directory = filedialog.askdirectory(
            title="选择包含PPT文件的文件夹"
        )
        
        if directory:
            # 更新状态
            self.status_var.set("正在扫描文件夹，请稍候...")
            self.root.update_idletasks()
            
            # 重置文件计数器
            self.file_scan_count = 0
            
            # 使用线程避免UI卡死
            def scan_directory():
                ppt_files = self.find_ppt_files(directory)
                
                # 在UI线程中处理结果
                self.root.after(0, lambda: self._add_files_to_list(ppt_files))
                self.root.after(0, lambda: self.status_var.set(
                    f"文件夹扫描完成。找到 {len(ppt_files)} 个PPT文件，共检查了 {self.file_scan_count} 个文件"))
            
            threading.Thread(target=scan_directory, daemon=True).start()
    
    def browse_ppt(self):
        """保留此方法以兼容可能的外部调用，转发到文件选择方法"""
        self.browse_ppt_files()
    
    def _add_files_to_list(self, new_files):
        """将文件添加到列表中（内部方法，被browse_ppt调用）"""
        if not new_files:
            return
            
        # 检查文件是否已在列表中
        added_count = 0
        
        # 先启用Text组件以添加内容
        self.files_text.configure(state=tk.NORMAL)
        
        for file_path in new_files:
            if file_path not in self.ppt_files:
                self.ppt_files.append(file_path)
                # 添加文件名到Text组件
                file_name = os.path.basename(file_path)
                self.files_text.insert(tk.END, f"{file_name}\n", "default")
                added_count += 1
        
        # 重新禁用Text组件
        self.files_text.configure(state=tk.DISABLED)
        
        print(f"添加了 {added_count} 个新文件，当前共有 {len(self.ppt_files)} 个PPT文件")
        
        # 更新状态显示
        self.status_var.set(f"已选择 {len(self.ppt_files)} 个文件")
    
    def on_file_click(self, event):
        """处理文件列表的点击事件"""
        # 获取点击的行号
        index = self.files_text.index(f"@{event.x},{event.y}")
        line = int(index.split(".")[0]) - 1  # 行号从1开始，转换为0开始的索引
        
        # 确保点击的是有效行
        if 0 <= line < len(self.ppt_files):
            # 先启用Text组件以修改标记
            self.files_text.configure(state=tk.NORMAL)
            
            # 检查当前行是否已选中
            current_tags = self.files_text.tag_names(f"{line+1}.0")
            
            # 默认标签总是存在的，检查是否有选择的标签
            if "selected" in current_tags:
                # 已经选中，现在取消选择
                self.files_text.tag_remove("selected", f"{line+1}.0", f"{line+1}.end")
                if line in self.selected_indices:
                    self.selected_indices.remove(line)
            else:
                # 未选中，现在选择（按下Ctrl键时为多选）
                if not (event.state & 0x0004):  # 检查Ctrl键是否按下
                    # 单选模式：先清除所有选择
                    for idx in self.selected_indices:
                        self.files_text.tag_remove("selected", f"{idx+1}.0", f"{idx+1}.end")
                    self.selected_indices = []
                
                # 添加新选择
                self.files_text.tag_add("selected", f"{line+1}.0", f"{line+1}.end")
                self.selected_indices.append(line)
                
            # 配置选中标签的外观
            self.files_text.tag_configure("selected", background="lightblue")
                
            # 重新禁用Text组件
            self.files_text.configure(state=tk.DISABLED)
    
    def delete_selected_files(self):
        """删除列表中选中的文件"""
        if not self.selected_indices:
            messagebox.showinfo("提示", "请先选择要删除的文件")
            return
        
        # 排序并反转，以便从后往前删除
        selected_to_remove = sorted(self.selected_indices, reverse=True)
        
        # 启用Text组件以删除内容
        self.files_text.configure(state=tk.NORMAL)
        
        # 从列表和Text组件中删除
        for idx in selected_to_remove:
            # 从原始文件列表中删除
            if 0 <= idx < len(self.ppt_files):
                del self.ppt_files[idx]
                # 从文本框中删除对应行
                self.files_text.delete(f"{idx+1}.0", f"{idx+2}.0")
        
        # 清理行号索引
        self.files_text.delete("1.0", tk.END)
        # 重新添加所有文件
        for file_path in self.ppt_files:
            file_name = os.path.basename(file_path)
            self.files_text.insert(tk.END, f"{file_name}\n", "default")
        
        # 重新禁用Text组件
        self.files_text.configure(state=tk.DISABLED)
        
        # 清空选择
        self.selected_indices = []
        
        print(f"已删除选中文件，当前剩余 {len(self.ppt_files)} 个PPT文件")
        
        # 更新状态显示
        if self.ppt_files:
            self.status_var.set(f"已选择 {len(self.ppt_files)} 个文件")
        else:
            self.status_var.set("未选择文件")
    
    def start_conversion(self):
        if not self.ppt_files:
            messagebox.showerror("错误", "请选择至少一个PPT文件")
            return
            
        # Clear log
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)
        
        tts_engine = self.tts_engine_var.get()
        
        # 科大讯飞参数
        xfyun_params = None
        if tts_engine == "xfyun":
            # 验证科大讯飞参数是否齐全
            if not self.xfyun_app_id.get() or not self.xfyun_api_key.get() or not self.xfyun_api_secret.get():
                messagebox.showerror("错误", "使用科大讯飞TTS需要填写完整的APPID、API Key和API Secret")
                return
                
            xfyun_params = {
                'app_id': self.xfyun_app_id.get(),
                'api_key': self.xfyun_api_key.get(),
                'api_secret': self.xfyun_api_secret.get(),
                'voice_name': self.xfyun_voice.get(),
                'speed': self.xfyun_speed.get()  # 添加语速参数
            }
        
        # 马克配音参数
        ttsmaker_params = None
        if tts_engine == "ttsmaker":
            # 验证语音ID是否为数字
            try:
                voice_id = int(self.ttsmaker_voice_id.get())
            except ValueError:
                messagebox.showerror("错误", "语音ID必须是数字")
                return
                
            ttsmaker_params = {
                'token': self.ttsmaker_token.get(),
                'voice_id': voice_id,
                'audio_speed': self.ttsmaker_speed.get()  # 添加语速参数
            }
        
        # 处理多音字替换设置
        pronunciation_dict = {}
        replacements_text = self.pronunciation_replacements.get().strip()
        if replacements_text:
            # 分割替换规则 (允许逗号或分号分隔)
            rules = re.split(r'[,，；，，]+', replacements_text)
            for rule in rules:
                if '=' in rule:
                    original, replacement = rule.split('=', 1)
                    if original and replacement:  # 确保两边都不为空
                        pronunciation_dict[original] = replacement
        
        # 处理RGB颜色值
        try:
            rgb_str = self.font_color_rgb.get().strip()
            rgb_parts = [int(x.strip()) for x in rgb_str.split(',')]
            if len(rgb_parts) != 3:
                raise ValueError("RGB值必须包含三个数字")
            for val in rgb_parts:
                if not (0 <= val <= 255):
                    raise ValueError("RGB值必须在0-255范围内")
            font_color = tuple(rgb_parts)
        except Exception as e:
            # 如果解析失败，使用默认值
            print(f"RGB值格式错误 ({e})，使用默认颜色(38, 74, 145)")
            font_color = (38, 74, 145)
        
        # 字幕设置参数
        subtitle_params = {
            'bg_color': self.subtitle_bg_color.get(),
            'font_size': self.font_size.get(),
            'font_color': font_color,
            'precise_subtitle': self.precise_subtitle.get()
        }
        
        # 添加水印设置
        watermark_params = None
        if self.use_watermark.get() and hasattr(self, '_full_watermark_path') and self._full_watermark_path:
            watermark_params = {
                'image_path': self._full_watermark_path,
                'opacity': self.watermark_opacity.get() / 100.0  # 转换为0-1范围
            }
            print(f"启用水印: {os.path.basename(self._full_watermark_path)}, 透明度: {self.watermark_opacity.get()}%")
        else:
            print("水印功能未启用或未选择图片")
        
        print("="*50)
        print(f"开始批量转换 {len(self.ppt_files)} 个文件")
        print(f"语音引擎：{tts_engine}")
        if tts_engine == "xfyun":
            print(f"科大讯飞发音人：{self.xfyun_voice.get()}")
            print(f"语速：{self.xfyun_speed.get()}")  # 添加语速日志输出
        elif tts_engine == "ttsmaker":
            print(f"马克配音语音ID：{self.ttsmaker_voice_id.get()}")
            print(f"语速：{self.ttsmaker_speed.get():.2f}倍")  # 修改为显示2位小数
        print(f"字幕背景颜色: {subtitle_params['bg_color']}")
        print(f"字幕字体大小: {subtitle_params['font_size']}")
        print(f"字幕字体颜色: RGB{font_color}")  # 添加字体颜色日志输出
        print(f"精准字幕: {'启用' if subtitle_params['precise_subtitle'] else '禁用'}")
        
        # 打印多音字替换设置
        if pronunciation_dict:
            print("多音字替换规则:")
            for orig, repl in pronunciation_dict.items():
                print(f"  {orig} → {repl}")
        else:
            print("未设置多音字替换规则")
        
        print("="*50)
        
        # Disable buttons during conversion
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.configure(state=tk.DISABLED)
        
        # Set status
        self.status_var.set("准备批量转换...")
        self.progress_var.set(0)
        
        # Start conversion in a separate thread
        conversion_thread = threading.Thread(
            target=self.run_batch_conversion,
            args=(self.ppt_files, tts_engine, xfyun_params, ttsmaker_params, subtitle_params, pronunciation_dict, watermark_params)
        )
        conversion_thread.daemon = True
        conversion_thread.start()
    
    def run_batch_conversion(self, ppt_files, tts_engine, xfyun_params=None, ttsmaker_params=None, subtitle_params=None, pronunciation_dict=None, watermark_params=None):
        """批量处理多个PPT文件"""
        total_files = len(ppt_files)
        success_count = 0
        failed_files = []
        
        # 计算每个文件的进度增量
        progress_increment = 100.0 / total_files
        
        for index, ppt_path in enumerate(ppt_files):
            # 更新状态
            file_name = os.path.basename(ppt_path)
            status_text = f"正在处理 ({index + 1}/{total_files}): {file_name}"
            self.root.after(0, lambda t=status_text: self.status_var.set(t))
            
            # 自动生成输出路径 - 使用相同目录，更改扩展名为.mp4
            output_path = os.path.splitext(ppt_path)[0] + ".mp4"
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if not os.path.exists(output_dir):
                try:
                    os.makedirs(output_dir, exist_ok=True)
                except Exception as e:
                    print(f"创建输出目录失败: {e}")
                    # 如果无法创建目录，使用临时目录
                    import tempfile
                    output_dir = tempfile.gettempdir()
                    output_filename = os.path.basename(output_path)
                    output_path = os.path.join(output_dir, output_filename)
                    print(f"将使用临时目录: {output_path}")
            
            print("\n" + "="*50)
            print(f"开始处理文件 {index + 1}/{total_files}: {ppt_path}")
            print(f"输出路径: {output_path}")
            print("="*50)
            
            try:
                # 调用单个文件转换函数
                ppt_to_video(
                    ppt_path, 
                    output_path, 
                    tts_engine, 
                    None,  # language 参数
                    xfyun_params, 
                    ttsmaker_params, 
                    subtitle_params, 
                    pronunciation_dict,
                    watermark_params  # 添加水印参数
                )
                
                print(f"文件 {file_name} 处理成功!")
                success_count += 1
                
                # 更新文件列表中的状态为成功(绿色)
                self.root.after(0, lambda idx=index: self.update_file_status(idx, "success"))
            
            except Exception as e:
                print(f"文件 {file_name} 处理失败: {str(e)}")
                print(traceback.format_exc())
                failed_files.append((file_name, str(e)))
                
                # 更新文件列表中的状态为失败(红色)
                self.root.after(0, lambda idx=index: self.update_file_status(idx, "failed"))
            
            # 更新进度条 - 每完成一个文件增加相应的进度
            current_progress = (index + 1) * progress_increment
            self.root.after(0, lambda p=current_progress: self.progress_var.set(p))
        
        # 处理完成，更新状态
        summary = f"转换完成! 成功: {success_count}/{total_files}"
        if failed_files:
            summary += f", 失败: {len(failed_files)}"
            
        self.root.after(0, lambda s=summary: self.status_var.set(s))
        
        # 显示处理汇总
        print("\n" + "="*50)
        print(f"批处理完成！成功: {success_count}/{total_files}")
        
        if failed_files:
            print("\n失败文件列表:")
            for name, error in failed_files:
                print(f"• {name}: {error}")
        
        print("="*50)
        
        # 重新启用按钮
        self.root.after(0, self.enable_buttons)
        
        # 显示汇总消息框
        if failed_files:
            message = f"转换完成！\n\n成功: {success_count}/{total_files}\n失败: {len(failed_files)}/{total_files}\n\n请查看日志了解详情。"
            self.root.after(0, lambda: messagebox.showinfo("批处理完成", message))
        else:
            self.root.after(0, lambda: messagebox.showinfo("成功", f"所有 {total_files} 个文件均已成功转换!"))
    
    def update_file_status(self, index, status):
        """更新文件列表中指定索引的文件状态颜色"""
        if index >= 0 and index < len(self.ppt_files):
            # 启用文本组件进行编辑
            self.files_text.configure(state=tk.NORMAL)
            
            # 为了安全起见，先删除所有可能的状态标签
            self.files_text.tag_remove("success", f"{index+1}.0", f"{index+1}.end")
            self.files_text.tag_remove("failed", f"{index+1}.0", f"{index+1}.end")
            
            # 添加指定的状态标签
            self.files_text.tag_add(status, f"{index+1}.0", f"{index+1}.end")
            
            # 禁用文本组件
            self.files_text.configure(state=tk.DISABLED)
            
            # 存储状态
            self.file_statuses[index] = status
    
    def enable_buttons(self):
        """重新启用所有按钮"""
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.configure(state=tk.NORMAL)
    
    def test_xfyun_api(self):
        """测试科大讯飞API连接是否正常"""
        from ppt_to_video_converter import xfyun_tts
        import tempfile
        import webbrowser
        import datetime  # 在函数内再次导入以确保可用
        
        # 清空日志
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)
        
        # 验证API信息是否填写完整
        if not self.xfyun_app_id.get() or not self.xfyun_api_key.get() or not self.xfyun_api_secret.get():
            messagebox.showerror("错误", "请填写完整的科大讯飞API信息")
            return
            
        # 创建临时文件用于测试
        test_file = tempfile.NamedTemporaryFile(suffix='.mp3', delete=False)
        test_file.close()
        
        print("正在测试科大讯飞API连接...")
        print(f"APPID: {self.xfyun_app_id.get()}")
        print(f"API Key: {self.xfyun_api_key.get()[:4]}...{self.xfyun_api_key.get()[-4:]}")
        print(f"API Secret: {self.xfyun_api_secret.get()[:4]}...{self.xfyun_api_secret.get()[-4:]}")
        print(f"发音人: {self.xfyun_voice.get()}")
        
        # 添加提示信息，特别针对科大讯飞反馈
        print("\n=== 重要提示 ===")
        print("科大讯飞官方指出，403错误通常由以下原因导致:")
        print("1. 系统时间与标准时间相差超过5分钟")
        print("2. IP白名单设置问题")
        print("\n请确保:")
        print("- 您的系统时间是准确的")
        print("- 关闭科大讯飞控制台的IP白名单功能")
        print("===============\n")
        
        # 先尝试检查系统时间
        try:
            # 如果安装了ntplib，检查系统时间
            import ntplib
            from time import ctime
            
            print("正在检查系统时间...")
            try:
                ntp_client = ntplib.NTPClient()
                response = ntp_client.request('pool.ntp.org', timeout=1)
                system_time = datetime.datetime.now()
                ntp_time = datetime.datetime.fromtimestamp(response.tx_time)
                time_diff = abs((system_time - ntp_time).total_seconds())
                
                print(f"系统时间: {system_time}")
                print(f"网络时间: {ntp_time}")
                print(f"时差: {time_diff:.1f}秒")
                
                if time_diff > 300:
                    if messagebox.askyesno("时间不同步警告", 
                        f"您的系统时间与标准时间相差过大({time_diff:.1f}秒)，这会导致科大讯飞API认证失败。\n\n"
                        f"是否立即调整系统时间？"):
                        # 尝试同步系统时间
                        import os
                        if os.name == 'nt':  # Windows
                            os.system('w32tm /resync')
                            messagebox.showinfo("时间同步", "已尝试同步系统时间。请重新测试API连接。")
                            return
                        elif os.name == 'posix':  # Linux/Mac
                            os.system('sudo ntpdate pool.ntp.org')
                            messagebox.showinfo("时间同步", "已尝试同步系统时间。请重新测试API连接。")
                            return
            except Exception as e:
                print(f"时间检查失败: {e}")
                print("继续尝试API连接...")
        except ImportError:
            print("无法检查系统时间（ntplib模块未安装）")
        
        # 添加按钮禁用
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.configure(state=tk.DISABLED)
        
        # 在单独的线程中测试以避免UI冻结
        def run_test():
            try:
                test_text = "这是一条测试消息，用于验证科大讯飞API是否正常工作。"
                result = xfyun_tts(
                    test_text, 
                    test_file.name,
                    self.xfyun_app_id.get(),
                    self.xfyun_api_key.get(),
                    self.xfyun_api_secret.get(),
                    voice=self.xfyun_voice.get()
                )
                
                # 重新启用按钮
                self.root.after(0, lambda: [widget.configure(state=tk.NORMAL) for widget in self.root.winfo_children() if isinstance(widget, ttk.Button)])
                
                if result:
                    self.root.after(0, lambda: messagebox.showinfo("成功", "科大讯飞API连接测试成功！已生成测试音频。"))
                    # 打开包含音频文件的文件夹
                    try:
                        import platform
                        if platform.system() == "Windows":
                            os.system(f'explorer /select,"{test_file.name}"')
                        elif platform.system() == "Darwin":  # macOS
                            os.system(f'open -R "{test_file.name}"')
                    except:
                        pass
                else:
                    # 在日志中看到403错误时，显示特殊对话框
                    def show_failure_dialog():
                        log_content = self.get_log_content().lower()
                        if "403" in log_content or "forbidden" in log_content:
                            result = messagebox.askquestion(
                                "API连接失败 (403错误)", 
                                "科大讯飞API返回403错误。根据官方反馈，这通常由以下原因导致：\n\n"
                                "1. 系统时间与标准时间相差超过5分钟\n"
                                "2. IP白名单设置问题\n\n"
                                "请确保：\n"
                                "- 您的系统时间是准确的\n"
                                "- 关闭科大讯飞控制台的IP白名单功能\n\n"
                                "是否打开科大讯飞控制台和故障排除指南？"
                            )
                            
                            if result == 'yes':
                                # 打开科大讯飞控制台和故障排除指南
                                webbrowser.open("https://console.xfyun.cn/")
                                
                                # 如果故障排除文档存在，则打开它
                                troubleshooting_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 
                                                                   "XFYUN_TROUBLESHOOTING.md")
                                if os.path.exists(troubleshooting_path):
                                    try:
                                        os.startfile(troubleshooting_path)
                                    except:
                                        pass
                        else:
                            messagebox.showerror("API连接失败", 
                                "科大讯飞API连接测试失败。请查看日志获取详细错误信息。")
                            
                        # 无论如何，尝试使用系统TTS作为备选
                        fallback_success = False
                        try:
                            import pyttsx3
                            engine = pyttsx3.init()
                            engine.save_to_file(test_text, test_file.name)
                            engine.runAndWait()
                            fallback_success = True
                            
                            if fallback_success:
                                messagebox.showinfo(
                                    "备用方案", 
                                    "已使用系统TTS生成备用音频。\n\n"
                                    "在解决科大讯飞API问题前，您可以临时使用系统TTS进行转换。"
                                )
                                # 打开包含音频文件的文件夹
                                try:
                                    import platform
                                    if platform.system() == "Windows":
                                        os.system(f'explorer /select,"{test_file.name}"')
                                    elif platform.system() == "Darwin":  # macOS
                                        os.system(f'open -R "{test_file.name}"')
                                except:
                                    pass
                        except Exception as e:
                            print(f"备用TTS也失败: {e}")
                            
                    self.root.after(0, show_failure_dialog)
            except Exception as e:
                print(f"测试过程出现异常: {e}")
                print(traceback.format_exc())
                self.root.after(0, lambda: [widget.configure(state=tk.NORMAL) for widget in self.root.winfo_children() if isinstance(widget, ttk.Button)])
                self.root.after(0, lambda: messagebox.showerror("错误", f"测试过程出现异常: {e}"))
        
        threading.Thread(target=run_test, daemon=True).start()
    
    def get_log_content(self):
        """获取日志内容"""
        self.log_text.configure(state=tk.NORMAL)
        content = self.log_text.get(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)
        return content
    
    def open_voice_list(self):
        """打开浏览器查看TTS Maker语音列表"""
        import webbrowser
        url = "https://ttsmaker.cn/voice-cloning-ai-voices"
        webbrowser.open(url)
        print("已打开浏览器查看语音列表")
    
    def test_ttsmaker_api(self):
        """测试马克配音API连接是否正常"""
        from ppt_to_video_converter import ttsmaker_tts
        import tempfile
        
        # 清空日志
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)
        
        # 验证voice_id是否为有效数字
        try:
            voice_id = int(self.ttsmaker_voice_id.get())
        except ValueError:
            messagebox.showerror("错误", "语音ID必须是数字")
            return
            
        # 验证语速是否为有效范围内的数字
        try:
            speed = float(self.ttsmaker_speed.get())
            if not (0.5 <= speed <= 2.0):
                messagebox.showerror("错误", "语速必须在0.5到2.0之间")
                return
                
            # 确保语速符合X.X0或X.X5格式
            decimal_part = int((speed * 100) % 100)
            if decimal_part % 5 != 0:
                # 自动修正为最接近的X.X0或X.X5格式
                corrected_speed = round(speed * 20) / 20
                self.ttsmaker_speed.set(corrected_speed)
                speed = corrected_speed
                
        except ValueError:
            messagebox.showerror("错误", "语速必须是数字")
            return
            
        # 创建临时文件用于测试
        test_file = tempfile.NamedTemporaryFile(suffix='.mp3', delete=False)
        test_file.close()
        
        print("正在测试马克配音API连接...")
        print(f"Token: {'*'*len(self.ttsmaker_token.get())}")
        print(f"语音ID: {voice_id}")
        
        # 添加按钮禁用
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.configure(state=tk.DISABLED)
        
        # 在单独的线程中测试以避免UI冻结
        def run_test():
            try:
                test_text = "这是一条测试消息，用于验证马克配音API是否正常工作。"
                result = ttsmaker_tts(
                    test_text, 
                    test_file.name,
                    token=self.ttsmaker_token.get(),
                    voice_id=voice_id,
                    audio_speed=self.ttsmaker_speed.get()  # 添加语速参数
                )
                
                # 重新启用按钮
                self.root.after(0, lambda: [widget.configure(state=tk.NORMAL) for widget in self.root.winfo_children() if isinstance(widget, ttk.Button)])
                
                if result:
                    self.root.after(0, lambda: messagebox.showinfo("成功", "马克配音API连接测试成功！已生成测试音频。"))
                    # 打开包含音频文件的文件夹
                    try:
                        import platform
                        if platform.system() == "Windows":
                            os.system(f'explorer /select,"{test_file.name}"')
                        elif platform.system() == "Darwin":  # macOS
                            os.system(f'open -R "{test_file.name}"')
                    except:
                        pass
                else:
                    self.root.after(0, lambda: messagebox.showerror("API连接失败", 
                        "马克配音API连接测试失败。请查看日志获取详细错误信息。"))
                    
                    # 无论如何，尝试使用系统TTS作为备选
                    fallback_success = False
                    try:
                        import pyttsx3
                        engine = pyttsx3.init()
                        engine.save_to_file(test_text, test_file.name)
                        engine.runAndWait()
                        fallback_success = True
                        
                        if fallback_success:
                            self.root.after(0, lambda: messagebox.showinfo(
                                "备用方案", 
                                "已使用系统TTS生成备用音频。\n\n"
                                "在解决马克配音API问题前，您可以临时使用系统TTS进行转换。"
                            ))
                            # 打开包含音频文件的文件夹
                            try:
                                import platform
                                if platform.system() == "Windows":
                                    os.system(f'explorer /select,"{test_file.name}"')
                                elif platform.system() == "Darwin":  # macOS
                                    os.system(f'open -R "{test_file.name}"')
                            except:
                                pass
                    except Exception as e:
                        print(f"备用TTS也失败: {e}")
            except Exception as e:
                print(f"测试过程出现异常: {e}")
                print(traceback.format_exc())
                self.root.after(0, lambda: [widget.configure(state=tk.NORMAL) for widget in self.root.winfo_children() if isinstance(widget, ttk.Button)])
                self.root.after(0, lambda: messagebox.showerror("错误", f"测试过程出现异常: {e}"))
        
        threading.Thread(target=run_test, daemon=True).start()
    
    def select_watermark_image(self):
        """选择水印图片文件"""
        filename = filedialog.askopenfilename(
            title="选择水印图片",
            filetypes=[("图片文件", "*.png *.jpg *.jpeg *.bmp"), ("所有文件", "*.*")]
        )
        
        if filename:
            # 检查文件大小
            try:
                file_size = os.path.getsize(filename) / (1024 * 1024)  # 转换为MB
                if file_size > 5:
                    messagebox.showwarning("文件过大", f"选择的图片文件大小为 {file_size:.1f}MB，建议使用小于5MB的图片作为水印。")
                
                # 如果文件有效，更新路径变量
                self.watermark_path.set(os.path.basename(filename))
                # 存储完整路径
                self._full_watermark_path = filename
                print(f"已选择水印图片: {filename}")
            except Exception as e:
                messagebox.showerror("错误", f"读取图片文件时出错: {e}")
                self.watermark_path.set("未选择图片")
                self._full_watermark_path = None
    
    def on_close(self):
        # Restore stdout
        sys.stdout = sys.__stdout__
        self.root.destroy()

# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = PPTToVideoApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()

