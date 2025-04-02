import os
import tempfile
import time
from moviepy.editor import *
import pyttsx3
import re
import win32com.client
import sys
import traceback
import requests
import numpy as np
from PIL import Image as PILImage, ImageDraw as PILImageDraw, ImageFont as PILImageFont  # Rename all PIL imports
import base64
import json
import websocket
import wave
import threading
import time
import hmac
import hashlib
import urllib
from urllib.parse import urlencode
import datetime
import ssl  # 添加ssl模块导入

# 添加字幕时长估算函数
def estimate_line_duration(text, cn_char_duration=0.2048, en_word_duration=0.35, en_char_duration=0.1, digit_duration=0.3233, punctuation_factor=0.8251):
    """
    Estimate line duration based on text content, including punctuation pauses, excluding ending punctuation.

    Args:
        text (str): Input text.
        cn_char_duration (float): Duration per Chinese character in seconds.
        en_word_duration (float): Duration per English word in seconds.
        en_char_duration (float): Duration per English character for uppercase words in seconds.
        digit_duration (float): Duration per digit in seconds.
        punctuation_factor (float): Factor to adjust punctuation pauses duration.

    Returns:
        float: Estimated duration in seconds.
    """
    # Count Chinese characters
    chinese_chars = re.findall(r'[\u4e00-\u9fff]', text)
    cn_duration = len(chinese_chars) * cn_char_duration

    # English duration calculation
    uppercase_words = re.findall(r'\b[A-Z]+\b', text)
    normal_words = re.findall(r'\b[a-zA-Z]*[a-z][a-zA-Z]*\b', text)

    uppercase_duration = sum(len(word) * en_char_duration for word in uppercase_words)
    normal_word_duration = len(normal_words) * en_word_duration

    en_duration = uppercase_duration + normal_word_duration

    # Count digits
    digits = re.findall(r'\d', text)
    digit_total_duration = len(digits) * digit_duration

    # Punctuation pause durations
    punctuation_pauses = {
        '，': 0.3, '、': 0.25, '；': 0.4,
        '。': 0.5, '！': 0.5, '？': 0.5,
        '：': 0.35, '—': 0.5, '……': 0.5
    }

    # Exclude all trailing punctuation (not just one)
    text_excluding_end = re.sub(r'[，、；。！？：—……]+$', '', text)

    punctuation_duration = sum(punctuation_pauses.get(char, 0) for char in text_excluding_end) * punctuation_factor

    # Total estimated line duration
    line_duration = cn_duration + en_duration + digit_total_duration + punctuation_duration

    return line_duration

def ppt_to_video(ppt_path, output_video_path, tts_engine="ttsmaker", language=None, xfyun_params=None, ttsmaker_params=None, subtitle_params=None, pronunciation_dict=None):
    """
    Convert PowerPoint presentation to video with narration.
    
    Parameters:
    - ppt_path: Path to PowerPoint file
    - output_video_path: Path for the output video
    - tts_engine: Text-to-speech engine to use ('pyttsx3', 'xfyun', or 'ttsmaker')
    - language: 未使用的参数，保留用于向后兼容
    - xfyun_params: Dictionary containing iFLYTEK parameters (only used if tts_engine is 'xfyun')
                    Required keys: 'app_id', 'api_key', 'api_secret', 'voice_name'
                    Optional keys: 'speed' (0-100)
    - ttsmaker_params: Dictionary containing TTS Maker parameters (only used if tts_engine is 'ttsmaker')
                    Required keys: 'token', 'voice_id'
    - subtitle_params: Dictionary containing subtitle formatting parameters
                    Optional keys: 'bg_color', 'font_size', 'font_color'
    - pronunciation_dict: Dictionary mapping characters to their preferred pronunciation
                    Example: {'压': '鸭', '参': '餐'}
    """
    # Convert paths to absolute paths
    ppt_path = os.path.abspath(ppt_path)
    output_video_path = os.path.abspath(output_video_path)
    
    # Check if the PowerPoint file exists
    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PowerPoint file not found: {ppt_path}")
    
    # 设置默认字幕参数
    if subtitle_params is None:
        subtitle_params = {}
    
    font_size = subtitle_params.get('font_size', 30)  # 默认字体大小为30
    bg_color_name = subtitle_params.get('bg_color', '白色半透明')  # 默认背景色
    font_color = subtitle_params.get('font_color', (38, 74, 145))  # 默认字体颜色为深蓝色
    
    # 背景颜色映射
    bg_color_map = {
        '白色半透明': (255, 255, 255, 200),  # 白色,透明度200/255
        '黑色半透明': (0, 0, 0, 180),        # 黑色,透明度180/255
        '蓝色半透明': (51, 153, 255, 180),   # 蓝色,透明度180/255
        '灰色半透明': (128, 128, 128, 200),  # 灰色,透明度200/255
        '无背景': None                       # 无背景
    }
    
    # 获取背景颜色
    bg_color = bg_color_map.get(bg_color_name, (255, 255, 255, 200))
    
    print(f"字幕设置: 背景颜色={bg_color_name}, 字体大小={font_size}, 字体颜色=RGB{font_color}")

    # Create PowerPoint application
    ppt_app = None
    presentation = None
    audio_clips = []  # Keep track of audio clips to ensure proper cleanup
    
    try:
        print(f"启动PowerPoint应用程序...")
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        # 不要设置 Visible = False，这会导致错误
        # 让 PowerPoint 保持可见状态运行
        print(f"PowerPoint应用程序启动成功")
        
        # Open the presentation
        print(f"打开演示文稿: {ppt_path}")
        # 移除 WithWindow=False 参数，使用默认参数打开演示文稿
        presentation = ppt_app.Presentations.Open(ppt_path)
        print(f"演示文稿打开成功，开始处理...")
        
        # Get total number of slides
        total_slides = presentation.Slides.Count
        print(f"演示文稿共有 {total_slides} 张幻灯片")
        
        # 总幻灯片数保持不变，但注意python下标从0开始，所以最后一页的下标为total_slides-1
        if total_slides < 2:
            raise ValueError("演示文稿必须至少包含两张幻灯片，最后一张用于存放旁白文字")
            
        # 确定哪一页是旁白页面（最后一页，索引total_slides-1）
        narration_slide_index = total_slides - 1
        
        # Extract narration from the narration slide (default: last slide)
        print(f"从第 {narration_slide_index+1} 张幻灯片提取旁白文字...")  # 日志显示页码从1开始
        narration_slide = presentation.Slides[narration_slide_index]
        narration_text = ""
        try:
            for shape in narration_slide.Shapes:
                if shape.HasTextFrame:
                    narration_text += shape.TextFrame.TextRange.Text + "\n"
        except Exception as e:
            print(f"提取文字时出错: {e}")
            
        print(f"提取到的文字: {narration_text[:100]}...")
        
        # 确定要处理的幻灯片范围
        print("打印所有幻灯片数量和索引信息:")
        print(f"PowerPoint报告总幻灯片数: {total_slides}")
        for i in range(0, total_slides):
            try:
                slide = presentation.Slides[i]  # 直接使用0-based下标
                slide_name = f"Slide {i+1}"
                print(f"  索引 {i} -> 第 {i+1} 页: {slide_name}")
            except Exception as e:
                print(f"  索引 {i}: 访问出错 - {str(e)}")

        # 确定要处理的幻灯片范围 - 处理除最后一页外的所有幻灯片
        slides_to_process = list(range(0, total_slides - 1))  # 从0到total_slides-2
        print(f"将处理以下幻灯片 (下标)： {slides_to_process}")
        
        # 尝试识别格式化模式 - 检测是否使用"pageX"的格式
        page_pattern = re.compile(r'page(\d+)\s*[:：]\s*(.+?)(?=page\d+\s*[:：]|$)', re.DOTALL | re.IGNORECASE)
        page_matches = page_pattern.findall(narration_text)
        
        # Parse narration text to match with slides
        narrations = {}  # 使用字典存储页码对应的旁白
        raw_narrations = {}  # 用于保存原始文本，不带"page n"前缀
        
        # 如果找到"pageX"格式的内容，直接使用它们
        if page_matches:
            print("检测到page格式的旁白标记")
            
            # 对匹配结果按页码排序
            sorted_matches = sorted(page_matches, key=lambda x: int(x[0]))
            
            # 详细打印页面匹配情况
            page_numbers = [int(match[0]) for match in sorted_matches]
            print(f"找到以下页面标记: {page_numbers}")
            
            # 检查是否缺少页码标记
            all_slides_indices = set(slides_to_process)
            
            # 更正：从pageX转换为PowerPoint索引
            # 现在“page1”对应下标0
            found_slides = set(int(match[0]) - 1 for match in sorted_matches)
            missing_pages = all_slides_indices - found_slides
            
            if missing_pages:
                # print(f"警告: 以下PPT索引缺少旁白标记: {sorted(missing_pages)}")
                # print("将为这些页面使用默认文本")
                error_msg = f"错误: 以下PPT索引缺少旁白标记: {missing_pages}"
                print(error_msg)
                raise ValueError(error_msg)
            
            # 保存文本，去掉"page n"前缀
            for match in sorted_matches:
                page_num = int(match[0])  # "pageX"中的X
                text = match[1].strip()
                
                # 使用页码作为PowerPoint的slide索引（"page1"对应下标0）
                ppt_slide_idx = page_num - 1
                print(f"为幻灯片索引 {ppt_slide_idx} (对应page{page_num})设置旁白")
                narrations[ppt_slide_idx] = text
                raw_narrations[ppt_slide_idx] = text

        else:
            # 尝试其他拆分模式
            slide_texts = re.split(r'(?:Slide\s+\d+[\.:]|^\d+[\.:])\s*', narration_text, flags=re.MULTILINE)
            # 清理提取的文本
            narrations = [text.strip() for text in slide_texts if text.strip()]
            raw_narrations = narrations.copy()  # 在这种情况下，原始文本和处理后的文本相同
            
            # 如果没有找到匹配的模式，尝试按行拆分
            if len(narrations) <= 1 or len(narrations) != total_slides - 1:
                error_msg = f"错误: 旁白数量 ({len(narrations)}) 与要处理的幻灯片数量 ({len(total_slides - 1)}) 不匹配!"
                print(error_msg)
                raise ValueError(error_msg)
                # lines = [line.strip() for line in narration_text.split('\n') if line.strip()]
                # if len(lines) >= total_slides - 1:
                #     narrations = lines[:total_slides - 1]  # 仅使用与幻灯片数量匹配的行数
                # else:
                #     # 如果仍然没有足够的旁白，重复可用的或使用占位符
                #     if len(narrations) > 0:
                #         temp_narrations = []
                #         for i in range(total_slides - 1):
                #             temp_narrations.append(narrations[i % len(narrations)])
                #         narrations = temp_narrations
                #     else:
                #         narrations = ["这是第 {} 张幻灯片".format(i+1) for i in range(total_slides - 1)]
        
        # 清理所有旁白文本中可能存在的"page n:"格式
        for page_num in narrations.keys():
            # 使用正则表达式移除"page n:"或"page n：" 的模式 - 同样忽略大小写
            narrations[page_num] = re.sub(r'^page\s*\d+[:：]\s*', '', narrations[page_num], flags=re.IGNORECASE)
            raw_narrations[page_num] = narrations[page_num]
            print(f"幻灯片 {page_num} 的旁白: {narrations[page_num][:30]}{'...' if len(narrations[page_num]) > 30 else ''}")
        
        print(f"解析得到 {len(narrations)} 段旁白文字")
        
        # 强化检查: 在自动填充默认旁白前，先检查旁白数量是否匹配幻灯片数量
        if len(narrations) != len(slides_to_process):
            error_message = (
                f"错误: 旁白数量 ({len(narrations)}) 与要处理的幻灯片数量 ({len(slides_to_process)}) 不匹配!\n"
                f"请检查PPT最后一页的旁白文本格式，确保:\n"
                f"1. 每个幻灯片都有对应的旁白文本 (当前旁白不足)\n"
                f"2. 没有多余的旁白文本 (当前旁白过多)\n"
                f"3. 如使用page标记，确保编号与实际幻灯片数量一致"
            )
            print("\n" + "!"*50)
            print(error_message)
            print("!"*50 + "\n")
            # 抛出异常，中止处理当前文件
            raise ValueError(error_message)
            
        # 为缺失的幻灯片自动添加旁白（如果严格检查已通过，这里只是为了向后兼容）
        # 实际上在上面的检查通过后，这里不应该再有缺失
        missing_narrations = []
        for slide_idx in slides_to_process:
            if slide_idx not in narrations:
                missing_narrations.append(slide_idx)
                default_text = f"这是第 {slide_idx+1} 张幻灯片"
                narrations[slide_idx] = default_text
                raw_narrations[slide_idx] = default_text
                print(f"幻灯片 {slide_idx} 使用默认旁白文本: {default_text}")
                
        if missing_narrations:
            print(f"警告: 自动为 {len(missing_narrations)} 个缺失旁白的幻灯片添加了默认文本")
        
        # 打印slide索引和narration的对应关系，帮助调试
        print("幻灯片索引与旁白对应关系:")
        for slide_idx in sorted(narrations.keys()):
            print(f"  幻灯片索引 {slide_idx} -> 旁白: {narrations[slide_idx][:30]}{'...' if len(narrations[slide_idx]) > 30 else ''}")
            
        # Create temporary directory for slide images and audio files
        temp_dir = tempfile.mkdtemp()  # Using mkdtemp instead of TemporaryDirectory to avoid auto-cleanup issues
        try:
            print(f"创建临时目录: {temp_dir}")
            
            # Save slides as images
            print("导出幻灯片为图片...")
            for processed_idx, ppt_idx in enumerate(slides_to_process, 1):
                try:

                    slide = presentation.Slides[ppt_idx]
                    slide_name = f"Slide {ppt_idx+1}"
                    print(f"正在导出 PPT索引 {ppt_idx} (第 {ppt_idx+1} 页, {slide_name}) -> 处理序号 {processed_idx}")
                    
                    img_path = os.path.join(temp_dir, f"slide_{processed_idx}.png")
                    
                    # 导出前切换到该幻灯片
                    presentation.Windows(1).View.GotoSlide(ppt_idx+1)
                    time.sleep(0.5)
                    
                    # 导出幻灯片
                    slide.Export(img_path, "PNG", 1920, 1080)
                    print(f"导出幻灯片 {processed_idx}/{len(slides_to_process)} 成功 (第 {ppt_idx+1} 页)")
                except Exception as e:
                    print(f"导出幻灯片 PPT索引 {ppt_idx} 失败: {e}")
                    print(traceback.format_exc())
                    # 尝试使用替代方法导出
                    try:
                        print(f"尝试使用替代方法导出幻灯片 {ppt_idx}...")
                        # 尝试使用PrintOut方法或其他屏幕截图方法
                        # 这里我们简单地生成一个空白图片作为替代
                        img = PILImage.new('RGB', (800, 600), color=(255, 255, 255))
                        d = PILImageDraw.Draw(img)
                        d.text((400, 300), f"幻灯片 {ppt_idx}", fill=(0, 0, 0))
                        img.save(img_path)
                        print(f"使用替代方法导出幻灯片 {ppt_idx} 成功")
                    except Exception as e2:
                        print(f"替代方法也失败: {e2}")
                        raise
            
            # Close PowerPoint
            print("关闭PowerPoint应用程序...")
            if presentation:
                presentation.Close()
            if ppt_app:
                ppt_app.Quit()
            presentation = None
            ppt_app = None
            print("PowerPoint应用程序已关闭")
            
            clips = []
            
            # Process each slide that needs to be included in the video
            print("开始生成音频...")
            # 创建一个映射表，记录处理序号到PPT索引的关系
            idx_to_ppt_map = {idx+1: ppt_idx for idx, ppt_idx in enumerate(slides_to_process)}
            for processed_idx in range(1, len(slides_to_process) + 1):
                ppt_idx = idx_to_ppt_map[processed_idx]
                img_path = os.path.join(temp_dir, f"slide_{processed_idx}.png")
                
                if not os.path.exists(img_path):
                    print(f"警告: 幻灯片图片不存在: {img_path}")
                    continue
                    
                # 使用PPT页码获取旁白文本
                text_to_speak = narrations.get(ppt_idx, f"这是第 {ppt_idx+1} 张幻灯片")
                subtitle_text = raw_narrations.get(ppt_idx, text_to_speak)
                
                print(f"处理幻灯片 {processed_idx}/{len(slides_to_process)} (对应PPT索引 {ppt_idx})")
                
                audio_path = os.path.join(temp_dir, f"audio_{processed_idx}.mp3")
                
                # 生成用于语音合成的文本，应用多音字替换规则
                tts_text = text_to_speak
                if pronunciation_dict:
                    # 应用所有替换规则
                    for original, replacement in pronunciation_dict.items():
                        tts_text = tts_text.replace(original, replacement)
                    
                    # 记录替换情况
                    if tts_text != text_to_speak:
                        print(f"已应用多音字优化，共替换 {sum(text_to_speak.count(orig) for orig in pronunciation_dict.keys())} 处")
                
                # Generate audio using the selected TTS engine
                audio_generated = False
                audio_clip = None
                
                # 先尝试使用指定的TTS引擎
                try:
                    if tts_engine.lower() == "xfyun" and xfyun_params:
                        # 使用科大讯飞语音合成
                        try:
                            # 获取发音人
                            voice_name = xfyun_params.get('voice_name', 'xiaoyan')
                            
                            # 添加停顿标记 - 根据不同发音人使用不同方式
                            text_with_pause = tts_text
                            cssml_voices = ['xiaoyan', 'xiaoyu', 'xiaofeng', 'xiaoqi', 'catherine', 'mary']
                            
                            # 检查是否需要使用CSSML方式添加停顿
                            if voice_name in cssml_voices:
                                # CSSML方式：添加<break>标签（开头和结尾）
                                text_with_pause = '<break time="500ms"/>' + text_with_pause + '<break time="600ms"/>'
                                # 设置ttp参数为cssml
                                xfyun_params['ttp'] = 'cssml'
                            else:
                                # 简单标记方式：添加[p500]和[p600]
                                text_with_pause =  '[p200]' + text_with_pause + '[p1000]'
                            
                            print(f"为旁白添加停顿标记，使用{'CSSML' if voice_name in cssml_voices else '简单'}标记方式")
                            print(f"每页结尾添加600ms的停顿")
                            
                            tts_result = xfyun_tts(
                                text_with_pause, 
                                audio_path, 
                                xfyun_params.get('app_id', ''), 
                                xfyun_params.get('api_key', ''),
                                xfyun_params.get('api_secret', ''),
                                voice=voice_name,
                                speed=xfyun_params.get('speed', 50),  # 使用可选的speed参数
                                ttp=xfyun_params.get('ttp', 'text')  # 传递ttp参数
                            )
                            if tts_result:
                                print(f"使用科大讯飞TTS生成音频成功")
                                audio_generated = True
                            else:
                                print(f"科大讯飞TTS失败，将尝试使用系统TTS...")
                        except Exception as xfyun_error:
                            print(f"科大讯飞TTS失败，错误: {xfyun_error}，将尝试使用系统TTS...")
                    
                    elif tts_engine.lower() == "ttsmaker" and ttsmaker_params:
                        # 使用马克配音TTS语音合成
                        try:
                            # 获取token和voice_id
                            token = ttsmaker_params.get('token', 'ttsmaker_demo_token')
                            voice_id = ttsmaker_params.get('voice_id', 1504)
                            audio_speed = ttsmaker_params.get('audio_speed', 1.0)  # 获取语速参数
                            
                            print(f"使用马克配音TTS生成音频，voice_id: {voice_id}, 语速: {audio_speed}")
                            
                            tts_result = ttsmaker_tts(
                                tts_text, 
                                audio_path, 
                                token=token,
                                voice_id=voice_id,
                                audio_speed=audio_speed  # 将语速参数传递给TTS函数
                            )
                            if tts_result:
                                print(f"使用马克配音TTS生成音频成功")
                                audio_generated = True
                            else:
                                print(f"马克配音TTS失败，将尝试使用系统TTS...")
                        except Exception as ttsmaker_error:
                            print(f"马克配音TTS失败，错误: {ttsmaker_error}，将尝试使用系统TTS...")
                    
                    # 如果其他TTS失败或者原本就选择了系统TTS
                    if not audio_generated or tts_engine.lower() == "pyttsx3":
                        engine = pyttsx3.init()
                        engine.save_to_file(tts_text, audio_path)
                        engine.runAndWait()
                        print(f"使用系统TTS生成音频成功")
                        audio_generated = True
                    
                    # 等待文件写入完成
                    time.sleep(1.0)
                    
                    if not os.path.exists(audio_path) or os.path.getsize(audio_path) == 0:
                        raise FileNotFoundError("音频文件未生成或为空")
                        
                except Exception as e:
                    print(f"生成音频时出错: {e}")
                    print(traceback.format_exc())
                    # 创建一个无声的音频文件作为备选
                    try:
                        print("创建备用无声音频文件...")
                        from moviepy.audio.AudioClip import AudioClip
                        silence = AudioClip(lambda t: 0, duration=5.0)  # 5秒静音
                        silence.write_audiofile(audio_path)
                        print("无声音频文件创建成功")
                        audio_generated = True
                    except Exception as silent_error:
                        print(f"创建无声音频失败: {silent_error}")
                
                if audio_generated:
                    try:
                        # Calculate duration based on audio length
                        audio_clip = AudioFileClip(audio_path)
                        audio_clips.append(audio_clip)  # Store for later cleanup
                        duration = audio_clip.duration
                        
                        print(f"音频时长: {duration}秒")
                        
                        # 检查是否为最后一页PPT，最后一页不添加字幕
                        is_last_slide = (processed_idx == len(slides_to_process))
                        
                        if is_last_slide:
                            # 最后一页PPT只播放声音，不显示字幕
                            print(f"处理最后一页PPT (索引 {ppt_idx})：只有语音，无字幕")
                            # 根据TTS引擎来决定是否需要剪裁音频
                            if tts_engine.lower() == "xfyun":
                                # 当使用科大讯飞TTS时，将原始音频剪去最后1秒
                                # audio_duration = audio_clip.duration
                                # if audio_duration > 1:
                                #     trimmed_audio_clip = audio_clip.subclip(0, audio_duration - 1)
                                #     print(f"使用科大讯飞TTS，剪裁最后1秒音频，从 {audio_duration}秒 减至 {trimmed_audio_clip.duration}秒")
                                # else:
                                #     trimmed_audio_clip = audio_clip  # 若音频长度小于等于1秒，则不剪裁
                                #     print(f"使用科大讯飞TTS，音频时长不足1秒，保持原长度 {audio_duration}秒")
                                
                                # # 图像时长也相应减少1秒
                                # img_clip = ImageClip(img_path).set_duration(trimmed_audio_clip.duration)
                                
                                # # 组合成视频片段
                                # video_clip = img_clip.set_audio(trimmed_audio_clip)

                                # 其他TTS引擎使用完整音频
                                print(f"使用 {tts_engine} TTS，保持完整音频时长 {audio_clip.duration}秒")
                                img_clip = ImageClip(img_path).set_duration(audio_clip.duration)
                                video_clip = img_clip.set_audio(audio_clip)
                            else:
                                # 其他TTS引擎使用完整音频
                                print(f"使用 {tts_engine} TTS，保持完整音频时长 {audio_clip.duration}秒")
                                img_clip = ImageClip(img_path).set_duration(audio_clip.duration)
                                video_clip = img_clip.set_audio(audio_clip)

                            clips.append(video_clip)
                        else:
                            # 非最后一页，正常添加字幕
                            # Use a simpler subtitle method with direct PIL drawing on the image
                            try:
                                # Load image and create a copy to draw on
                                base_image = PILImage.open(img_path)
                                
                                # Create a function to render text at the bottom of the image
                                def add_subtitles_to_image(image, text, font_size=28, bg_color=None, bg_color_name='白色半透明', font_color=(38, 74, 145)):
                                    # Make a copy to avoid modifying the original
                                    img_with_text = image.copy()
                                    draw = PILImageDraw.Draw(img_with_text)
                                    
                                    # Try to use a system font
                                    try:
                                        font = PILImageFont.truetype('simhei.ttf', font_size)
                                    except Exception:
                                        try:
                                            font = PILImageFont.truetype('arial.ttf', font_size)
                                        except Exception:
                                            font = PILImageFont.load_default()
                                    
                                    # Get text dimensions
                                    if hasattr(font, 'getbbox'):
                                        bbox = font.getbbox(text)
                                        text_width = bbox[2] - bbox[0]
                                        text_height = bbox[3] - bbox[1]
                                    else:
                                        text_width, text_height = font.getsize(text)
                                    
                                    # Position at bottom center
                                    x_pos = (image.width - text_width) // 2
                                    y_pos = image.height - text_height - 30
                                    
                                    # Draw background (仅当选择了背景颜色时)
                                    if bg_color is not None:  # 根据设置决定是否绘制背景
                                        bg_padding = 10
                                        bg_left = x_pos - bg_padding
                                        bg_right = x_pos + text_width + bg_padding
                                        bg_top = y_pos - 5
                                        bg_bottom = y_pos + text_height + 5
                                        draw.rectangle([bg_left, bg_top, bg_right, bg_bottom], fill=bg_color)
                                    
                                    # Draw text (使用指定的字体颜色)
                                    text_color = font_color
                                    
                                    # 根据背景调整文本颜色以确保足够的对比度
                                    if bg_color_name in ['黑色半透明', '蓝色半透明']:
                                        # 如果是深色背景但字体颜色也很深，则自动调整为白色
                                        r, g, b = font_color
                                        brightness = (0.299 * r + 0.587 * g + 0.114 * b) / 255
                                        if brightness < 0.5:  # 如果字体颜色较暗
                                            text_color = (255, 255, 255)  # 使用白色代替
                                            print(f"自动调整字体颜色为白色，以提高在深色背景上的可见度")
                                            
                                    draw.text((x_pos, y_pos), text, fill=text_color, font=font)
                                    
                                    return img_with_text
                                
                                # Split text into lines based on punctuation and limit line length
                                def split_into_lines(text, max_chars=28):  # 增加默认最大字符数
                                    # Split by punctuation first
                                    import re
                                    
                                    # 调试输出
                                    print(f"原始文本: {text}")
                                    
                                    # 第一步: 首先标记所有可能的产品型号，确保它们不被分割
                                    # 扩展产品型号匹配模式，确保一定能捕获"LP6286"这种格式
                                    alphanumeric_patterns = []
                                    
                                    # 使用多个特定的模式进行匹配
                                    # 1. 标准格式: 字母+数字组合(如LP6286)
                                    patterns = [
                                        # 常见模式: 字母开头后跟数字(可能包含连字符、下划线或点)
                                        r'[A-Za-z]+\d+(?:[A-Za-z0-9\-_\.]*)',
                                        # 特殊处理带数字中缀的型号(如TPS7A47)
                                        r'[A-Za-z]+\d+[A-Za-z]+\d+',
                                        # 特殊处理SOT23-6等封装格式
                                        r'SOT\d+-\d+',
                                        # 处理以数字开头的型号(如1.8V, 3.3V)
                                        r'\d+\.\d+[A-Za-z]+'
                                    ]
                                    
                                    # 先把文本复制一份用于处理
                                    working_text = text
                                    
                                    # 保存所有匹配到的模式及其位置
                                    all_matches = []
                                    
                                    # 用每个模式进行匹配
                                    for pattern in patterns:
                                        for match in re.finditer(pattern, working_text):
                                            matched_text = match.group()
                                            start_pos = match.start()
                                            end_pos = match.end()
                                            all_matches.append((matched_text, start_pos, end_pos))
                                            print(f"找到产品型号: '{matched_text}' 位置: {start_pos}-{end_pos}")
                                    
                                    # 按照起始位置降序排序，以避免替换时的位置偏移
                                    all_matches.sort(key=lambda x: x[1], reverse=True)

                                    # 在所有匹配结束后添加去重处理
                                    all_matches_unique = []
                                    seen_positions = set()
                                    for match in all_matches:
                                        matched_text, start_pos, end_pos = match
                                        position_key = (start_pos, end_pos)  # 创建位置键
                                        if position_key not in seen_positions:
                                            all_matches_unique.append(match)
                                            seen_positions.add(position_key)
                                            
                                    all_matches = all_matches_unique  # 使用去重后的结果
                                    
                                    # 替换所有匹配项为特殊标记
                                    protected_text = text
                                    replacement_map = {}  # 保存标记到原始文本的映射
                                    
                                    for i, (pattern, start, end) in enumerate(all_matches):
                                        marker = f"<PROTECTED{i}>"
                                        replacement_map[marker] = pattern
                                        protected_text = protected_text[:start] + marker + protected_text[end:]
                                    
                                    if all_matches:
                                        print(f"保护的模式: {[m[0] for m in all_matches]}")
                                        print(f"处理后的文本: {protected_text}")
                                    
                                    # 再保护小数点，避免在小数点处断句
                                    protected_text = re.sub(r'(\d+)\.(\d+)', r'\1<DECIMAL>\2', protected_text)
                                    
                                    # 按标点分割文本(使用句号、问号等)
                                    sentences = re.split(r'([。！？\.!?;；]+)', protected_text)
                                    
                                    # 重组句子和标点
                                    proper_sentences = []
                                    for i in range(0, len(sentences) - 1, 2):
                                        if i + 1 < len(sentences):
                                            proper_sentences.append(sentences[i] + sentences[i + 1])
                                        else:
                                            proper_sentences.append(sentences[i])
                                    
                                    # 处理最后一个元素
                                    if len(sentences) % 2 == 1:
                                        if proper_sentences:
                                            proper_sentences[-1] += sentences[-1]
                                        else:
                                            proper_sentences.append(sentences[-1])
                                    
                                    # 进一步分割长句子
                                    result = []
                                    for sentence in proper_sentences:
                                        # 检查句子长度，如果短则直接添加
                                        if len(sentence) <= max_chars:
                                            # 直接恢复所有保护标记并添加
                                            restored_sentence = restore_all_markers(sentence, replacement_map)
                                            result.append(restored_sentence)
                                        else:
                                            # 对长句进行再切分，但避免打断保护标记
                                            parts = split_with_marker_protection(sentence, max_chars, replacement_map)
                                            result.extend(parts)
                                    
                                    # 输出最终分行结果
                                    print(f"分行结果 ({len(result)}行):")
                                    for i, line in enumerate(result):
                                        print(f"  行{i+1}: {line}")
                                                                        
                                    return result
                                
                                # 恢复所有特殊标记为原始文本
                                def restore_all_markers(text, replacement_map):
                                    restored = text
                                    
                                    # 先恢复产品型号标记
                                    for marker, original in replacement_map.items():
                                        restored = restored.replace(marker, original)
                                    
                                    # 再恢复小数点
                                    restored = restored.replace('<DECIMAL>', '.')
                                    
                                    return restored
                                
                                # 安全地分割长文本，确保不会打断特殊标记
                                def split_with_marker_protection(text, max_chars, replacement_map):
                                    result_parts = []
                                    
                                    # 先检查是否有保护标记
                                    has_protected_markers = any(marker in text for marker in replacement_map.keys())
                                    
                                    # 如果有保护标记，使用更安全的分割方法
                                    if has_protected_markers:
                                        # 修改：将文本根据逗号分割，但保留逗号作为每段的结尾
                                        segments = []
                                        last_end = 0
                                        
                                        # 找到所有逗号位置并分段
                                        for i, char in enumerate(text):
                                            if char in ',，':
                                                segments.append(text[last_end:i+1])  # 包含逗号
                                                last_end = i+1
                                        
                                        # 添加最后一段(如果有)
                                        if last_end < len(text):
                                            segments.append(text[last_end:])
                                        
                                        current_line = ""
                                        for segment in segments:
                                            # 如果当前行加上这段不会太长
                                            if len(current_line + segment) <= max_chars:
                                                current_line += segment
                                            else:
                                                # 保存当前行并开始新行
                                                if current_line:
                                                    restored_line = restore_all_markers(current_line, replacement_map)
                                                    result_parts.append(restored_line)
                                                
                                                # 如果段落本身超过最大长度
                                                if len(segment) > max_chars:
                                                    # 检查是否包含保护标记
                                                    marker_found = False
                                                    start_idx = 0
                                                    
                                                    # 循环查找每个标记的位置
                                                    for marker in replacement_map.keys():
                                                        marker_pos = segment.find(marker)
                                                        if marker_pos != -1:
                                                            marker_found = True
                                                            # 如果标记在max_chars附近，特殊处理
                                                            if marker_pos < max_chars < marker_pos + len(marker):
                                                                # 拆分点在标记之前
                                                                restored_before = restore_all_markers(segment[:marker_pos], replacement_map)
                                                                if restored_before:
                                                                    result_parts.append(restored_before)
                                                                
                                                                # 标记和后面的内容放在下一行
                                                                remaining = segment[marker_pos:]
                                                                # 恢复保护标记
                                                                restored_remaining = restore_all_markers(remaining, replacement_map)
                                                                
                                                                # 如果剩余部分仍然超长，再次递归分割
                                                                if len(restored_remaining) > max_chars:
                                                                    sub_parts = split_with_marker_protection(
                                                                        restored_remaining, max_chars, {}  # 空映射，因为已经恢复了
                                                                    )
                                                                    result_parts.extend(sub_parts)
                                                                else:
                                                                    result_parts.append(restored_remaining)
                                                                
                                                                start_idx = len(segment)  # 标记已处理完整个部分
                                                                break
                                                    
                                                    # 如果没有找到标记或标记位置不需要特殊处理
                                                    if not marker_found or start_idx == 0:
                                                        # 使用常规分割方法
                                                        # 将该部分恢复所有标记后再分割
                                                        restored_part = restore_all_markers(segment, replacement_map)
                                                        
                                                        # 尝试在空格处分割
                                                        space_idx = restored_part.rfind(' ', 0, max_chars)
                                                        if space_idx > 0:
                                                            result_parts.append(restored_part[:space_idx+1])
                                                            
                                                            # 剩余部分如果超长，递归处理
                                                            if len(restored_part) - space_idx - 1 > max_chars:
                                                                sub_parts = split_with_marker_protection(
                                                                    restored_part[space_idx+1:], max_chars, {}
                                                                )
                                                                result_parts.extend(sub_parts)
                                                            else:
                                                                result_parts.append(restored_part[space_idx+1:])
                                                        else:
                                                            # 如果找不到空格，就强制分割，最大程度避开标点符号
                                                            segments = []
                                                            
                                                            # 查找所有可能的分割点(空格或标点)
                                                            potential_breaks = []
                                                            for j, char in enumerate(restored_part):
                                                                if char in ' ，,。.;；':
                                                                    potential_breaks.append(j)
                                                            
                                                            current_pos = 0
                                                            while current_pos < len(restored_part):
                                                                # 寻找最近的不超过max_chars的分割点
                                                                next_break = None
                                                                for pb in potential_breaks:
                                                                    if current_pos < pb < current_pos + max_chars:
                                                                        next_break = pb
                                                                
                                                                # 如果找到合适的分割点
                                                                if next_break is not None:
                                                                    segments.append(restored_part[current_pos:next_break+1])
                                                                    current_pos = next_break + 1
                                                                else:
                                                                    # 没有找到合适分割点，只能强制分割
                                                                    end_pos = min(current_pos + max_chars, len(restored_part))
                                                                    segments.append(restored_part[current_pos:end_pos])
                                                                    current_pos = end_pos
                                                            
                                                            result_parts.extend(segments)
                                                else:
                                                    # 段落长度合适，直接恢复标记并添加
                                                    restored_part = restore_all_markers(segment, replacement_map)
                                                    result_parts.append(restored_part)
                                                    
                                                # 重置当前行
                                                current_line = ""
                                        
                                        # 处理最后一行
                                        if current_line:
                                            restored_line = restore_all_markers(current_line, replacement_map)
                                            result_parts.append(restored_line)
                                    else:
                                        # 没有保护标记，使用简单的分割逻辑
                                        # 先恢复所有标记
                                        restored_text = restore_all_markers(text, replacement_map)
                                        
                                        # 使用修改后的辅助函数
                                        result_parts = split_simple_text(restored_text, max_chars)
                                    
                                    return result_parts
                                
                                # 简单文本分割，不考虑保护标记
                                def split_simple_text(text, max_chars):
                                    parts = []
                                    
                                    # 修改：将文本根据逗号分割，但保留逗号作为每段的结尾
                                    segments = []
                                    last_end = 0
                                    
                                    # 找到所有逗号位置并分段
                                    for i, char in enumerate(text):
                                        if char in ',，':
                                            segments.append(text[last_end:i+1])  # 包含逗号
                                            last_end = i+1
                                    
                                    # 添加最后一段(如果有)
                                    if last_end < len(text):
                                        segments.append(text[last_end:])
                                    
                                    current_line = ""
                                    for segment in segments:
                                        if len(current_line + segment) <= max_chars:
                                            current_line += segment
                                        else:
                                            if current_line:
                                                parts.append(current_line)
                                            
                                            # 如果段落本身超过最大长度
                                            if len(segment) > max_chars:
                                                # 寻找合适的断句点
                                                for j in range(0, len(segment), max_chars):
                                                    end_idx = min(j + max_chars, len(segment))
                                                    chunk = segment[j:end_idx]
                                                    parts.append(chunk)
                                                current_line = ""
                                            else:
                                                current_line = segment
                                    
                                    # 添加最后一行
                                    if current_line:
                                        parts.append(current_line)
                                    
                                    return parts
                                
                                # 新增函数：移除字符串末尾的标点符号
                                def remove_ending_punctuation(text):
                                    # 定义中英文标点符号列表
                                    punctuation_marks = ['.', '。', ',', '，', '?', '？', '!', '！', ';', '；', ':', '：']
                                    
                                    # 检查字符串是否以标点符号结尾
                                    if text and text[-1] in punctuation_marks:
                                        return text[:-1]  # 移除最后一个字符
                                    return text
                                
                                # 修改：不再使用整个音频时长估计显示时间，而是逐行生成音频并测量实际时长
                                lines = split_into_lines(subtitle_text)
                                if not lines:
                                    lines = [""]  # Ensure we have at least one line
                                    
                                print(f"该页分成 {len(lines)} 行字幕")
                                
                                # 存储每行的音频片段和对应时长
                                line_audio_clips = []
                                line_durations = []
                                total_line_duration = 0.0
                                
                                # 根据精准字幕设置选择处理方式
                                use_precise_subtitles = subtitle_params.get('precise_subtitle', False)
                                print(f"字幕模式: {'精准字幕' if use_precise_subtitles else '估算字幕'}")
                                
                                if use_precise_subtitles:
                                    # 精准字幕模式: 为每行单独生成音频，并获取实际时长
                                    print("使用精准字幕模式，开始逐行生成音频...")
                                    
                                    for i, line in enumerate(lines):
                                        line_text = line.strip()
                                        if not line_text:
                                            continue
                                            
                                        line_audio_path = os.path.join(temp_dir, f"line_audio_{processed_idx}_{i}.mp3")
                                        
                                        # 为科大讯飞TTS添加特殊处理（首句句首和尾句句尾添加停顿）
                                        line_tts_text = line_text
                                        if tts_engine.lower() == "xfyun":
                                            # 为第一行添加句首停顿
                                            if i == 0:
                                                if xfyun_params.get('ttp') == 'cssml':
                                                    line_tts_text = '<break time="500ms"/>' + line_tts_text
                                                else:
                                                    line_tts_text = line_tts_text                                        
                                            # 为最后一行添加句尾停顿
                                            elif i == len(lines) - 1:
                                                if xfyun_params.get('ttp') == 'cssml':
                                                    line_tts_text = line_tts_text + '<break time="600ms"/>'
                                                else:
                                                    line_tts_text = line_tts_text + '[p800]'
                                            else:            
                                                if xfyun_params.get('ttp') == 'cssml':
                                                    line_tts_text = line_tts_text + '<break time="100ms"/>'
                                                else:
                                                    line_tts_text = line_tts_text

                                        # 应用多音字替换规则
                                        if pronunciation_dict:
                                            for original, replacement in pronunciation_dict.items():
                                                line_tts_text = line_tts_text.replace(original, replacement)
                                        
                                        # 生成该行的音频
                                        line_audio_generated = False
                                        
                                        if tts_engine.lower() == "xfyun" and xfyun_params:
                                            # 使用科大讯飞
                                            tts_result = xfyun_tts(
                                                line_tts_text, 
                                                line_audio_path, 
                                                xfyun_params.get('app_id', ''), 
                                                xfyun_params.get('api_key', ''),
                                                xfyun_params.get('api_secret', ''),
                                                voice=xfyun_params.get('voice_name', 'xiaoyan'),
                                                speed=xfyun_params.get('speed', 50),
                                                ttp=xfyun_params.get('ttp', 'text')
                                            )
                                            line_audio_generated = tts_result
                                        elif tts_engine.lower() == "ttsmaker" and ttsmaker_params:
                                            # 使用马克配音
                                            tts_result = ttsmaker_tts(
                                                line_tts_text, 
                                                line_audio_path, 
                                                token=ttsmaker_params.get('token', 'ttsmaker_demo_token'),
                                                voice_id=ttsmaker_params.get('voice_id', 1504),
                                                audio_speed=ttsmaker_params.get('audio_speed', 1.0)
                                            )
                                            line_audio_generated = tts_result
                                        else:
                                            # 使用系统TTS
                                            engine = pyttsx3.init()
                                            engine.save_to_file(line_tts_text, line_audio_path)
                                            engine.runAndWait()
                                            line_audio_generated = True
                                            time.sleep(0.5)  # 等待文件写入
                                            
                                        # 检查音频是否生成成功
                                        if line_audio_generated and os.path.exists(line_audio_path) and os.path.getsize(line_audio_path) > 0:
                                            # 读取音频并获取时长
                                            line_audio_clip = AudioFileClip(line_audio_path)
                                            line_audio_clips.append(line_audio_clip)
                                            line_duration = line_audio_clip.duration
                                            line_durations.append(line_duration)
                                            total_line_duration += line_duration
                                            print(f"  行 {i+1}: '{line_text[:30]}{'...' if len(line_text) > 30 else ''}' - 音频时长: {line_duration:.2f}秒")
                                        else:
                                            # 音频生成失败时添加默认时长
                                            print(f"  行 {i+1}: 音频生成失败，使用默认时长")
                                            line_durations.append(2.0)  # 默认2秒
                                            total_line_duration += 2.0
                                else:
                                    # 估算字幕模式: 使用estimate_line_duration函数估算每行时长
                                    print("使用估算字幕模式，计算每行估计时长...")
                                    
                                    # 为每行估算时长
                                    for i, line in enumerate(lines):
                                        line_text = line.strip()
                                        if not line_text:
                                            # 空行使用最小时长
                                            estimated_duration = 0.5  # 0.5秒为空行
                                            line_durations.append(estimated_duration)
                                            total_line_duration += estimated_duration
                                            print(f"  行 {i+1}: [空行] - 估计时长: {estimated_duration:.2f}秒")
                                            continue
                                            
                                        # 使用estimate_line_duration函数估算时长
                                        estimated_duration = estimate_line_duration(line_text)
                                        # 应用衰减系数，每行递增4%，但最多增加20%
                                        decay_factor = min(1.2, 1.0 + (i * 0.04))
                                        estimated_duration *= decay_factor
                                        line_durations.append(estimated_duration)
                                        total_line_duration += estimated_duration
                                        
                                        print(f"  行 {i+1}: '{line_text[:30]}{'...' if len(line_text) > 30 else ''}' - 估计时长: {estimated_duration:.2f}秒")
                                
                                # 连接所有音频片段
                                print("使用整体音频，总时长:", audio_clip.duration, "秒")
                                
                                # 首先获取整页音频的总时长，作为参考
                                actual_audio_duration = audio_clip.duration
                                
                                # 根据各行时长占比，计算实际时长
                                real_line_durations = []
                                
                                if total_line_duration > 0:
                                    # 按比例计算每行实际时长
                                    for i, duration in enumerate(line_durations):
                                        proportion = duration / total_line_duration
                                        real_duration = proportion * actual_audio_duration
                                        real_line_durations.append(real_duration)
                                        print(f"  行 {i+1} 占比: {proportion:.2%}, 实际时长: {real_duration:.2f}秒")
                                else:
                                    # 如果总估计时长为0，平均分配时间
                                    equal_duration = actual_audio_duration / max(len(lines), 1)
                                    real_line_durations = [equal_duration] * len(lines)
                                    print(f"  无法计算时长占比，每行平均分配: {equal_duration:.2f}秒")
                                
                                # Calculate frame counts for each line
                                fps = 24   #字幕略微提前
                                frame_counts = [int(d * fps) for d in real_line_durations]
                                total_frames = sum(frame_counts)
                                
                                # Create frames directory
                                frames_dir = os.path.join(temp_dir, f"frames_{processed_idx}")
                                os.makedirs(frames_dir, exist_ok=True)
                                
                                # Generate frames for each line
                                frame_files = []
                                frame_idx = 0
                                
                                print(f"生成字幕帧，共 {len(lines)} 行文本")
                                for i, (line, frames_for_line) in enumerate(zip(lines, frame_counts)):
                                    line_text = line.strip()
                                    
                                    # 去除行尾标点符号
                                    line_text = remove_ending_punctuation(line_text)
                                    
                                    if not line_text:
                                        # For empty lines, just use base image
                                        frame_img = base_image.copy()
                                    else:
                                        frame_img = add_subtitles_to_image(
                                            base_image, 
                                            line_text,
                                            font_size=font_size,     # 传递字体大小
                                            bg_color=bg_color,       # 传递背景颜色
                                            bg_color_name=bg_color_name,  # 传递背景颜色名称
                                            font_color=font_color    # 传递字体颜色
                                        )
                                    
                                    # Save each frame for this line
                                    for _ in range(frames_for_line):
                                        frame_path = os.path.join(frames_dir, f"frame_{frame_idx:04d}.png")
                                        frame_img.save(frame_path)
                                        frame_files.append(frame_path)
                                        frame_idx += 1
                                    
                                    print(f"  第 {i+1}/{len(lines)} 行: {line_text[:20]}{'...' if len(line_text) > 20 else ''} - {frames_for_line} 帧")
                                
                                # If we have fewer frames than needed, repeat the last frame
                                while len(frame_files) < total_frames:
                                    frame_path = os.path.join(frames_dir, f"frame_{frame_idx:04d}.png")
                                    if frame_files:
                                        # Copy the last frame
                                        shutil.copy(frame_files[-1], frame_path)
                                    else:
                                        # Use base image if no frames were created
                                        base_image.save(frame_path)
                                    frame_files.append(frame_path)
                                    frame_idx += 1
                                
                                # Create a clip from image sequence
                                video_clip = ImageSequenceClip(frame_files, fps=fps)
                                
                                # Add audio
                                video_clip = video_clip.set_audio(audio_clip)
                                print("字幕已添加到视频，按行显示")
                                
                            except Exception as subtitle_error:
                                print(f"添加字幕时出错: {subtitle_error}")
                                print(traceback.format_exc())
                                # Fallback to no subtitles
                                img_clip = ImageClip(img_path).set_duration(duration)
                                video_clip = img_clip.set_audio(audio_clip)
                            
                            # 获取原始时长
                            original_duration = video_clip.duration
                            audio_duration = audio_clip.duration

                            # 确保视频和音频足够长才进行裁剪
                            if original_duration > 0.9 and audio_duration > 0.9:
                                # 同时裁剪视频和音频
                                video_clip = video_clip.subclip(0, original_duration - 0.9)
                                audio_clip = audio_clip.subclip(0, audio_duration - 0.9)
                                print(f"视频和音频时长已缩短0.9秒，当前时长: {video_clip.duration:.2f}秒")

                            # 然后设置音频
                            video_clip = video_clip.set_audio(audio_clip)
                            clips.append(video_clip)
                    except Exception as clip_error:
                        print(f"处理视频剪辑时出错: {clip_error}")
                        print(traceback.format_exc())
                        
                        # Fallback to simpler approach without subtitles
                        try:
                            img_clip = ImageClip(img_path).set_duration(duration)
                            video_clip = img_clip.set_audio(audio_clip)
                            clips.append(video_clip)
                        except Exception as fallback_error:
                            print(f"使用备选方案时出错: {fallback_error}")
            
            # Concatenate all clips
            if clips:
                print(f"合成 {len(clips)} 个片段为最终视频...")
                
                # 移除交叉淡入淡出效果，因为它会导致黑屏问题
                # 直接拼接视频片段，依靠音频中的静音缓冲区避免噪音
                final_clip = concatenate_videoclips(clips, method="compose")  # 仅使用compose方法而非crossfade
                
                # 添加整体淡入淡出效果
                if final_clip.duration > 2.0:
                    final_clip = final_clip.fadein(0.0).fadeout(0.5)
                
                # 多线程导出视频以提高效率
                try:
                    # 获取CPU核心数，用于设置线程数
                    import multiprocessing
                    cpu_count = multiprocessing.cpu_count()
                    # 使用CPU核心数的75%作为线程数，最少2个线程，最多12个线程
                    threads = max(2, min(12, int(cpu_count * 0.75)))
                    print(f"检测到 {cpu_count} 个CPU核心，将使用 {threads} 个线程进行视频导出")
                    
                    # 使用更高效的FFMPEG参数
                    ffmpeg_params = [
                        "-preset", "medium",  # 中等压缩率，平衡速度和质量
                        "-crf", "21",         # 恒定质量因子，较好的视频质量，数值越小质量越高
                        "-movflags", "+faststart",  # 优化文件结构以便快速开始播放
                        "-profile:v", "high", # 使用高配置文件提高兼容性
                        "-tune", "stillimage" # 针对幻灯片优化
                    ]
                    
                    # 提高音频质量设置
                    print(f"开始多线程导出视频到: {output_video_path}")
                    final_clip.write_videofile(
                        output_video_path, 
                        fps=24, 
                        codec='libx264', 
                        audio_codec='aac', 
                        audio_bitrate='192k',
                        threads=threads,  # 使用多线程
                        ffmpeg_params=ffmpeg_params,  # 使用优化的FFMPEG参数
                        logger=None  # 不使用进度条，减少控制台输出
                    )
                    print(f"视频导出成功: {output_video_path}")
                except Exception as e:
                    print(f"多线程导出失败 ({e})，将使用单线程导出...")
                    # 单线程备选方案
                    final_clip.write_videofile(
                        output_video_path, 
                        fps=24, 
                        codec='libx264', 
                        audio_codec='aac', 
                        audio_bitrate='192k'
                    )
                
                print(f"视频创建成功: {output_video_path}")
            else:
                # 如果clips列表为空，提供更明确的错误信息
                print(f"错误: 没有成功创建任何视频片段。以下是处理过的幻灯片索引: {slides_to_process}")
                raise ValueError("没有处理任何幻灯片。请检查是否正确提取了旁白文字和导出了幻灯片图片。")
        finally:
            # Clean up resources before removing temp directory
            print("清理资源...")
            # Close all audio clips
            for clip in audio_clips:
                try:
                    clip.close()
                except:
                    pass
            
            # Clean up temporary directory manually
            try:
                # Wait a bit to ensure files are released
                time.sleep(2)
                
                # Remove temporary files one by one with error handling
                for root, dirs, files in os.walk(temp_dir, topdown=False):
                    for file in files:
                        try:
                            file_path = os.path.join(root, file)
                            if os.path.exists(file_path):
                                os.remove(file_path)
                        except Exception as e:
                            print(f"无法删除临时文件 {file}: {e}")
                    
                    for dir in dirs:
                        try:
                            dir_path = os.path.join(root, dir)
                            if os.path.exists(dir_path):
                                os.rmdir(dir_path)
                        except Exception as e:
                            print(f"无法删除临时目录 {dir}: {e}")
                
                # Try to remove the root temp directory
                try:
                    if os.path.exists(temp_dir):
                        os.rmdir(temp_dir)
                except Exception as e:
                    print(f"无法删除临时根目录: {e}")
            except Exception as cleanup_error:
                print(f"清理临时文件时出错: {cleanup_error}")
                print("继续执行，临时文件可能需要手动清理")
                
    except Exception as e:
        print(f"转换过程中出错: {e}")
        print(traceback.format_exc())
        # Ensure PowerPoint is closed even if there's an error
        try:
            if presentation:
                presentation.Close()
            if ppt_app:
                ppt_app.Quit()
        except Exception as cleanup_error:
            print(f"关闭PowerPoint时出错: {cleanup_error}")
        raise Exception(f"转换失败: {str(e)}")

def xfyun_tts(text, output_file, app_id, api_key, api_secret, voice="xiaoyan", speed=50, volume=50, ttp="text"):
    """
    使用科大讯飞云平台的TTS API生成语音（基于官方WebSocket示例）
    
    Parameters:
    - text: 要转换为语音的文本
    - output_file: 输出的音频文件路径
    - app_id: 科大讯飞应用ID
    - api_key: API密钥
    - api_secret: API密钥密码
    - voice: 发音人，默认为 xiaoyan
    - speed: 语速，范围是0~100，默认为50
    - volume: 音量，范围是0~100，默认为50
    - ttp: 文本类型，"text"为普通文本，"cssml"为CSSML文本
    
    Returns:
    - 成功返回True，失败返回False
    """
    print(f"使用科大讯飞WebSocket API生成语音...")
    print(f"APPID: {app_id}, 发音人: {voice}, 语速: {speed}, 文本类型: {ttp}")
    print(f"文本长度: {len(text)} 字符")
    
    # 删除之前可能存在的临时PCM文件
    temp_pcm = os.path.join(tempfile.gettempdir(), "xfyun_temp.pcm")
    if os.path.exists(temp_pcm):
        try:
            os.remove(temp_pcm)
        except:
            pass
    
    # 保存运行状态和结果
    tts_success = False
    tts_done = False
    tts_error = None
    audio_received = False  # 新增: 跟踪是否接收到音频数据
    
    # 从wsgiref.handlers导入时间格式化函数
    from wsgiref.handlers import format_date_time
    from time import mktime
    
    class XFyunWsParam:
        """科大讯飞WebSocket参数类"""
        def __init__(self, app_id, api_key, api_secret, text):
            self.APPID = app_id
            self.APIKey = api_key
            self.APISecret = api_secret
            self.Text = text
            
            # 公共参数
            self.CommonArgs = {"app_id": self.APPID}
            
            # 业务参数 - 按照官方格式设置
            self.BusinessArgs = {
                "aue": "raw",  # 使用raw格式，便于后续处理
                "auf": "audio/L16;rate=16000",  # 音频采样率
                "vcn": voice,  # 发音人
                "tte": "utf8",  # 文本编码
                "speed": speed,  # 语速
                "volume": volume,  # 音量
                "ttp": ttp  # 文本类型，普通文本或CSSML
            }
            
            # 数据
            self.Data = {
                "status": 2,  # 2表示一次性发送全部数据
                "text": str(base64.b64encode(self.Text.encode('utf-8')), "UTF8")
            }
        
        def create_url(self):
            """生成鉴权URL"""
            url = 'wss://tts-api.xfyun.cn/v2/tts'
            
            # 生成RFC1123格式的时间戳
            now = datetime.datetime.now()
            date = format_date_time(mktime(now.timetuple()))
            
            # 拼接字符串
            signature_origin = "host: " + "ws-api.xfyun.cn" + "\n"
            signature_origin += "date: " + date + "\n"
            signature_origin += "GET " + "/v2/tts " + "HTTP/1.1"
            
            print(f"签名原始字符串: \n{signature_origin}")
            
            # 进行hmac-sha256加密
            signature_sha = hmac.new(self.APISecret.encode('utf-8'), 
                                      signature_origin.encode('utf-8'),
                                      digestmod=hashlib.sha256).digest()
            signature_sha_base64 = base64.b64encode(signature_sha).decode(encoding='utf-8')
            
            authorization_origin = f'api_key="{self.APIKey}", algorithm="hmac-sha256", headers="host date request-line", signature="{signature_sha_base64}"'
            authorization = base64.b64encode(authorization_origin.encode('utf-8')).decode(encoding='utf-8')
            
            # 将请求的鉴权参数组合为字典
            v = {
                "authorization": authorization,
                "date": date,
                "host": "ws-api.xfyun.cn"
            }
            
            # 拼接鉴权参数，生成url
            url = url + '?' + urlencode(v)
            print(f"WebSocket URL: {url.split('?')[0]}?authorization=***&date={date}&host=ws-api.xfyun.cn")
            return url
    
    # WebSocket回调函数
    def on_message(ws, message):
        nonlocal tts_success, tts_done, tts_error, audio_received
        try:
            message = json.loads(message)
            code = message["code"]
            sid = message["sid"]
            
            if code != 0:
                error_msg = message.get("message", "未知错误")
                print(f"错误信息: {error_msg}, 代码: {code}")
                tts_error = f"科大讯飞错误: {error_msg} (代码: {code})"
                tts_done = True
                return
            
            # 只有成功的响应才会有data字段
            if "data" in message:
                audio = message["data"].get("audio", "")
                status = message["data"].get("status", 0)
                
                if audio:
                    audio_received = True
                    # 解码音频数据并保存
                    audio_data = base64.b64decode(audio)
                    with open(temp_pcm, 'ab') as f:
                        f.write(audio_data)
                    
                    print(f"已接收音频数据片段，当前状态: {status}")
                
                if status == 2:  # 所有数据接收完成
                    print("WebSocket数据传输完成")
                    tts_success = audio_received  # 只有在接收到音频数据时才算成功
                    tts_done = True
                    ws.close()
        except Exception as e:
            print(f"处理WebSocket消息时出错: {e}")
            print(traceback.format_exc())
            tts_error = str(e)
            tts_done = True
    
    def on_error(ws, error):
        nonlocal tts_done, tts_error
        print(f"WebSocket错误: {error}")
        tts_error = f"WebSocket错误: {error}"
        tts_done = True
    
    def on_close(ws, close_status_code, close_msg):
        nonlocal tts_done
        print(f"WebSocket连接关闭: 状态码={close_status_code}, 消息={close_msg}")
        if not tts_done:  # 如果不是正常完成导致的关闭
            tts_done = True
            # 如果连接提前关闭但已接收部分数据，可能也是成功的
            if audio_received and not tts_success:
                tts_success = True  # 添加缺失的代码块 - 将连接标记为成功
    
    def on_open(ws):
        print("WebSocket连接已建立，发送数据...")
        
        def run(*args):
            try:
                # 组装消息数据
                data = {
                    "common": ws_param.CommonArgs,
                    "business": ws_param.BusinessArgs,
                    "data": ws_param.Data
                }
                
                # 发送JSON数据
                ws.send(json.dumps(data))
                print("文本数据已发送，等待响应...")
            except Exception as e:
                print(f"发送数据时出错: {e}")
                print(traceback.format_exc())
                nonlocal tts_error, tts_done
                tts_error = f"发送数据时出错: {e}"
                tts_done = True
                ws.close()
        
        # 启动线程发送数据
        thread = threading.Thread(target=run)
        thread.daemon = True
        thread.start()

    # 主要处理逻辑 - 使用新的try块
    try:
        # 初始化WebSocket参数
        ws_param = XFyunWsParam(app_id, api_key, api_secret, text)
        ws_url = ws_param.create_url()
        
        # 启用跟踪以便调试
        websocket.enableTrace(False)
        
        # 创建WebSocket连接
        ws = websocket.WebSocketApp(
            ws_url, 
            on_message=on_message, 
            on_error=on_error, 
            on_close=on_close
        )
        ws.on_open = on_open
        
        # 启动WebSocket客户端（非阻塞模式）
        ws_thread = threading.Thread(
            target=ws.run_forever, 
            kwargs={
                "sslopt": {"cert_reqs": ssl.CERT_NONE},
                "ping_interval": 5,  # 发送ping的间隔(秒)
                "ping_timeout": 3    # ping超时(秒)
            }
        )
        ws_thread.daemon = True
        ws_thread.start()
        
        # 等待处理完成，最多等待30秒
        timeout = 30
        start_time = time.time()
        print(f"等待WebSocket响应，超时时间: {timeout}秒...")
        while not tts_done and (time.time() - start_time < timeout):
            time.sleep(0.1)
        
        # 如果超时但未完成，则标记为错误
        if not tts_done:
            print(f"WebSocket请求超时（{timeout}秒）")
            tts_error = f"请求超时（{timeout}秒）"
            ws.close()
        
        # 检查是否成功，并进行后续处理
        if tts_success and os.path.exists(temp_pcm) and os.path.getsize(temp_pcm) > 0:
            print(f"语音合成成功，PCM大小: {os.path.getsize(temp_pcm)} 字节")
            
            # 根据输出格式进行处理
            if output_file.lower().endswith('.mp3'):
                # PCM -> MP3 转换
                print("转换PCM到MP3格式...")
                
                # 将PCM数据转换为WAV格式
                print("创建临时WAV文件...")
                temp_wav = temp_pcm + ".wav"
                with wave.open(temp_wav, 'wb') as wf:
                    wf.setnchannels(1)  # 单声道
                    wf.setsampwidth(2)  # 16位
                    wf.setframerate(16000)  # 16kHz
                    with open(temp_pcm, 'rb') as pcm_file:
                        wf.writeframes(pcm_file.read())
                
                # 将WAV转换为MP3
                print("将WAV转换为MP3...")
                audio_clip = AudioFileClip(temp_wav)
                audio_clip.write_audiofile(output_file, bitrate="192k", fps=16000, verbose=False, logger=None)
                
                # 清理临时文件
                try:
                    if os.path.exists(temp_wav):
                        os.remove(temp_wav)
                    if os.path.exists(temp_pcm):
                        os.remove(temp_pcm)
                except Exception as e:
                    print(f"清理临时文件失败: {e}")
            else:
                # 直接复制PCM文件到目标路径
                import shutil
                shutil.copy(temp_pcm, output_file)
                # 删除临时PCM文件
                try:
                    if os.path.exists(temp_pcm):
                        os.remove(temp_pcm)
                except:
                    pass
            
            print(f"音频文件已生成: {output_file}")
            return True
        else:
            if tts_error:
                print(f"语音合成失败: {tts_error}")
            else:
                print("语音合成失败，未生成有效音频数据")
                
            return False
    except Exception as e:
        print(f"语音合成过程中出错: {e}")
        print(traceback.format_exc())
        return False

def ttsmaker_tts(text, output_file, token="ttsmaker_demo_token", voice_id=1504, audio_format='mp3', audio_speed=1.0, audio_volume=0):
    """
    使用马克配音API生成语音
    
    Parameters:
    - text: 要转换为语音的文本
    - output_file: 输出的音频文件路径
    - token: API token，默认使用demo token
    - voice_id: 语音ID，默认为1504
    - audio_format: 音频格式，默认为mp3
    - audio_speed: 语音速度，范围0.5-2.0，默认1.0
    - audio_volume: 音量增益，范围0-10，默认0
    
    Returns:
    - 成功返回True，失败返回False
    """
    print(f"使用马克配音API生成语音...")
    print(f"Token: {'*'*len(token)}, Voice ID: {voice_id}")
    print(f"文本长度: {len(text)} 字符")
    
    # API URL
    url = 'https://api.ttsmaker.cn/v1/create-tts-order'
    
    # 请求头
    headers = {'Content-Type': 'application/json; charset=utf-8'}
    
    # 请求参数
    params = {
        'token': token,
        'text': text,
        'voice_id': int(voice_id),
        'audio_format': audio_format,
        'audio_speed': float(audio_speed),  # 确保使用浮点数
        'audio_volume': int(audio_volume),
        'text_paragraph_pause_time': 0
    }
    
    try:
        print("发送TTS请求...")
        response = requests.post(url, headers=headers, data=json.dumps(params))
        
        # 检查响应
        if response.status_code != 200:
            print(f"API请求失败: HTTP状态码 {response.status_code}")
            print(f"响应内容: {response.text}")
            return False
        
        # 解析JSON响应
        response_data = response.json()
        
        # 检查API响应状态
        if response_data.get('status') != 'success':
            error_code = response_data.get('error_code', '未知')
            error_details = response_data.get('error_details', '未知错误')
            print(f"API返回错误: {error_code} - {error_details}")
            return False
        
        # 获取音频文件URL
        audio_url = response_data.get('audio_file_url')
        if not audio_url:
            print("API响应中未找到音频URL")
            return False
            
        # 显示API使用情况
        if 'token_status' in response_data:
            token_status = response_data['token_status']
            used = token_status.get('current_cycle_characters_used', 0)
            available = token_status.get('current_cycle_characters_available', 0)
            print(f"Token使用情况: 已使用 {used} 字符, 可用 {available} 字符")
        
        # 下载音频文件到临时文件
        print(f"下载音频文件: {audio_url}")
        
        # 添加SSL处理和重试逻辑
        retry_count = 0
        max_retries = 3
        backoff_factor = 0.5  # 每次重试的延迟增加
        success = False
        
        while retry_count < max_retries and not success:
            try:
                # 添加超时和禁用SSL验证选项以处理SSL错误
                session = requests.Session()
                session.mount('https://', requests.adapters.HTTPAdapter(max_retries=2))
                
                audio_response = session.get(
                    audio_url, 
                    timeout=30,
                    verify=False  # 注意：在生产环境中不推荐禁用SSL验证
                )
                
                if audio_response.status_code == 200:
                    success = True
                else:
                    print(f"下载失败 (HTTP {audio_response.status_code})，重试 {retry_count+1}/{max_retries}...")
                    retry_count += 1
                    time.sleep(backoff_factor * (2 ** retry_count))  # 指数回退
            except Exception as e:
                print(f"下载出错: {e}，重试 {retry_count+1}/{max_retries}...")
                retry_count += 1
                time.sleep(backoff_factor * (2 ** retry_count))
        
        if not success:
            print("所有下载尝试都失败")
            return False
        
        # 创建临时文件用于保存原始音频
        temp_audio_file = output_file + ".temp." + audio_format
        with open(temp_audio_file, 'wb') as f:
            f.write(audio_response.content)
            
        print(f"原始音频文件已保存到临时文件: {temp_audio_file}")
        
        # 使用moviepy为音频添加前后静音
        try:
            print("为音频添加前后0.5秒静音...")
            from moviepy.editor import AudioFileClip, concatenate_audioclips, AudioClip
            
            # 加载原始音频
            original_audio = AudioFileClip(temp_audio_file)
            
            # 创建0.5秒静音片段
            #silence_duration = 0.3  # 0.7秒静音
            silence_clip1 = AudioClip(lambda t: 0, duration=0.6)
            silence_clip2 = AudioClip(lambda t: 0, duration=0.6)
            
            # 拼接：静音 + 原始音频 + 静音
            #final_audio = concatenate_audioclips([silence_clip1, original_audio, silence_clip2])
            final_audio = concatenate_audioclips([ original_audio, silence_clip2])
            
            # 保存到目标文件
            final_audio.write_audiofile(output_file, verbose=False, logger=None)
            
            # 关闭音频剪辑以释放资源
            original_audio.close()
            final_audio.close()
            
            # 删除临时文件
            try:
                if os.path.exists(temp_audio_file):
                    os.remove(temp_audio_file)
            except Exception as e:
                print(f"删除临时文件失败: {e}")
                
            print(f"已添加前后静音处理，音频文件已保存到: {output_file}")
            return True
        except Exception as e:
            print(f"添加静音处理失败: {e}")
            print(traceback.format_exc())
            
            # 如果添加静音失败，使用原始音频
            print("使用原始音频（无静音处理）...")
            try:
                # 将临时文件重命名为最终输出文件
                if os.path.exists(output_file):
                    os.remove(output_file)
                os.rename(temp_audio_file, output_file)
                print(f"音频文件已保存到: {output_file} (无静音处理)")
                return True
            except Exception as e2:
                print(f"保存原始音频失败: {e2}")
                return False
        
    except Exception as e:
        print(f"马克配音TTS调用过程中出现异常: {e}")
        print(traceback.format_exc())
        return False
