from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
import threading
import traceback
import pyautogui
import cv2
import numpy as np
from PIL import ImageGrab, Image, ImageTk
import re
import os
import sys
import json

# 导入配置模块和AI评分模块
import config
from ai_scoring import DeepSeekScorer

# 添加OCR支持
try:
    import pytesseract
    from PIL import Image as PILImage
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# 全局变量控制运行状态
running = False

# 创建图片存储目录
IMAGE_DIR = "image"
if not os.path.exists(IMAGE_DIR):
    os.makedirs(IMAGE_DIR)

# 用于存储两个区域的坐标
score_area_rect = None       # 分数输入区域
submit_area_rect = None      # 提交按钮区域
answer_area_rect = None      # 答案文本区域

# 记录区域是否已选择
score_area_selected = False
submit_area_selected = False
answer_area_selected = False  # 添加答案区域选择状态

# 新增全局变量
score_input_rect = None    # 分数输入框区域
score_input_selected = False
final_submit_rect = None   # 最终提交按钮区域
final_submit_selected = False
scroll_amount = -3         # 滚动量，负值表示向下滚动
current_question = 1       # 当前题目序号

# 新增题号相关变量
question_number_rect = None    # 题号区域
question_number_selected = False

def toggle_grading():
    global running
    running = not running
    
    if running:
        # 检查是否已选择所有必要区域
        if not (score_area_selected and submit_area_selected):
            messagebox.showwarning("警告", "请先选择分数区域和提交按钮区域")
            running = False
            return
            
        # 如果启用了AI评分但没有选择答案区域
        if config.config.get("use_ai_scoring", False) and not answer_area_selected:
            messagebox.showwarning("警告", "已启用AI评分，但未选择答案区域")
            running = False
            return
            
        btn_toggle.config(text="暂停", bg="red")
        status_label.config(text="状态：运行中")
        log_message("开始自动批改")
        # 在新线程中启动批改，避免界面卡顿
        threading.Thread(target=start_grading, daemon=True).start()
    else:
        btn_toggle.config(text="开始", bg="green")
        status_label.config(text="状态：已暂停")
        log_message("批改已暂停")

def log_message(message):
    """向日志窗口添加消息"""
    log_area.config(state="normal")
    log_area.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
    log_area.see(tk.END)
    log_area.config(state="disabled")

# 创建新的选择区域方法 - 使用鼠标拖动来选择
def capture_screen_region(window_title="选择区域"):
    global root
    
    # 最小化主窗口
    root.iconify()
    time.sleep(1)
    
    try:
        # 截取整个屏幕
        screen_img = pyautogui.screenshot()
        screen_np = np.array(screen_img)
        screen_np = cv2.cvtColor(screen_np, cv2.COLOR_RGB2BGR)
        
        # 创建一个新窗口显示截图
        cv2.namedWindow(window_title, cv2.WINDOW_NORMAL)
        cv2.setWindowProperty(window_title, cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
        
        # 全局变量记录选择的区域
        selected_rect = [0, 0, 0, 0]
        selecting = False
        
        # 鼠标回调函数
        def mouse_callback(event, x, y, flags, param):
            nonlocal selecting, selected_rect, screen_np
            
            if event == cv2.EVENT_LBUTTONDOWN:
                selecting = True
                selected_rect[0], selected_rect[1] = x, y
                
            elif event == cv2.EVENT_MOUSEMOVE and selecting:
                img_copy = screen_np.copy()
                cv2.rectangle(img_copy, (selected_rect[0], selected_rect[1]), (x, y), (0, 255, 0), 2)
                cv2.imshow(window_title, img_copy)
                
            elif event == cv2.EVENT_LBUTTONUP:
                selecting = False
                selected_rect[2], selected_rect[3] = x, y
                
                # 确保坐标是从左上到右下
                if selected_rect[0] > selected_rect[2]:
                    selected_rect[0], selected_rect[2] = selected_rect[2], selected_rect[0]
                if selected_rect[1] > selected_rect[3]:
                    selected_rect[1], selected_rect[3] = selected_rect[3], selected_rect[1]
                
                # 在图像上绘制最终选择的矩形
                img_copy = screen_np.copy()
                cv2.rectangle(img_copy, (selected_rect[0], selected_rect[1]), 
                             (selected_rect[2], selected_rect[3]), (0, 255, 0), 2)
                cv2.imshow(window_title, img_copy)
        
        # 设置鼠标回调
        cv2.setMouseCallback(window_title, mouse_callback)
        
        # 显示截图
        cv2.imshow(window_title, screen_np)
        log_message("请在截图上拖动鼠标选择区域，选择完成后按Enter键确认")
        
        # 等待用户选择区域
        while True:
            key = cv2.waitKey(1) & 0xFF
            if key == 13:  # Enter键
                if selected_rect[2] - selected_rect[0] > 10 and selected_rect[3] - selected_rect[1] > 10:
                    break
                else:
                    log_message("选择的区域太小，请重新选择")
            elif key == 27:  # ESC键
                cv2.destroyWindow(window_title)
                root.deiconify()
                return None
        
        # 关闭窗口
        cv2.destroyWindow(window_title)
        
        # 截取选中的区域并保存
        region_img = screen_np[selected_rect[1]:selected_rect[3], selected_rect[0]:selected_rect[2]]
        
        # 恢复主窗口
        root.deiconify()
        
        return (selected_rect[0], selected_rect[1], selected_rect[2], selected_rect[3], region_img)
    
    except Exception as e:
        log_message(f"截取屏幕区域时出错: {e}")
        log_message(traceback.format_exc())
        
        # 恢复主窗口
        root.deiconify()
        return None

def select_score_area():
    """选择分数输入区域"""
    global score_area_rect, score_area_selected
    
    log_message("请选择分数输入区域...")
    result = capture_screen_region("选择分数输入区域")
    
    if result:
        left, top, right, bottom, region_img = result
        score_area_rect = (left, top, right, bottom)
        score_area_selected = True
        
        # 保存区域截图到image文件夹
        score_img_path = os.path.join(IMAGE_DIR, "score_area.png")
        cv2.imwrite(score_img_path, region_img)
        log_message(f"已设置分数区域: 左={left}, 上={top}, 右={right}, 下={bottom}")
        log_message(f"已保存分数区域截图到 {score_img_path}")
        
        # 更新按钮状态
        btn_select_score.config(bg="green")
        
        # 显示截图预览
        show_preview(region_img, "分数区域预览")
        
        return True
    
    return False

def select_submit_area():
    """选择提交按钮区域"""
    global submit_area_rect, submit_area_selected
    
    log_message("请选择提交按钮区域...")
    result = capture_screen_region("选择提交按钮区域")
    
    if result:
        left, top, right, bottom, region_img = result
        submit_area_rect = (left, top, right, bottom)
        submit_area_selected = True
        
        # 保存区域截图到image文件夹
        submit_img_path = os.path.join(IMAGE_DIR, "submit_area.png")
        cv2.imwrite(submit_img_path, region_img)
        log_message(f"已设置提交按钮区域: 左={left}, 上={top}, 右={right}, 下={bottom}")
        log_message(f"已保存提交按钮区域截图到 {submit_img_path}")
        
        # 更新按钮状态
        btn_select_submit.config(bg="green")
        
        # 显示截图预览
        show_preview(region_img, "提交按钮区域预览")
        
        return True
    
    return False

def select_answer_area():
    """选择答案文本区域"""
    global answer_area_rect, answer_area_selected
    
    if not OCR_AVAILABLE:
        messagebox.showerror("错误", "OCR功能不可用，请安装pytesseract库")
        log_message("OCR功能不可用，请安装pytesseract库: pip install pytesseract")
        log_message("并安装Tesseract OCR: https://github.com/UB-Mannheim/tesseract/wiki")
        return False
    
    log_message("请选择答案文本区域...")
    result = capture_screen_region("选择答案文本区域")
    
    if result:
        left, top, right, bottom, region_img = result
        answer_area_rect = (left, top, right, bottom)
        answer_area_selected = True
        
        # 保存区域截图到image文件夹
        answer_img_path = os.path.join(IMAGE_DIR, "answer_area.png")
        cv2.imwrite(answer_img_path, region_img)
        log_message(f"已设置答案文本区域: 左={left}, 上={top}, 右={right}, 下={bottom}")
        log_message(f"已保存答案区域截图到 {answer_img_path}")
        
        # 尝试OCR识别图像中的文本
        try:
            # 转换图像格式用于OCR
            pil_img = PILImage.fromarray(cv2.cvtColor(region_img, cv2.COLOR_BGR2RGB))
            text = pytesseract.image_to_string(pil_img, lang='chi_sim+eng')
            
            if text and len(text.strip()) > 10:
                log_message(f"OCR识别结果预览: {text[:50]}...")
            else:
                log_message("OCR未能识别出足够的文本，请检查选择区域或图像质量")
        except Exception as e:
            log_message(f"OCR识别测试失败: {e}")
        
        # 更新按钮状态
        btn_select_answer.config(bg="green")
        
        # 显示截图预览
        show_preview(region_img, "答案文本区域预览")
        
        return True
    
    return False

def select_score_input_area():
    """选择分数输入框区域"""
    global score_input_rect, score_input_selected
    
    log_message("请选择分数输入框区域(右上角的0-10分输入框)...")
    result = capture_screen_region("选择分数输入框区域")
    
    if result:
        left, top, right, bottom, region_img = result
        score_input_rect = (left, top, right, bottom)
        score_input_selected = True
        
        # 保存区域截图到image文件夹
        img_path = os.path.join(IMAGE_DIR, "score_input_area.png")
        cv2.imwrite(img_path, region_img)
        log_message(f"已设置分数输入框区域: 左={left}, 上={top}, 右={right}, 下={bottom}")
        log_message(f"已保存分数输入框区域截图到 {img_path}")
        
        # 更新按钮状态
        btn_select_score_input.config(bg="green")
        
        # 显示截图预览
        show_preview(region_img, "分数输入框区域预览")
        
        return True
    
    return False

def select_final_submit_area():
    """选择最终提交按钮区域"""
    global final_submit_rect, final_submit_selected
    
    log_message("请选择最终提交按钮区域(页面右下角)...")
    result = capture_screen_region("选择最终提交按钮区域")
    
    if result:
        left, top, right, bottom, region_img = result
        final_submit_rect = (left, top, right, bottom)
        final_submit_selected = True
        
        # 保存区域截图到image文件夹
        img_path = os.path.join(IMAGE_DIR, "final_submit_area.png")
        cv2.imwrite(img_path, region_img)
        log_message(f"已设置最终提交按钮区域: 左={left}, 上={top}, 右={right}, 下={bottom}")
        log_message(f"已保存最终提交按钮区域截图到 {img_path}")
        
        # 更新按钮状态
        btn_select_final_submit.config(bg="green")
        
        # 显示截图预览
        show_preview(region_img, "最终提交按钮区域预览")
        
        return True
    
    return False

def show_preview(img, title="预览"):
    """显示预览图像"""
    try:
        # 创建新窗口
        preview_window = tk.Toplevel(root)
        preview_window.title(title)
        preview_window.attributes("-topmost", True)
        
        # 调整图像大小，确保不会太大
        max_height = 300
        height, width = img.shape[:2]
        if height > max_height:
            ratio = max_height / height
            new_width = int(width * ratio)
            img = cv2.resize(img, (new_width, max_height))
        
        # 转换为PIL格式
        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(img)
        img_tk = ImageTk.PhotoImage(image=img)
        
        # 显示图像
        label = tk.Label(preview_window, image=img_tk)
        label.image = img_tk  # 保持引用
        label.pack(padx=10, pady=10)
        
        # 添加确认按钮
        btn_confirm = tk.Button(preview_window, text="确认", command=preview_window.destroy)
        btn_confirm.pack(pady=5)
        
    except Exception as e:
        log_message(f"显示预览图像时出错: {e}")
        log_message(traceback.format_exc())

def enter_score():
    """在分数输入区域输入配置的分数"""
    try:
        if not score_area_selected:
            log_message("未选择分数区域，跳过输入分数步骤")
            return False
        
        # 计算分数输入区域的中心位置
        left, top, right, bottom = score_area_rect
        center_x = (left + right) // 2
        center_y = (top + bottom) // 2
        
        # 点击分数输入区域
        log_message(f"点击分数区域中心位置: ({center_x}, {center_y})")
        pyautogui.click(center_x, center_y)
        time.sleep(0.5)
        
        # 获取分数 - 使用AI评分或配置的固定分数
        score = config.config["score"]
        
        # 如果启用了AI评分功能
        if config.config.get("use_ai_scoring", False) and config.config.get("deepseek_api_key"):
            # 尝试对当前页面内容进行AI评分
            try:
                log_message("尝试使用AI评估答案...")
                
                # 使用剪贴板方法获取答案文本
                answer_text = get_answer_text_from_clipboard()
                
                if answer_text and len(answer_text.strip()) > 10:
                    log_message(f"获取到答案: {answer_text[:30]}...")
                    
                    # 获取题目内容（如果有）
                    question_text = config.config.get("question_content", "")
                    
                    # 使用AI评分
                    scorer = DeepSeekScorer()
                    ai_score, reason = scorer.score_answer(
                        answer_text, 
                        expected_score=None, 
                        question_text=question_text
                    )
                    
                    if ai_score is not None:
                        score = str(ai_score)
                        log_message(f"AI评分结果: {score}")
                    else:
                        log_message(f"AI评分失败: {reason}，使用默认分数")
                else:
                    log_message("未能获取有效答案文本，使用默认分数")
            except Exception as e:
                log_message(f"AI评分过程出错: {e}")
                log_message(traceback.format_exc())
        
        # 输入分数
        pyautogui.write(score)
        log_message(f"已输入分数：{score}")
        
        # 按Enter键确认
        pyautogui.press('enter')
        log_message("已按Enter键确认")
        
        return True
        
    except Exception as e:
        log_message(f"输入分数过程出错: {e}")
        log_message(traceback.format_exc())
        return False

def find_and_click_submit_button():
    """点击提交按钮"""
    try:
        if not submit_area_selected:
            log_message("未选择提交按钮区域")
            return False
        
        # 计算提交按钮区域的中心位置
        left, top, right, bottom = submit_area_rect
        center_x = (left + right) // 2
        center_y = (top + bottom) // 2
        
        log_message(f"点击提交按钮区域中心位置: ({center_x}, {center_y})")
        pyautogui.click(center_x, center_y)
        log_message("已点击提交按钮")
        return True
        
    except Exception as e:
        log_message(f"点击提交按钮过程出错: {e}")
        log_message(traceback.format_exc())
        return False

def show_config_window():
    """显示配置窗口"""
    config_window = tk.Toplevel(root)
    config_window.title("配置")
    config_window.geometry("600x400")
    config_window.resizable(True, False)
    config_window.attributes("-topmost", True)
    
    # 创建主框架，使用滚动条
    main_frame = tk.Frame(config_window)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    # 添加滚动条
    canvas = tk.Canvas(main_frame)
    scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    # 添加控件
    tk.Label(scrollable_frame, text="分数设置:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
    score_var = tk.StringVar(value=config.config["score"])
    score_entry = tk.Entry(scrollable_frame, textvariable=score_var, width=10)
    score_entry.grid(row=0, column=1, padx=10, pady=10, sticky="w")
    
    tk.Label(scrollable_frame, text="页面刷新等待时间(秒):").grid(row=1, column=0, padx=10, pady=10, sticky="w")
    wait_var = tk.IntVar(value=config.config["wait_time"])
    wait_entry = tk.Entry(scrollable_frame, textvariable=wait_var, width=10)
    wait_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")
    
    tk.Label(scrollable_frame, text="提交后等待时间(秒):").grid(row=2, column=0, padx=10, pady=10, sticky="w")
    after_submit_var = tk.IntVar(value=config.config["after_submit_wait"])
    after_submit_entry = tk.Entry(scrollable_frame, textvariable=after_submit_var, width=10)
    after_submit_entry.grid(row=2, column=1, padx=10, pady=10, sticky="w")
    
    # AI评分配置
    tk.Label(scrollable_frame, text="DeepSeek API密钥:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
    api_key_var = tk.StringVar(value=config.config.get("deepseek_api_key", ""))
    api_key_entry = tk.Entry(scrollable_frame, textvariable=api_key_var, width=40, show="*")
    api_key_entry.grid(row=3, column=1, padx=10, pady=10, sticky="w")
    
    # 使用AI评分复选框
    use_ai_var = tk.BooleanVar(value=config.config.get("use_ai_scoring", False))
    use_ai_check = tk.Checkbutton(scrollable_frame, text="使用AI自动评分", variable=use_ai_var)
    use_ai_check.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="w")
    
    # OCR语言设置
    tk.Label(scrollable_frame, text="OCR语言:").grid(row=5, column=0, padx=10, pady=10, sticky="w")
    ocr_lang_var = tk.StringVar(value=config.config.get("ocr_language", "chi_sim+eng"))
    ocr_lang_entry = tk.Entry(scrollable_frame, textvariable=ocr_lang_var, width=15)
    ocr_lang_entry.grid(row=5, column=1, padx=10, pady=10, sticky="w")
    
    # 题目内容
    tk.Label(scrollable_frame, text="题目内容:").grid(row=6, column=0, padx=10, pady=10, sticky="nw")
    question_var = tk.StringVar(value=config.config.get("question_content", ""))
    question_text = tk.Text(scrollable_frame, width=50, height=10)
    question_text.insert("1.0", question_var.get())
    question_text.grid(row=6, column=1, padx=10, pady=10, sticky="w")
    
    # 添加多题目打分配置
    tk.Label(scrollable_frame, text="多题目默认分数(0-10):").grid(row=8, column=0, padx=10, pady=10, sticky="w")
    multi_score_var = tk.StringVar(value=config.config.get("multi_question_score", "10"))
    multi_score_entry = tk.Entry(scrollable_frame, textvariable=multi_score_var, width=10)
    multi_score_entry.grid(row=8, column=1, padx=10, pady=10, sticky="w")
    
    # 随机变化分数复选框
    randomize_var = tk.BooleanVar(value=config.config.get("randomize_scores", False))
    randomize_check = tk.Checkbutton(scrollable_frame, text="分数随机微调(±1)", variable=randomize_var)
    randomize_check.grid(row=9, column=0, columnspan=2, padx=10, pady=5, sticky="w")
    
    # 添加智能滚动选项
    smart_scrolling_var = tk.BooleanVar(value=config.config.get("use_smart_scrolling", True))
    smart_scrolling_check = tk.Checkbutton(scrollable_frame, text="使用智能定位滚动", variable=smart_scrolling_var)
    smart_scrolling_check.grid(row=10, column=0, columnspan=2, padx=10, pady=5, sticky="w")
    
    # 添加滚动步长设置
    tk.Label(scrollable_frame, text="智能滚动步长:").grid(row=11, column=0, padx=10, pady=10, sticky="w")
    step_size_var = tk.StringVar(value=str(config.config.get("scroll_step_size", "-10")))
    step_size_entry = tk.Entry(scrollable_frame, textvariable=step_size_var, width=10)
    step_size_entry.grid(row=11, column=1, padx=10, pady=10, sticky="w")
    
    def save_settings():
        try:
            # 更新配置
            config.config["score"] = score_var.get()
            config.config["wait_time"] = wait_var.get()
            config.config["after_submit_wait"] = after_submit_var.get()
            config.config["deepseek_api_key"] = api_key_var.get()
            config.config["use_ai_scoring"] = use_ai_var.get()
            config.config["ocr_language"] = ocr_lang_var.get()
            config.config["question_content"] = question_text.get("1.0", tk.END).strip()
            config.config["multi_question_score"] = multi_score_var.get()
            config.config["randomize_scores"] = randomize_var.get()
            config.config["use_smart_scrolling"] = smart_scrolling_var.get()
            config.config["scroll_step_size"] = step_size_var.get()
            
            # 保存配置
            if config.save_config():
                log_message("配置已保存")
                config_window.destroy()
            else:
                messagebox.showerror("错误", "保存配置失败")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置时出错: {e}")
    
    # 保存按钮
    save_btn = tk.Button(scrollable_frame, text="保存", command=save_settings)
    save_btn.grid(row=12, column=0, columnspan=2, pady=20)

def start_grading():
    global running
    
    try:
        # 循环执行批改，直到用户暂停
        while running:
            try:
                # 第一步：输入分数
                log_message("尝试输入分数...")
                if enter_score():
                    status_label.config(text=f"状态：已输入分数")
                    # 检查是否暂停
                    if not check_running(1):
                        break
                else:
                    log_message("未能输入分数，尝试直接提交")
                
                # 第二步：查找并点击提交按钮
                log_message("点击提交按钮...")
                if find_and_click_submit_button():
                    log_message("已点击提交按钮，完成一次批改")
                    status_label.config(text="状态：已提交，等待下一个...")
                    # 使用分段等待代替单个长等待
                    wait_time = config.config["wait_time"]
                    if not check_running(wait_time):  # 使用配置的等待时间
                        break
                    log_message("页面刷新等待完成，准备处理下一个")
                else:
                    log_message("未能点击提交按钮，跳过当前项目")
                    if not check_running(2):
                        break
                    continue
                
            except Exception as e:
                error_msg = f"批改过程中发生错误: {str(e)}"
                log_message(error_msg)
                log_message(traceback.format_exc())
                status_label.config(text=f"状态：发生错误")
                if not check_running(2):
                    break
            
            # 每次批改完成后暂停一下，避免过快操作
            after_submit_wait = config.config["after_submit_wait"]
            if not check_running(after_submit_wait):  # 使用配置的等待时间
                break
    
    except Exception as e:
        error_msg = f"发生严重错误：{e}"
        log_message(error_msg)
        log_message(traceback.format_exc())
        messagebox.showerror("错误", error_msg)
        btn_toggle.config(text="开始", bg="green")
        status_label.config(text="状态：已暂停")
        running = False

def check_running(seconds):
    """分段等待并检查running状态"""
    global running
    for _ in range(int(seconds * 10)):  # 将秒拆分成0.1秒的间隔
        if not running:
            return False
        time.sleep(0.1)
    return running

def on_closing():
    global running
    running = False
    root.destroy()

def get_answer_text_from_clipboard():
    """从剪贴板获取答案文本"""
    global answer_area_rect
    
    try:
        if not answer_area_selected or not answer_area_rect:
            log_message("未选择答案区域，请先选择答案区域")
            return None
            
        # 计算答案区域的中心位置
        left, top, right, bottom = answer_area_rect
        center_x = (left + right) // 2
        center_y = (top + bottom) // 2
        
        # 弹出提示窗口
        copy_window = tk.Toplevel(root)
        copy_window.title("复制文本")
        copy_window.geometry("300x150")
        copy_window.attributes("-topmost", True)
        
        tk.Label(copy_window, text="请按以下步骤操作:", font=("Arial", 10, "bold")).pack(pady=5)
        tk.Label(copy_window, text="1. 点击'点击区域'按钮").pack(anchor=tk.W, padx=20)
        tk.Label(copy_window, text="2. 选择答案文本并复制(Ctrl+C)").pack(anchor=tk.W, padx=20)
        tk.Label(copy_window, text="3. 点击'获取文本'按钮").pack(anchor=tk.W, padx=20)
        
        answer_text = [None]  # 使用列表存储，以便在内部函数中修改
        
        def click_area():
            """点击答案区域"""
            copy_window.iconify()  # 最小化窗口
            time.sleep(0.5)
            # 点击答案区域
            pyautogui.click(center_x, center_y)
            time.sleep(0.5)
            copy_window.deiconify()  # 恢复窗口
        
        def get_text():
            """从剪贴板获取文本"""
            try:
                answer_text[0] = root.clipboard_get()
                if answer_text[0] and len(answer_text[0].strip()) > 5:
                    log_message(f"已获取答案文本: {answer_text[0][:30]}...")
                    copy_window.destroy()
                else:
                    messagebox.showwarning("警告", "剪贴板中没有足够的文本，请重试")
            except Exception as e:
                log_message(f"获取剪贴板内容出错: {e}")
                messagebox.showerror("错误", f"获取剪贴板内容出错: {e}")
        
        button_frame = tk.Frame(copy_window)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="点击区域", command=click_area).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="获取文本", command=get_text).pack(side=tk.LEFT, padx=10)
        
        # 等待窗口关闭
        root.wait_window(copy_window)
        return answer_text[0]
        
    except Exception as e:
        log_message(f"获取答案文本过程出错: {e}")
        log_message(traceback.format_exc())
        return None

def set_scroll_amount():
    """设置滚动量"""
    global scroll_amount
    
    scroll_window = tk.Toplevel(root)
    scroll_window.title("设置滚动量")
    scroll_window.geometry("350x200")
    scroll_window.attributes("-topmost", True)
    
    # 添加说明文字
    tk.Label(scroll_window, text="设置滚动量(负值向下滚动，正值向上滚动):", font=("Arial", 9)).pack(pady=10)
    tk.Label(scroll_window, text="推荐值: 普通页面 -10 ~ -50, 长页面可用 -100 ~ -300", font=("Arial", 8, "italic")).pack()
    
    # 创建输入框架
    input_frame = tk.Frame(scroll_window)
    input_frame.pack(pady=15)
    
    tk.Label(input_frame, text="滚动量:").pack(side=tk.LEFT, padx=5)
    
    # 创建输入框，默认值为当前滚动量
    scroll_var = tk.StringVar(value=str(scroll_amount))
    scroll_entry = tk.Entry(input_frame, textvariable=scroll_var, width=10)
    scroll_entry.pack(side=tk.LEFT, padx=5)
    
    # 保存设置
    def save_scroll():
        try:
            global scroll_amount
            new_value = int(scroll_var.get())
            scroll_amount = new_value
            log_message(f"滚动量已设置为: {scroll_amount}")
            scroll_window.destroy()
        except ValueError:
            messagebox.showerror("错误", "请输入有效的整数值")
    
    # 测试滚动
    def test_scroll():
        try:
            test_value = int(scroll_var.get())
            log_message(f"测试滚动量: {test_value}")
            # 最小化窗口进行测试
            scroll_window.iconify()
            time.sleep(0.5)
            pyautogui.scroll(test_value)
            time.sleep(0.5)
            scroll_window.deiconify()
        except ValueError:
            messagebox.showerror("错误", "请输入有效的整数值")
    
    # 添加按钮
    button_frame = tk.Frame(scroll_window)
    button_frame.pack(pady=15)
    
    tk.Button(button_frame, text="测试滚动", command=test_scroll, width=10).pack(side=tk.LEFT, padx=15)
    tk.Button(button_frame, text="保存", command=save_scroll, width=10).pack(side=tk.LEFT, padx=15)

def toggle_multipage_grading():
    """切换多题目打分状态"""
    global running
    
    if running:
        stop_multipage_grading()
    else:
        start_multipage_grading()

def start_multipage_grading():
    """开始多题目滚动打分"""
    global running, current_question
    
    try:
        # 检查必要区域是否已选择
        if not score_input_selected:
            messagebox.showwarning("警告", "请先选择分数输入框区域")
            return
            
        if not final_submit_selected:
            messagebox.showwarning("警告", "请先选择最终提交按钮区域")
            return
            
        running = True
        current_question = 1
        btn_start_multipage.config(text="暂停", bg="red")
        status_label.config(text="状态：多题目打分中")
        log_message("开始多题目打分流程")
        
        # 在新线程中启动批改，避免界面卡顿
        threading.Thread(target=run_multipage_grading, daemon=True).start()
    except Exception as e:
        log_message(f"启动多题目打分出错: {e}")
        log_message(traceback.format_exc())

def stop_multipage_grading():
    """停止多题目滚动打分"""
    global running
    
    running = False
    btn_start_multipage.config(text="开始多题目打分", bg="purple")
    status_label.config(text="状态：已暂停")
    log_message("多题目打分已暂停")

def run_multipage_grading():
    """执行多题目滚动打分的主循环"""
    global running, current_question
    
    try:
        while running:
            # 1. 输入当前题目的分数
            log_message(f"第{current_question}题: 输入分数...")
            if not input_score_for_question():
                # 如果无法输入分数，可能是已到达底部，尝试提交
                log_message("无法继续输入分数，尝试点击提交按钮")
                click_final_submit()
                break
                
            # 2. 等待短暂时间
            if not check_running(1):
                break
                
            # 3. 滚动到下一题
            log_message(f"滚动到下一题...")
            scroll_to_next_question()
            
            # 4. 等待页面稳定
            if not check_running(1.5):
                break
                
            # 5. 递增题目计数
            current_question += 1
            
        # 如果因为running=False而退出循环
        if not running:
            log_message("多题目打分已停止")
            btn_start_multipage.config(text="开始多题目打分", bg="purple")
            status_label.config(text="状态：已暂停")
    except Exception as e:
        error_msg = f"多题目打分过程出错: {str(e)}"
        log_message(error_msg)
        log_message(traceback.format_exc())
        messagebox.showerror("错误", error_msg)
        stop_multipage_grading()

def input_score_for_question():
    """为当前题目输入分数"""
    try:
        if not score_input_selected:
            log_message("未选择分数输入框区域")
            return False
            
        # 计算分数输入框中心位置
        left, top, right, bottom = score_input_rect
        center_x = (left + right) // 2
        center_y = (top + bottom) // 2
        
        # 点击分数输入框
        log_message(f"点击分数输入框: ({center_x}, {center_y})")
        pyautogui.click(center_x, center_y)
        time.sleep(0.5)
        
        # 从配置中获取分数，或根据题目算法确定分数
        score = get_score_for_current_question()
        
        # 输入分数
        pyautogui.write(str(score))
        log_message(f"输入分数: {score}")
        
        # 按Enter键确认
        pyautogui.press('enter')
        time.sleep(0.2)
        
        return True
    except Exception as e:
        log_message(f"输入分数过程出错: {e}")
        log_message(traceback.format_exc())
        return False

def get_score_for_current_question():
    """获取当前题目的分数"""
    # 可以根据题目类型或配置返回不同分数
    # 这里简单使用配置中的固定分数，后续可以扩展
    base_score = int(config.config.get("multi_question_score", "10"))
    
    # 如果需要随机变化一点分数
    if config.config.get("randomize_scores", False):
        import random
        random_offset = random.choice([-1, 0, 0, 0, 1])  # 大部分情况保持不变
        score = max(0, min(10, base_score + random_offset))
        return score
    
    return base_score

def scroll_to_next_question():
    """智能滚动到下一个题目的分数输入框"""
    global scroll_amount, score_input_rect, current_question
    
    try:
        if not score_input_selected:
            log_message("未选择分数输入框区域，无法智能定位")
            # 回退到普通滚动
            pyautogui.scroll(scroll_amount)
            return
            
        # 记录当前分数框位置，用于后续比较
        original_left, original_top, original_right, original_bottom = score_input_rect
        original_center_x = (original_left + original_right) // 2
        original_center_y = (original_top + original_bottom) // 2
        
        # 获取分数输入框的模板图像
        template = pyautogui.screenshot(region=(original_left, original_top, original_right-original_left, original_bottom-original_top))
        template = np.array(template)
        template = cv2.cvtColor(template, cv2.COLOR_RGB2BGR)
        
        # 使用小增量滚动
        max_attempts = 20  # 增加最大尝试次数，以适应小增量滚动
        current_attempt = 0
        # 使用较小的滚动步长，确保不会跳过目标
        scroll_step = max(-70, min(-50, scroll_amount // 5))  # 限制在-5到-20之间
        if scroll_amount > 0:
            scroll_step = min(20, max(5, scroll_amount // 5))  # 处理向上滚动的情况
            
        log_message(f"使用小增量滚动模式，步长: {scroll_step}")
        
        found = False
        last_max_val = 0  # 记录最后一次的匹配度
        best_match = None  # 记录最佳匹配
        
        while current_attempt < max_attempts and not found and running:
            # 执行小增量滚动
            pyautogui.scroll(scroll_step)
            log_message(f"滚动尝试 {current_attempt+1}/{max_attempts}, 步长: {scroll_step}")
            time.sleep(0.5)  # 等待页面稳定
            
            # 屏幕截图
            screen = pyautogui.screenshot()
            screen = np.array(screen)
            screen = cv2.cvtColor(screen, cv2.COLOR_RGB2BGR)
            
            # 用模板匹配查找分数输入框
            result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            log_message(f"匹配度: {max_val:.2f}")
            
            # 设置匹配阈值
            threshold = 0.8
            
            if max_val >= threshold:
                # 找到匹配
                match_x, match_y = max_loc
                match_w, match_h = original_right-original_left, original_bottom-original_top
                match_center_x = match_x + match_w // 2
                match_center_y = match_y + match_h // 2
                
                # 判断是否是新的分数框，而不是原来的
                horizontal_diff = abs(match_center_x - original_center_x)
                vertical_diff = abs(match_center_y - original_center_y)
                
                log_message(f"位置差异: 水平={horizontal_diff}像素, 垂直={vertical_diff}像素")
                
                # 水平位置应接近，垂直位置必须有明显差异(至少40像素)
                if horizontal_diff < 50 and vertical_diff > 40:
                    log_message(f"找到新题目分数框，位置: ({match_center_x}, {match_center_y})")
                    
                    # 额外检查：确保输入框为空白
                    score_input_rect = (match_x, match_y, match_x + match_w, match_y + match_h)
                    if verify_empty_score():
                        found = True
                        break
                    else:
                        log_message("检测到的输入框可能已有内容，继续寻找")
                        time.sleep(0.2)
                
            # 在任何情况下，记录最佳匹配
            if max_val > last_max_val and max_val > 0.6:
                last_x, last_y = max_loc if max_loc else (0, 0)
                
                # 确保当前检测位置不是原始位置
                if abs(last_y - original_center_y) > 30:
                    last_max_val = max_val
                    best_match = (max_val, max_loc)
            
            current_attempt += 1
        
        if not found:
            # 如果没找到高度匹配的，但有较好的匹配，也可以使用
            if best_match and best_match[0] > 0.65:
                match_x, match_y = best_match[1]
                match_w, match_h = original_right-original_left, original_bottom-original_top
                
                log_message(f"使用次优匹配: 匹配度={best_match[0]:.2f}")
                score_input_rect = (match_x, match_y, match_x + match_w, match_y + match_h)
                return True
            else:
                log_message("无法找到下一题的分数输入框，可能已到达最后一题")
                return False
        
        return True
        
    except Exception as e:
        log_message(f"智能滚动过程出错: {e}")
        log_message(traceback.format_exc())
        # 出错时回退到普通滚动
        pyautogui.scroll(scroll_amount)
        return False

def verify_empty_score():
    """验证当前分数框是否为空白（未填写）"""
    try:
        if not score_input_selected:
            return True
            
        # 截取分数框图像
        left, top, right, bottom = score_input_rect
        score_img = pyautogui.screenshot(region=(left, top, right-left, bottom-top))
        
        # 将图像转为灰度
        score_img_np = np.array(score_img)
        gray = cv2.cvtColor(score_img_np, cv2.COLOR_RGB2GRAY)
        
        # 检查平均亮度
        avg_brightness = np.mean(gray)
        log_message(f"分数框亮度: {avg_brightness:.1f}")
        
        # 根据亮度判断是否为空白框
        # 此阈值需要根据实际情况调整
        if avg_brightness > 200:  # 亮度高通常表示空白
            log_message("验证成功: 分数框为空")
            return True
        else:
            log_message("警告: 分数框可能已有内容")
            return False
    except Exception as e:
        log_message(f"验证分数框出错: {e}")
        return True  # 验证失败时默认继续

def click_final_submit():
    """点击最终提交按钮，提交所有评分"""
    try:
        if not final_submit_selected:
            log_message("未选择最终提交按钮区域")
            return False
            
        # 获取提交按钮中心位置
        left, top, right, bottom = final_submit_rect
        center_x = (left + right) // 2
        center_y = (top + bottom) // 2
        
        log_message(f"点击最终提交按钮: ({center_x}, {center_y})")
        
        # 点击提交按钮
        pyautogui.click(center_x, center_y)
        log_message("已点击最终提交按钮")
        
        # 不停止程序，等待页面加载新内容
        time.sleep(3.0)  # 等待更长时间，确保页面刷新
        log_message("页面已提交，准备处理下一批题目...")
        
        return True
    except Exception as e:
        log_message(f"点击最终提交按钮出错: {e}")
        return False

def create_score_box_template():
    """创建分数输入框的模板"""
    try:
        # 基于屏幕上的一个分数输入框创建模板
        if not score_input_selected:
            log_message("请先选择一个分数输入框作为模板")
            return None
            
        # 获取选定的分数框图像
        left, top, right, bottom = score_input_rect
        template = pyautogui.screenshot(region=(left, top, right-left, bottom-top))
        template = np.array(template)
        template = cv2.cvtColor(template, cv2.COLOR_RGB2BGR)
        
        # 保存模板
        template_path = os.path.join(IMAGE_DIR, "score_box_template.png")
        cv2.imwrite(template_path, template)
        log_message(f"已创建分数框模板: {template_path}")
        
        return template
    except Exception as e:
        log_message(f"创建模板出错: {e}")
        return None

def find_all_visible_score_boxes():
    """在当前屏幕上查找所有可见的分数输入框"""
    try:
        template = create_score_box_template()
        if template is None:
            return None
            
        # 截取当前屏幕
        screen = pyautogui.screenshot()
        screen = np.array(screen)
        screen = cv2.cvtColor(screen, cv2.COLOR_RGB2BGR)
        
        # 使用模板匹配查找所有分数框
        result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
        
        # 设置匹配阈值
        threshold = 0.8
        
        # 找到所有匹配
        locations = np.where(result >= threshold)
        h, w = template.shape[:2]
        
        # 过滤重复位置
        filtered_boxes = []
        last_x, last_y = -100, -100  # 初始化上一个位置
        
        for y, x in zip(*locations[::-1]):
            # 如果与上一个框距离太近，认为是重复检测
            if abs(x - last_x) > 20 or abs(y - last_y) > 20:
                filtered_boxes.append((x, y, w, h))
                last_x, last_y = x, y
        
        log_message(f"共找到 {len(filtered_boxes)} 个分数输入框")
        
        # 按垂直位置排序
        filtered_boxes.sort(key=lambda box: box[1])
        
        return filtered_boxes
    except Exception as e:
        log_message(f"查找分数框出错: {e}")
        log_message(traceback.format_exc())
        return None

def improved_multipage_scoring():
    """改进的多题目打分方法"""
    try:
        # 找到当前可见的所有分数框
        visible_boxes = find_all_visible_score_boxes()
        if not visible_boxes:
            log_message("未找到分数输入框，请检查模板")
            return False
            
        # 依次为每个分数框输入分数
        for i, (x, y, w, h) in enumerate(visible_boxes):
            if not running:
                break
                
            center_x = x + w//2
            center_y = y + h//2
            
            log_message(f"处理第 {i+1} 个分数框，位置: ({center_x}, {center_y})")
            
            # 点击分数框
            pyautogui.click(center_x, center_y)
            time.sleep(0.3)
            
            # 输入分数
            score = get_score_for_current_question()
            pyautogui.write(str(score))
            log_message(f"输入分数: {score}")
            time.sleep(0.2)
        
        # 处理完可见的输入框后滚动页面
        log_message("已处理所有可见分数框，滚动页面继续...")
        pyautogui.scroll(scroll_amount)
        time.sleep(1.5)  # 给页面滚动足够的时间
        
        # 截取新页面再次查找分数框
        new_visible_boxes = find_all_visible_score_boxes()
        
        # 检查是否有新的分数框出现
        if new_visible_boxes and len(new_visible_boxes) > 0:
            # 查看是否有新框出现（垂直位置不同）
            has_new_boxes = False
            for new_box in new_visible_boxes:
                is_new = True
                for old_box in visible_boxes:
                    # 如果新旧框位置接近，则认为是同一个框
                    if abs(new_box[0] - old_box[0]) < 20 and abs(new_box[1] - old_box[1]) < 20:
                        is_new = False
                        break
                if is_new:
                    has_new_boxes = True
                    break
                    
            if has_new_boxes:
                log_message("检测到新的分数框，继续处理...")
                return improved_multipage_scoring()  # 递归调用
            else:
                log_message("未检测到新的分数框，可能已到达页面底部")
                # 点击提交按钮
                return click_final_submit()
        else:
            log_message("未检测到更多分数框，准备提交...")
            return click_final_submit()
            
    except Exception as e:
        log_message(f"改进的多题目打分出错: {e}")
        log_message(traceback.format_exc())
        return False

def toggle_improved_scoring():
    """启动/停止智能打分"""
    global running
    
    if not running:
        # 开始打分
        running = True
        btn_start_multipage.config(text="停止打分", bg="red")
        status_label.config(text="状态：正在执行智能打分")
        
        # 检查OCR是否可用，选择相应的方法
        if OCR_AVAILABLE and os.path.exists(pytesseract.pytesseract.tesseract_cmd):
            log_message("启动基于OCR的题号智能打分")
            threading.Thread(target=question_based_scoring, daemon=True).start()
        else:
            log_message("OCR不可用，使用模板匹配方法打分")
            threading.Thread(target=question_based_scoring_without_ocr, daemon=True).start()
    else:
        # 停止打分
        running = False
        btn_start_multipage.config(text="开始多题目打分", bg="purple")
        status_label.config(text="状态：已停止打分")
        log_message("已停止智能打分")

def select_question_number_area():
    """选择题号文本区域"""
    global question_number_rect, question_number_selected
    
    log_message('请选择一个题号区域（如"第1题:"）')
    
    question_number_rect = capture_screen_region("选择题号区域")
    
    if question_number_rect:
        question_number_selected = True
        log_message(f"题号区域已选择: {question_number_rect}")
        
        # 截取并保存题号图像，用于调试
        left, top, right, bottom = question_number_rect
        img = pyautogui.screenshot(region=(left, top, right-left, bottom-top))
        img_path = os.path.join(IMAGE_DIR, "question_number.png")
        img.save(img_path)
        log_message(f"题号图像已保存: {img_path}")
    else:
        log_message("题号区域选择已取消")

def find_all_question_numbers():
    """识别当前屏幕上所有的题号"""
    try:
        # 截取整个屏幕
        screen = pyautogui.screenshot()
        screen_np = np.array(screen)
        screen_bgr = cv2.cvtColor(screen_np, cv2.COLOR_RGB2BGR)
        
        question_numbers = []
        
        # 如果OCR可用，使用OCR识别题号
        if OCR_AVAILABLE and question_number_selected:
            # 先截取一个宽范围的区域
            # 假设题号集中在屏幕的左半部分
            screen_width = screen.width
            screen_height = screen.height
            
            # 截取左侧区域
            left_area = screen_bgr[:, :int(screen_width*0.5)]
            
            # 使用OCR识别文本
            ocr_results = pytesseract.image_to_data(left_area, lang='chi_sim', 
                                                   output_type=pytesseract.Output.DICT)
            
            # 查找所有"第X题"文本
            for i, text in enumerate(ocr_results['text']):
                match = re.search(r'第(\d+)题', text)
                if match:
                    question_num = int(match.group(1))
                    x = ocr_results['left'][i]
                    y = ocr_results['top'][i]
                    w = ocr_results['width'][i]
                    h = ocr_results['height'][i]
                    conf = ocr_results['conf'][i]
                    
                    if conf > 60:  # 只接受置信度较高的结果
                        # 找到题号右侧的分数输入框
                        # 分数框通常在题号右侧并位于同一水平线附近
                        input_box = find_score_input_near_question(question_num, x+w, y, screen_bgr)
                        
                        if input_box:
                            question_numbers.append((question_num, input_box))
                            log_message(f"找到第{question_num}题，输入框位置: {input_box}")
            
            # 按题号排序
            question_numbers.sort(key=lambda x: x[0])
            
            return question_numbers
        else:
            log_message("OCR不可用或题号区域未选择，无法执行基于题号的打分")
            return None
    except Exception as e:
        log_message(f"识别题号出错: {e}")
        log_message(traceback.format_exc())
        return None

def find_score_input_near_question(question_num, start_x, start_y, screen_bgr):
    """找到题号附近的分数输入框"""
    try:
        # 在当前行右侧搜索"得分："文本
        # 这里可以用模板匹配或OCR来识别
        if OCR_AVAILABLE:
            # 定义搜索区域（题号右侧到屏幕右边缘）
            height, width = screen_bgr.shape[:2]
            search_width = min(700, width - start_x)  # 限制搜索区域宽度
            search_height = 120  # 向下扩展一些，以防"得分"文本在题号下方
            
            # 确保搜索区域在屏幕内
            search_x = max(0, start_x)
            search_y = max(0, start_y - 30)  # 向上偏移一点以增加容错性
            search_width = min(search_width, width - search_x)
            search_height = min(search_height, height - search_y)
            
            search_area = screen_bgr[search_y:search_y+search_height, 
                                    search_x:search_x+search_width]
            
            # 保存搜索区域图像，用于调试
            search_img_path = os.path.join(IMAGE_DIR, f"search_q{question_num}.png")
            cv2.imwrite(search_img_path, search_area)
            
            # 使用OCR识别文本
            search_results = pytesseract.image_to_data(search_area, lang='chi_sim', 
                                                     output_type=pytesseract.Output.DICT)
            
            # 查找"得分："文本
            for i, text in enumerate(search_results['text']):
                if "得分" in text:
                    score_x = search_results['left'][i]
                    score_y = search_results['top'][i]
                    score_w = search_results['width'][i]
                    score_h = search_results['height'][i]
                    
                    # 分数输入框通常在"得分："文本右侧
                    input_x = search_x + score_x + score_w + 5
                    input_y = search_y + score_y
                    input_w = 80  # 估计宽度
                    input_h = score_h
                    
                    log_message(f"第{question_num}题旁找到'得分'文本，估计输入框位置: ({input_x}, {input_y})")
                    
                    # 返回输入框的位置
                    return (input_x, input_y, input_x + input_w, input_y + input_h)
            
            # 如果没有找到"得分："文本，尝试直接定位输入框
            # 典型的分数输入框是一个浅色背景的矩形区域
            # 这里可以通过边缘检测或模板匹配来找到
            
            # 作为后备方案，我们假设输入框在题号右侧固定距离处
            backup_input_x = search_x + 300  # 根据页面布局估计
            backup_input_y = start_y
            backup_input_w = 60
            backup_input_h = 30
            
            log_message(f"未找到'得分'文本，使用估计位置: ({backup_input_x}, {backup_input_y})")
            return (backup_input_x, backup_input_y, backup_input_x + backup_input_w, backup_input_y + backup_input_h)
        
    except Exception as e:
        log_message(f"寻找输入框出错: {e}")
        return None

def question_based_scoring():
    """基于题号的智能打分主函数"""
    global running
    
    try:
        # 初始化已处理题号集合
        processed_questions = set()
        total_processed = 0
        consecutive_no_new = 0
        
        while running:
            # 识别当前页面上所有题号
            question_boxes = find_all_question_numbers()
            
            if not question_boxes or len(question_boxes) == 0:
                log_message("当前页面未发现题号，尝试滚动页面...")
                pyautogui.scroll(scroll_amount)
                time.sleep(1.0)
                consecutive_no_new += 1
                
                # 如果连续多次没有发现新题号，可能已到达页面底部
                if consecutive_no_new >= 3:
                    log_message("连续多次未发现新题号，可能已完成所有题目")
                    break
                    
                continue
            
            # 处理新发现的题号
            new_found = False
            
            for question_num, box in question_boxes:
                # 跳过已处理的题号
                if question_num in processed_questions:
                    continue
                    
                new_found = True
                consecutive_no_new = 0
                
                log_message(f"处理第{question_num}题")
                
                # 计算输入框中心位置
                left, top, right, bottom = box
                center_x = (left + right) // 2
                center_y = (top + bottom) // 2
                
                # 点击输入框
                pyautogui.click(center_x, center_y)
                time.sleep(0.3)
                
                # 输入分数
                score = get_score_for_current_question()
                pyautogui.write(str(score))
                log_message(f"为第{question_num}题输入分数: {score}")
                
                # 按回车确认
                time.sleep(0.2)
                pyautogui.press('enter')
                time.sleep(0.1)
                
                # 记录已处理的题号
                processed_questions.add(question_num)
                total_processed += 1
            
            # 如果当前页面没有新题号，滚动页面
            if not new_found:
                log_message("当前页面未发现新题号，滚动页面...")
                pyautogui.scroll(scroll_amount)
                time.sleep(1.0)
                consecutive_no_new += 1
                
                # 如果连续多次没有发现新题号，可能已到达页面底部
                if consecutive_no_new >= 3:
                    log_message(f"连续多次未发现新题号，可能已完成所有题目")
                    break
            else:
                # 如果有新题号被处理，继续滚动查找更多
                log_message("继续滚动查找更多题目...")
                pyautogui.scroll(scroll_amount)
                time.sleep(1.0)
        
        log_message(f"基于题号的智能打分已完成，共处理{total_processed}道题目")
        
        # 点击提交按钮
        if total_processed > 0 and running:
            log_message("准备提交所有评分...")
            click_final_submit()
            
        # 恢复按钮状态
        if running:
            running = False
            root.after(0, lambda: btn_start_multipage.config(text="开始多题目打分", bg="purple"))
            root.after(0, lambda: status_label.config(text="状态：已完成智能打分"))
            
    except Exception as e:
        log_message(f"基于题号的智能打分出错: {e}")
        log_message(traceback.format_exc())
        
        # 恢复按钮状态
        running = False
        root.after(0, lambda: btn_start_multipage.config(text="开始多题目打分", bg="purple"))
        root.after(0, lambda: status_label.config(text="状态：出错"))

def question_based_scoring_without_ocr():
    """不依赖OCR的智能打分函数 - 连续处理多批次"""
    global running, current_question, score_input_rect
    
    try:
        if not score_input_selected:
            log_message("请先选择分数输入框")
            return False
            
        # 外部循环 - 处理多批次题目
        batch_count = 0
        total_all_batches = 0
        
        while running:
            log_message(f"======= 开始处理第 {batch_count+1} 批题目 =======")
            
            # 保存已处理的位置，每批次重置
            processed_positions = []
            
            # 获取分数输入框模板 - 每批次重新截取
            left, top, right, bottom = score_input_rect
            template = pyautogui.screenshot(region=(left, top, right-left, bottom-top))
            template = np.array(template)
            template = cv2.cvtColor(template, cv2.COLOR_RGB2BGR)
            template_width = right - left
            template_height = bottom - top
            
            consecutive_no_find = 0
            total_processed = 0
            
            # 内部循环 - 处理当前批次的所有题目
            while running and consecutive_no_find < 5:
                # 截取当前屏幕
                screen = pyautogui.screenshot()
                screen_np = np.array(screen)
                screen_bgr = cv2.cvtColor(screen_np, cv2.COLOR_RGB2BGR)
                
                # 模板匹配查找所有分数输入框
                result = cv2.matchTemplate(screen_bgr, template, cv2.TM_CCOEFF_NORMED)
                threshold = 0.8
                
                # 找到所有匹配位置
                locations = np.where(result >= threshold)
                points = list(zip(*locations[::-1]))  # (x,y)坐标列表
                
                # 按垂直位置排序 - 确保从上到下处理
                points.sort(key=lambda pt: pt[1])
                
                # 查找未处理的框
                found_new = False
                
                for pt in points:
                    x, y = pt
                    
                    # 跳过已处理的位置 - 严格检查
                    is_processed = False
                    for px, py in processed_positions:
                        # 使用更严格的距离判断
                        if abs(x - px) < template_width/2 and abs(y - py) < template_height/2:
                            is_processed = True
                            break
                    
                    if not is_processed:
                        # 找到一个未处理的框
                        center_x = x + template_width // 2
                        center_y = y + template_height // 2
                        
                        log_message(f"处理分数框 #{total_processed+1}, 位置: ({center_x}, {center_y})")
                        
                        # 点击输入框
                        pyautogui.click(center_x, center_y)
                        time.sleep(0.4)  # 稍微增加等待时间，确保点击生效
                        
                        # 输入分数
                        score = get_score_for_current_question()
                        pyautogui.write(str(score))
                        log_message(f"输入分数: {score}")
                        time.sleep(0.3)
                        
                        # 确认输入 - 按回车键
                        pyautogui.press('enter')
                        time.sleep(0.2)
                        
                        # 记录已处理位置
                        processed_positions.append((x, y))
                        total_processed += 1
                        found_new = True
                        
                        # 重要：每次只处理一个分数框，然后滚动
                        break
                
                # 是否找到新的框决定下一步操作
                if found_new:
                    log_message("已处理一个分数框，滚动页面继续...")
                    consecutive_no_find = 0
                else:
                    log_message("当前页面未找到新的分数框，滚动页面...")
                    consecutive_no_find += 1
                    if consecutive_no_find >= 3:
                        log_message("连续多次未找到新分数框，当前批次可能已完成")
                
                # 无论是否找到，都滚动到下一个
                pyautogui.scroll(scroll_amount)
                time.sleep(1.2)  # 增加等待时间，确保页面完全加载
            
            log_message(f"当前批次打分完成，共处理 {total_processed} 个分数框")
            total_all_batches += total_processed
            
            # 点击提交按钮
            if total_processed > 0 and running:
                log_message("准备提交当前批次...")
                if click_final_submit():
                    batch_count += 1
                    log_message(f"已完成 {batch_count} 批次，总计评分 {total_all_batches} 题")
                    time.sleep(2.0)  # 等待页面完全加载
                else:
                    log_message("提交失败，停止处理")
                    break
            else:
                # 如果没有处理任何框，可能是已经全部完成
                log_message("本批次未找到可处理的题目，可能已全部完成")
                if batch_count == 0:
                    # 如果第一批就没找到，可能是初始选择有问题
                    log_message("未能找到任何可处理的分数框，请检查选择的模板")
                break
        
        log_message(f"=== 多批次处理结束，共完成 {batch_count} 批次，总计 {total_all_batches} 题 ===")
        return True
        
    except Exception as e:
        log_message(f"连续批次打分出错: {e}")
        log_message(traceback.format_exc())
        return False

# 创建置顶窗口
root = tk.Tk()
root.title("自动批改助手")
root.geometry("500x450")
root.attributes("-topmost", True)
root.resizable(True, True)

# 添加状态标签
status_label = tk.Label(root, text="状态：准备就绪", font=("Arial", 10))
status_label.pack(pady=5)

# 添加说明
instruction = tk.Label(root, text="请先手动打开浏览器并导航到批改页面，然后选择分数区域和提交按钮区域", font=("Arial", 8, "bold"))
instruction.pack(pady=3)

# 添加区域选择按钮框架
region_frame = tk.Frame(root)
region_frame.pack(pady=5)

# 添加选择分数输入区域按钮
btn_select_score = tk.Button(region_frame, text="1.选择分数区域", command=select_score_area, bg="blue", fg="white", width=15, height=1)
btn_select_score.pack(side=tk.LEFT, padx=5)

# 添加选择提交按钮区域按钮
btn_select_submit = tk.Button(region_frame, text="2.选择提交按钮区域", command=select_submit_area, bg="blue", fg="white", width=15, height=1)
btn_select_submit.pack(side=tk.LEFT, padx=5)

# 添加选择答案文本区域按钮
btn_select_answer = tk.Button(region_frame, text="3.选择答案区域", command=select_answer_area, bg="blue", fg="white", width=15, height=1)
btn_select_answer.pack(side=tk.LEFT, padx=5)

# 添加操作按钮框架
button_frame = tk.Frame(root)
button_frame.pack(pady=5)

# 添加开始/暂停按钮
btn_toggle = tk.Button(button_frame, text="开始", command=toggle_grading, bg="green", fg="white", width=10, height=1)
btn_toggle.pack(side=tk.LEFT, padx=5)

# 添加配置按钮
config_btn = tk.Button(button_frame, text="配置", command=show_config_window, bg="orange", fg="white", width=10, height=1)
config_btn.pack(side=tk.LEFT, padx=5)

# 添加日志区域
log_frame = tk.Frame(root)
log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

log_label = tk.Label(log_frame, text="运行日志:")
log_label.pack(anchor=tk.W)

log_area = scrolledtext.ScrolledText(log_frame, height=15)
log_area.pack(fill=tk.BOTH, expand=True)
log_area.config(state="disabled")

# 添加多题目打分模式框架
multi_frame = tk.Frame(root)
multi_frame.pack(pady=5)

# 添加选择分数输入框区域按钮
btn_select_score_input = tk.Button(multi_frame, text="选择分数输入框", command=select_score_input_area, 
                                 bg="blue", fg="white", width=12, height=1)
btn_select_score_input.pack(side=tk.LEFT, padx=5)

# 添加选择最终提交按钮区域按钮
btn_select_final_submit = tk.Button(multi_frame, text="选择最终提交按钮", command=select_final_submit_area, 
                                  bg="blue", fg="white", width=12, height=1)
btn_select_final_submit.pack(side=tk.LEFT, padx=5)

# 添加滚动量设置按钮
btn_set_scroll = tk.Button(multi_frame, text="设置滚动量", command=set_scroll_amount, 
                          bg="blue", fg="white", width=10, height=1)
btn_set_scroll.pack(side=tk.LEFT, padx=5)

# 添加选择题号区域按钮
btn_select_question = tk.Button(multi_frame, text="选择题号区域", command=select_question_number_area, 
                               bg="blue", fg="white", width=12, height=1)
btn_select_question.pack(side=tk.LEFT, padx=5)

# 修改开始多题目按钮，使用基于题号的方法
btn_start_multipage = tk.Button(multi_frame, text="开始多题目打分", command=toggle_improved_scoring, 
                              bg="purple", fg="white", width=12, height=1)
btn_start_multipage.pack(side=tk.LEFT, padx=5)

# 关闭窗口时清理资源
root.protocol("WM_DELETE_WINDOW", on_closing)

# 启动时记录
log_message("程序已启动")
log_message("请选择操作模式: 1.单题目批改 2.多题目滚动批改")

# 在导入 pytesseract 后添加
if OCR_AVAILABLE:
    # 指定 Tesseract 路径
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# 启动GUI主循环
root.mainloop()
