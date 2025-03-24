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

# 全局变量控制运行状态
running = False

# 用于存储两个区域的坐标
score_area_rect = None       # 分数输入区域
submit_area_rect = None      # 提交按钮区域

# 记录区域是否已选择
score_area_selected = False
submit_area_selected = False

def toggle_grading():
    global running
    running = not running
    
    if running:
        # 检查是否已选择所有必要区域
        if not (score_area_selected and submit_area_selected):
            messagebox.showwarning("警告", "请先选择分数区域和提交按钮区域")
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
        
        # 保存区域截图
        cv2.imwrite("score_area.png", region_img)
        log_message(f"已设置分数区域: 左={left}, 上={top}, 右={right}, 下={bottom}")
        log_message("已保存分数区域截图到 score_area.png")
        
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
        
        # 保存区域截图
        cv2.imwrite("submit_area.png", region_img)
        log_message(f"已设置提交按钮区域: 左={left}, 上={top}, 右={right}, 下={bottom}")
        log_message("已保存提交按钮区域截图到 submit_area.png")
        
        # 更新按钮状态
        btn_select_submit.config(bg="green")
        
        # 显示截图预览
        show_preview(region_img, "提交按钮区域预览")
        
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
    """在分数输入区域输入100分"""
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
        
        # 输入100分
        pyautogui.write("100")
        log_message("已输入100分")
        
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

def start_grading():
    global running
    
    try:
        # 循环执行批改，直到用户暂停
        while running:
            try:
                # 第一步：输入100分
                log_message("尝试输入100分...")
                if enter_score():
                    status_label.config(text="状态：已输入100分")
                    time.sleep(1)
                else:
                    log_message("未能输入分数，尝试直接提交")
                
                # 第二步：查找并点击提交按钮
                log_message("点击提交按钮...")
                if find_and_click_submit_button():
                    log_message("已点击提交按钮，完成一次批改")
                    status_label.config(text="状态：已提交，等待下一个...")
                    time.sleep(5)  # 将等待时间设为5秒，确保页面完全刷新
                    log_message("页面刷新等待完成，准备处理下一个")
                else:
                    log_message("未能点击提交按钮，跳过当前项目")
                    time.sleep(2)
                    continue
                
            except Exception as e:
                error_msg = f"批改过程中发生错误: {str(e)}"
                log_message(error_msg)
                log_message(traceback.format_exc())
                status_label.config(text=f"状态：发生错误")
                time.sleep(2)
            
            # 每次批改完成后暂停一下，避免过快操作
            time.sleep(1)
    
    except Exception as e:
        error_msg = f"发生严重错误：{e}"
        log_message(error_msg)
        log_message(traceback.format_exc())
        messagebox.showerror("错误", error_msg)
        btn_toggle.config(text="开始", bg="green")
        status_label.config(text="状态：已暂停")
        running = False

def on_closing():
    global running
    running = False
    root.destroy()

# 创建置顶窗口
root = tk.Tk()
root.title("自动批改助手")
root.geometry("500x400")
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

# 添加操作按钮框架
button_frame = tk.Frame(root)
button_frame.pack(pady=5)

# 添加开始/暂停按钮
btn_toggle = tk.Button(button_frame, text="开始", command=toggle_grading, bg="green", fg="white", width=10, height=1)
btn_toggle.pack(side=tk.LEFT, padx=5)

# 添加日志区域
log_frame = tk.Frame(root)
log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

log_label = tk.Label(log_frame, text="运行日志:")
log_label.pack(anchor=tk.W)

log_area = scrolledtext.ScrolledText(log_frame, height=15)
log_area.pack(fill=tk.BOTH, expand=True)
log_area.config(state="disabled")

# 关闭窗口时清理资源
root.protocol("WM_DELETE_WINDOW", on_closing)

# 启动时记录
log_message("程序已启动")
log_message("请手动打开浏览器并导航到批改页面，然后选择分数区域和提交按钮区域")

# 启动GUI主循环
root.mainloop()
