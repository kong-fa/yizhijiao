"""
自动批改助手配置文件
"""

# 默认配置
DEFAULT_CONFIG = {
    "score": "100",                # 默认评分
    "wait_time": 5,                # 页面刷新等待时间（秒）
    "after_submit_wait": 1,        # 提交后等待时间（秒）
    "deepseek_api_key": "",        # DeepSeek API密钥
    "use_ai_scoring": False,       # 是否使用AI评分
    "ocr_language": "chi_sim+eng", # OCR识别语言（中文简体+英文）
    "question_content": "",        # 题目内容
    "multi_question_score": "10",  # 多题目默认分数
    "randomize_scores": False,     # 分数是否随机微调
    "use_smart_scrolling": True,   # 是否使用智能滚动
    "scroll_step_size": -10        # 智能滚动的步长
}

# 当前配置（初始化为默认值）
config = DEFAULT_CONFIG.copy()

def save_config():
    """保存配置到文件"""
    import json
    import os
    
    try:
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"保存配置失败: {e}")
        return False

def load_config():
    """从文件加载配置"""
    import json
    import os
    
    global config
    
    try:
        if os.path.exists("config.json"):
            with open("config.json", "r", encoding="utf-8") as f:
                loaded_config = json.load(f)
                # 更新配置，但保留默认值中存在而加载的配置中不存在的项
                for key in DEFAULT_CONFIG:
                    if key in loaded_config:
                        config[key] = loaded_config[key]
            return True
        else:
            # 如果配置文件不存在，创建默认配置
            save_config()
            return False
    except Exception as e:
        print(f"加载配置失败: {e}")
        return False

# 初始化时加载配置
load_config() 