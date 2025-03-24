"""
自动评分模块 - 使用AI模型自动评估题目答案的得分
"""

import requests
import json
import time
import traceback

class DeepSeekScorer:
    """使用DeepSeek模型的自动评分工具"""
    
    def __init__(self, api_key=None):
        """初始化评分器"""
        self.api_key = api_key
        self.api_url = "https://api.deepseek.com/v1/chat/completions"  # DeepSeek API地址
        
        # 如果没有提供API密钥，尝试从配置文件加载
        if not self.api_key:
            try:
                # 尝试从配置文件加载API密钥
                from config import config
                if "deepseek_api_key" in config:
                    self.api_key = config["deepseek_api_key"]
            except (ImportError, KeyError):
                pass
    
    def score_answer(self, answer_text, expected_score=None, question_text=None):
        """
        使用AI评估答案得分
        
        参数:
            answer_text (str): 学生答案文本
            expected_score (int, optional): 期望分数，用于提示AI
            question_text (str, optional): 题目内容
            
        返回:
            tuple: (分数, 评分理由)
        """
        if not self.api_key:
            return None, "未设置API密钥，无法使用AI评分"
            
        try:
            # 构建提示词
            prompt = f"""你是一个原生Java安卓教学助手，请根据以下学生答案给出合理的分数（60-100分）。
            """
            
            # 如果有题目内容，则添加到提示中
            if question_text and question_text.strip():
                prompt += f"""
题目:
{question_text}

"""
            
            prompt += f"""学生答案:
{answer_text}

请只返回一个数字分数（60-100之间的整数），不要有任何其他文字。例如：95"""

            if expected_score:
                prompt += f"\n\n参考：该题目预期分数为{expected_score}分左右。"
            
            # 构建API请求
            payload = {
                "model": "Pro/deepseek-ai/DeepSeek-V3",  # DeepSeek聊天模型
                "messages": [
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                "stream": False,
                "max_tokens": 50,  # 只需要短回复
                "temperature": 0.1,  # 较低的随机性，保证评分稳定
                "top_p": 0.9
            }
            
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }
            
            # 发送请求
            response = requests.post(
                self.api_url, 
                json=payload, 
                headers=headers,
                timeout=30
            )
            
            # 解析响应
            if response.status_code == 200:
                result = response.json()
                score_text = result.get("choices", [{}])[0].get("message", {}).get("content", "")
                
                # 提取纯数字
                import re
                score_match = re.search(r'(\d+)', score_text)
                if score_match:
                    score = int(score_match.group(1))
                    # 确保分数在0-100范围内
                    score = max(0, min(100, score))
                    return score, "AI自动评分"
                else:
                    return None, f"无法从AI回复中提取分数：{score_text}"
            else:
                return None, f"AI评分请求失败：HTTP {response.status_code}, {response.text}"
                
        except Exception as e:
            error_details = traceback.format_exc()
            return None, f"AI评分过程出错: {str(e)}\n{error_details}"

# 简单封装的评分函数
def ai_score(answer_text, expected_score=None, question_text=None, api_key=None):
    """
    使用AI评估答案得分的简便函数
    
    参数:
        answer_text (str): 学生答案文本
        expected_score (int, optional): 期望分数
        question_text (str, optional): 题目内容
        api_key (str, optional): DeepSeek API密钥
        
    返回:
        int: 分数（0-100）或None（评分失败）
    """
    scorer = DeepSeekScorer(api_key)
    score, reason = scorer.score_answer(answer_text, expected_score, question_text)
    return score 