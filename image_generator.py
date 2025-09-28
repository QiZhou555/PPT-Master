import requests
import time
import json
from PIL import Image
from io import BytesIO
import os
import traceback

def generate_image(prompt, output_path):
    """
    使用ModelScope API生成图片
    
    Args:
        prompt: 图片生成提示词
        output_path: 保存图片的路径
        
    Returns:
        bool: 生成成功返回True，失败返回False
    """
    try:
        base_url = 'https://api-inference.modelscope.cn/'
        api_key = "ms-4b457ae8-4cfd-4504-8ec2-8dc2fb930454" # ModelScope Token

        common_headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }

        # 发送生成图片的请求
        response = requests.post(
            f"{base_url}v1/images/generations",
            headers={**common_headers, "X-ModelScope-Async-Mode": "true"},
            data=json.dumps({
                "model": "Qwen/Qwen-Image", # ModelScope Model-Id
                "prompt": prompt
            }, ensure_ascii=False).encode('utf-8')
        )

        response.raise_for_status()
        task_id = response.json()["task_id"]
        print(f"开始生成图片: {prompt}")

        # 轮询任务状态
        max_retries = 30  # 最多尝试30次，每次间隔5秒，最多等待150秒
        retry_count = 0
        
        while retry_count < max_retries:
            result = requests.get(
                f"{base_url}v1/tasks/{task_id}",
                headers={**common_headers, "X-ModelScope-Task-Type": "image_generation"},
            )
            result.raise_for_status()
            data = result.json()

            if data["task_status"] == "SUCCEED":
                # 下载并保存图片
                image_response = requests.get(data["output_images"][0])
                image_response.raise_for_status()
                image = Image.open(BytesIO(image_response.content))
                
                # 确保目录存在
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                
                # 保存图片
                image.save(output_path)
                print(f"图片生成成功并保存到: {output_path}")
                return True
            elif data["task_status"] == "FAILED":
                print(f"图片生成失败: {data.get('error_msg', '未知错误')}")
                return False

            retry_count += 1
            print(f"等待图片生成... ({retry_count}/{max_retries})")
            time.sleep(5)
            
        print("图片生成超时")
        return False
        
    except Exception as e:
        print(f"图片生成过程中发生错误: {str(e)}")
        print(f"错误详情: {traceback.format_exc()}")
        return False


def generate_image_prompts(topic, sections):
    """
    根据PPT主题和章节内容生成图片提示词
    
    Args:
        topic: PPT主题
        sections: PPT章节列表
        
    Returns:
        list: 每个章节对应的图片提示词列表
    """
    prompts = []
    
    # 为标题页生成提示词
    title_prompt = f"专业商务风格的{topic}主题封面图，包含相关视觉元素，简约大气，适合PPT封面"
    prompts.append({"type": "title", "prompt": title_prompt, "section_index": -1})
    
    # 为每个章节生成提示词
    for i, section in enumerate(sections):
        section_title = section['title']
        section_content = section['content']
        
        # 根据章节标题和内容生成相关的提示词
        # 这里可以根据不同章节类型生成不同风格的提示词
        if "概述" in section_title or "简介" in section_title:
            prompt = f"专业风格的{topic}概述图，包含相关图标和简洁图形，适合商务PPT"
        elif "概念" in section_title or "理论" in section_title:
            prompt = f"关于{section_title}的信息图表，展示核心概念，使用简洁的图标和文字，商务风格"
        elif "技术" in section_title or "方法" in section_title:
            prompt = f"展示{topic}中{section_title}的技术插图，包含流程图或图表，专业风格"
        elif "案例" in section_title or "应用" in section_title:
            prompt = f"{topic}的{section_title}场景图，展示实际应用效果，专业商务风格"
        elif "未来" in section_title or "展望" in section_title:
            prompt = f"{topic}的未来发展趋势图，使用现代科技风格，适合PPT演示"
        else:
            prompt = f"与{topic}中{section_title}相关的专业插图，适合商务PPT使用"
        
        prompts.append({"type": "section", "prompt": prompt, "section_index": i})
    
    return prompts