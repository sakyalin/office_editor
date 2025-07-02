import os
import sys
from pathlib import Path
from dotenv import load_dotenv

def load_env_variables():
    """
    加载环境变量，支持开发环境和打包后的环境
    """
    # 尝试从当前目录加载.env文件
    if load_dotenv():
        print("从当前目录加载了.env文件")
        return True
    
    # 如果当前目录没有.env文件，尝试从应用程序目录加载
    if getattr(sys, 'frozen', False):
        # 如果是打包后的应用程序
        application_path = Path(sys._MEIPASS)
        env_path = application_path / '.env'
        if env_path.exists():
            load_dotenv(env_path)
            print(f"从应用程序目录加载了.env文件: {env_path}")
            return True
    
    # 如果都没有找到.env文件，尝试从父目录加载
    parent_dir = Path.cwd().parent
    env_path = parent_dir / '.env'
    if env_path.exists():
        load_dotenv(env_path)
        print(f"从父目录加载了.env文件: {env_path}")
        return True
    
    print("警告: 未找到.env文件")
    return False

def get_api_key():
    """
    获取API密钥，优先从环境变量获取，如果没有则返回空字符串
    """
    # 加载环境变量
    load_env_variables()
    
    # 从环境变量获取API密钥
    api_key = os.environ.get("OPENAI_API_KEY", "")
    
    # 如果环境变量中没有API密钥，尝试从.env文件中直接读取
    if not api_key:
        try:
            # 尝试从当前目录的.env文件中读取
            env_path = Path('.env')
            if env_path.exists():
                with open(env_path, 'r') as f:
                    for line in f:
                        if line.startswith('OPENAI_API_KEY='):
                            api_key = line.strip().split('=', 1)[1]
                            # 去除可能的引号
                            api_key = api_key.strip('"\'')
                            break
        except Exception as e:
            print(f"读取.env文件出错: {e}")
    
    return api_key

def get_api_base_url():
    """
    获取API基础URL，优先从环境变量获取，如果没有则返回默认值
    """
    # 加载环境变量
    load_env_variables()
    
    # 从环境变量获取API基础URL，如果没有则使用默认值
    api_base_url = os.environ.get("OPENAI_API_BASE_URL", "https://api.gptsapi.net/v1/")
    
    # 如果环境变量中没有API基础URL，尝试从.env文件中直接读取
    if api_base_url == "https://api.gptsapi.net/v1/":
        try:
            # 尝试从当前目录的.env文件中读取
            env_path = Path('.env')
            if env_path.exists():
                with open(env_path, 'r') as f:
                    for line in f:
                        if line.startswith('OPENAI_API_BASE_URL='):
                            api_base_url = line.strip().split('=', 1)[1]
                            # 去除可能的引号
                            api_base_url = api_base_url.strip('"\'')
                            break
        except Exception as e:
            print(f"读取.env文件出错: {e}")
    
    return api_base_url

# 测试代码
if __name__ == "__main__":
    # 测试API密钥获取
    api_key = get_api_key()
    if api_key:
        # 打印部分API密钥，保护隐私
        masked_key = api_key[:4] + "*" * (len(api_key) - 8) + api_key[-4:]
        print(f"成功获取API密钥: {masked_key}")
    else:
        print("未找到API密钥")
    
    # 测试API基础URL获取
    api_base_url = get_api_base_url()
    print(f"成功获取API基础URL: {api_base_url}")