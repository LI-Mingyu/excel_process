from typing import Optional, List
import os
import json
from dotenv import load_dotenv
from llm_sandbox import SandboxSession
from openai import OpenAI

# 加载 .env 文件中的环境变量
load_dotenv()

# 初始化 OpenAI 客户端
client = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY"),
    base_url=os.getenv("OPENAI_API_BASE")
)

def run_code(lang: str, code: str, libraries: Optional[List] = None) -> str:
    """
    Run code in a sandboxed environment.
    :param lang: The language of the code.
    :param code: The code to run.
    :param libraries: The libraries to use, it is optional.
    :return: The output of the code.
    """
    with SandboxSession(lang=lang, verbose=False) as session:
        result = session.run(code, libraries).text
        # 确保结果不为空
        if not result or result.strip() == '':
            return "执行成功，但没有输出结果。请尝试添加print语句来显示结果。"
        return result

def copy_file_to_sandbox(local_path: str, sandbox_path: str) -> str:
    """
    将本地文件复制到沙盒环境中
    :param local_path: 本地文件路径
    :param sandbox_path: 沙盒中的目标路径
    :return: 操作结果
    """
    try:
        with SandboxSession(lang="python", keep_template=True) as session:
            session.copy_to_runtime(local_path, sandbox_path)
        return f"文件已成功复制到沙盒: {local_path} -> {sandbox_path}"
    except Exception as e:
        return f"复制文件失败: {str(e)}"

def copy_file_from_sandbox(sandbox_path: str, local_path: str) -> str:
    """
    从沙盒环境复制文件到本地
    :param sandbox_path: 沙盒中的文件路径
    :param local_path: 本地目标路径
    :return: 操作结果
    """
    try:
        with SandboxSession(lang="python", keep_template=True) as session:
            session.copy_from_runtime(sandbox_path, local_path)
        return f"文件已成功从沙盒复制: {sandbox_path} -> {local_path}"
    except Exception as e:
        return f"复制文件失败: {str(e)}"

def run_agent(user_input: str, excel_file_path: Optional[str] = None):
    """
    使用 OpenAI SDK 调用大模型并执行工具调用
    :param user_input: 用户输入
    :param excel_file_path: Excel 文件路径（可选）
    :return: 模型的最终输出
    """
    # 定义工具
    tools = [
        {
            "type": "function",
            "function": {
                "name": "run_code",
                "description": "Run code in a sandboxed environment.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "lang": {
                            "type": "string",
                            "description": "The language of the code. Must be one of ['python', 'java', 'javascript', 'cpp', 'go', 'ruby']"
                        },
                        "code": {
                            "type": "string",
                            "description": "The code to run."
                        },
                        "libraries": {
                            "type": "array",
                            "items": {
                                "type": "string"
                            },
                            "description": "The libraries to use, it is optional."
                        }
                    },
                    "required": ["lang", "code"]
                }
            }
        },
        {
            "type": "function",
            "function": {
                "name": "copy_file_to_sandbox",
                "description": "Copy a file from local to sandbox environment.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "local_path": {
                            "type": "string",
                            "description": "Local file path."
                        },
                        "sandbox_path": {
                            "type": "string",
                            "description": "Destination path in sandbox."
                        }
                    },
                    "required": ["local_path", "sandbox_path"]
                }
            }
        },
        {
            "type": "function",
            "function": {
                "name": "copy_file_from_sandbox",
                "description": "Copy a file from sandbox to local environment.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "sandbox_path": {
                            "type": "string",
                            "description": "File path in sandbox."
                        },
                        "local_path": {
                            "type": "string",
                            "description": "Destination local path."
                        }
                    },
                    "required": ["sandbox_path", "local_path"]
                }
            }
        }
    ]

    # 准备用户输入
    messages = [{"role": "user", "content": user_input}]
    
    # 如果提供了Excel文件路径，先将其复制到沙盒中
    if excel_file_path:
        # 修改为只提供目标目录路径，不包含文件名
        sandbox_dir = "/sandbox"
        # 从原始路径中提取文件名
        excel_filename = os.path.basename(excel_file_path)
        # 构建完整的沙盒路径（仅用于告知模型）
        sandbox_excel_path = os.path.join(sandbox_dir, excel_filename)
        
        # 调用复制函数时只传递目录路径
        copy_result = copy_file_to_sandbox(excel_file_path, sandbox_dir)
        print(copy_result)
        
        # 告诉模型Excel文件的位置
        messages[0]["content"] += f"\n\nExcel文件已上传到沙盒环境，路径为: {sandbox_excel_path}"
    
    # 打印用户输入（prompt）
    print("\n===== 用户输入 (Prompt) =====")
    print(messages[0]["content"])
    print("===========================\n")
    
    # 循环处理工具调用
    generated_files = []  # 用于跟踪生成的文件
    
    while True:
        response = client.chat.completions.create(
            model="qwen-max",  # 可以根据需要更换为其他模型
            messages=messages,
            tools=tools,
            tool_choice="auto"
        )
        
        response_message = response.choices[0].message
        messages.append(response_message)
        
        # 打印大模型的输出
        print("\n===== 大模型输出 =====")
        print(f"角色: {response_message.role}")
        print(f"内容: {response_message.content}")
        
        # 如果有工具调用，也打印出来
        if hasattr(response_message, 'tool_calls') and response_message.tool_calls:
            print("\n工具调用:")
            for tool_call in response_message.tool_calls:
                print(f"  - 函数: {tool_call.function.name}")
                print(f"  - 参数: {tool_call.function.arguments}")
        print("===================\n")
        
        # 检查是否有工具调用
        if hasattr(response_message, 'tool_calls') and response_message.tool_calls:
            # 处理工具调用
            for tool_call in response_message.tool_calls:
                function_name = tool_call.function.name
                function_args = json.loads(tool_call.function.arguments)
                
                if function_name == "run_code":
                    # 解析参数
                    lang = function_args.get("lang")
                    code = function_args.get("code")
                    libraries = function_args.get("libraries")
                    
                    # 执行代码
                    result = run_code(lang, code, libraries)
                elif function_name == "copy_file_to_sandbox":
                    local_path = function_args.get("local_path")
                    sandbox_path = function_args.get("sandbox_path")
                    result = copy_file_to_sandbox(local_path, sandbox_path)
                elif function_name == "copy_file_from_sandbox":
                    sandbox_path = function_args.get("sandbox_path")
                    local_path = function_args.get("local_path")
                    result = copy_file_from_sandbox(sandbox_path, local_path)
                
                # 将工具执行结果添加到消息中
                tool_message = {
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "name": function_name,
                    "content": result
                }
                messages.append(tool_message)
                
                # 打印工具执行结果
                print("\n===== 工具执行结果 =====")
                print(f"工具: {function_name}")
                print(f"结果: {result}")
                print("=======================\n")
            
            # 继续对话
            continue
        else:
            # 没有工具调用，处理生成的文件并返回最终结果
            final_result = response_message.content
            
            # 如果有生成的文件，将它们复制到本地并添加到结果中
            if generated_files:
                print("\n===== 处理生成的文件 =====")
                local_files = []
                
                for sandbox_file in generated_files:
                    # 提取文件名
                    file_name = os.path.basename(sandbox_file)
                    # 创建本地目标路径
                    local_file = os.path.join(os.getcwd(), "output_" + file_name)
                    
                    # 复制文件到本地
                    try:
                        copy_result = copy_file_from_sandbox(sandbox_file, local_file)
                        print(copy_result)
                        local_files.append(local_file)
                    except Exception as e:
                        print(f"复制文件失败: {str(e)}")
                
                # 将本地文件路径添加到结果中
                if local_files:
                    file_paths_text = "\n\n生成的文件已复制到本地:\n" + "\n".join(local_files)
                    final_result += file_paths_text
                
                print("===========================\n")
            
            return final_result

if __name__ == "__main__":
    print("Excel 文件处理工具")
    print("-" * 50)
    
    # 获取用户输入的Excel文件路径
    excel_path = input("请输入要处理的Excel文件路径 (留空则不指定文件): ")
    if excel_path and not os.path.exists(excel_path):
        print(f"错误: 文件 '{excel_path}' 不存在")
        exit(1)
    
    # 获取用户的处理需求
    user_request = input("请输入您的Excel处理需求: ")
    
    # 执行处理
    print("\n正在处理，请稍候...\n")
    result = run_agent(user_request, excel_path if excel_path else None)
    
    print("\n处理结果:")
    print("-" * 50)
    print(result)