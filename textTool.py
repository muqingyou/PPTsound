import os
import re
from openai import OpenAI

def analyze_kp_lines(input_file, output_file):
    # 初始化OpenAI客户端
    client = OpenAI(
        api_key=os.getenv("DASHSCOPE_API_KEY"),
        base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
    )
    
    # 读取输入文件
    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # 处理每一行
    processed_lines = []
    
    for line in lines:
        line = line.strip()
        if not line:
            processed_lines.append('')
            continue
        
        # 检查是否是[KP]开头的行
        if line.startswith('[KP]'):
            # 提取KP的内容
            kp_content = line[4:].strip()
            
            # 构建分析提示
            analysis_prompt = f"""请分析以下文本，判断[KP]是在进行故事叙述，还是在扮演NPC说话：

{kp_content}

如果是故事叙述，请保持原样；如果是扮演NPC说话，请识别NPC的名字。

输出格式：
- 类型：叙述/扮演
- NPC名字：如果是扮演，请输出NPC名字；如果是叙述，请输出"无"
"""
            
            # 调用API进行分析
            completion = client.chat.completions.create(
                model="tongyi-xiaomi-analysis-flash",
                messages=[
                    {
                        'role': 'user',
                        'content': analysis_prompt
                    }
                ],
                temperature=0,
                extra_body={
                    "top_k": 1
                }
            )
            
            # 解析API响应
            result = completion.choices[0].message.content
            print(f"分析结果: {result}")
            
            # 提取分析结果
            type_match = re.search(r'类型：(叙述|扮演)', result)
            npc_match = re.search(r'NPC名字：(.*)', result)
            
            if type_match and npc_match:
                analysis_type = type_match.group(1)
                npc_name = npc_match.group(1).strip()
                
                # 检查是否是有效的NPC名字（不是"无"或包含"无"的解释）
                if analysis_type == "扮演" and npc_name != "无" and not npc_name.startswith("无") and "无（" not in npc_name:
                    # 替换为NPC名字
                    processed_line = f"[{npc_name}] {kp_content}"
                else:
                    # 保持原样
                    processed_line = line
            else:
                # 分析失败，保持原样
                processed_line = line
        else:
            # 非KP行，保持原样
            processed_line = line
        
        processed_lines.append(processed_line)
    
    # 写入输出文件
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(processed_lines))
    
    print(f"处理完成，结果已写入 {output_file}")

if __name__ == "__main__":
    input_file = "e:\\VSCode\\replay\\textToBeCleaned.txt"
    output_file = "e:\\VSCode\\replay\\textCleaned.txt"
    analyze_kp_lines(input_file, output_file)