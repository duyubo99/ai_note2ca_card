import pandas as pd
import os
import json
import time
from genText2 import genText


gmyx_json = {
  # 基础信息
  '品牌名称': '',  
  '产品线': '',       # 产品型号+适用阶段（如L码夜用型）
  '内容类型': '',      # 新增★细分类型：产品测评/剧情种草/专家背书
  
  # 品牌信任指标  
  '官方账号发布': False, 
  '行业KOL标识': False,
  '账号粉丝量级': '', # ★补充账号垂直领域: 母婴/测评/剧情
  '权威认证提及': [], # ★FDA/三甲医院等认证资质引用

  # 产品功能解构  
  '核心痛点': [],     # ★如[红屁屁、侧漏、闷热] 
  '功能诉求': [],     # 细化到具体场景：「夜用防漏」「透气散热」
  '技术护城河': [],   # ★专利技术/独家研发（如热感显温技术）
  '适用群体特征': {}, # ★年龄段/体重/敏感肌等标签
  
  # 感知验证手段
  '效果可视化': [],   # ★热成像/吸水性实测等验证方式
  '对比实验': False,   # 是否呈现竞品对比
  
  # 消费决策因子
  '价格带标识': '',    # ★超高端>4元/片 | 高端3-4元 
  '节点营销标识': [],
  '用户证言密度': 0,   # 视频中消费者证言出现次数
  '赠品策略': [],      # ★试用装/组合优惠信息
}

result_list = {}

# 检查点文件路径
CHECKPOINT_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "result", "checkpoint.json")

# 获取当前文件的绝对路径
current_dir = os.path.dirname(os.path.abspath(__file__))
# 构建Excel文件的相对路径
excel_path = os.path.join(os.path.dirname(current_dir), "data", "抖音-618.xlsx")

# 确保结果目录存在
result_dir = os.path.join(os.path.dirname(current_dir), "result", "json")
os.makedirs(result_dir, exist_ok=True)

# 读取Excel文件
df = pd.read_excel(excel_path)

# 获取列名作为keys
keys = df.columns.tolist()
# print("表格列名(keys):", keys)
# print("\n" + "-" * 50 + "\n")


# 另一种更简洁的方法：直接将DataFrame转换为字典列表
# print("\n使用to_dict方法转换的结果:")
records = df.to_dict(orient='records')

# 加载检查点（如果存在）
start_index = 0
if os.path.exists(CHECKPOINT_FILE):
    try:
        with open(CHECKPOINT_FILE, 'r', encoding='utf-8') as f:
            checkpoint_data = json.load(f)
            start_index = checkpoint_data.get('last_processed_index', 0) + 1
            print(f"从检查点恢复，从索引 {start_index} 开始处理")
    except Exception as e:
        print(f"读取检查点文件失败: {e}")

# 创建或清空结果文件（如果从头开始）
if start_index == 0:
    with open(os.path.join(result_dir, "json1.txt"), "w", encoding="utf-8") as file:
        file.write("")

# 批量处理的大小
BATCH_SIZE = 3
batch_records = []
batch_indices = []

for i, record in enumerate(records[start_index:], start=start_index):
    print(f"第{i+1}行添加到批处理")
    batch_records.append(record)
    batch_indices.append(i)
    
    # 当累积了5条记录或者是最后一批数据时，调用大模型接口
    if len(batch_records) >= BATCH_SIZE or i == len(records) - 1:
        print(f"\n批量处理 {len(batch_records)} 条记录:")
        
        # 构建批量处理的提示
        batch_data = ""
        for idx, rec in enumerate(batch_records):
            batch_data += f"记录 {batch_indices[idx]+1}:\n{rec}\n\n"
        
        prompt = f"""
        请根据以下多条数据和json的key维度进行内容转换:

        原始数据:
        {batch_data}

        参考结构:
        {gmyx_json}

        输出格式要求：
        对每条记录，按以下格式输出，每条记录之间用两个换行符分隔（\n\n）:

        序号：[记录编号]
        品牌名称:xxx
        产品线:xxx
        ...

        \n\n

        序号：[下一条记录编号]
        品牌名称:xxx
        ...

        要求:
        1. 按购买意向重构数据维度，用于本地知识库构建
        2. 将原始数据中的信息映射到目标结构中的相应字段
        3. 保留原始数据中的关键信息
        4. 如果原始数据中没有对应目标结构的某些字段，可以留空或填写"未提及"
        5. 对于列表类型的字段，请提取相关的多个要点
        6. 确保每条记录之间有两个换行符分隔（\n\n）
        7. 每条记录的开头都要包含"序号：[记录编号]"

        请返回按照目标结构格式化后的数据。
        """

        # 调用大模型接口
        try:
            result = genText(prompt)
            
            # 确保结果中的记录之间有两个换行符分隔
            # 将结果写入文件
            with open(os.path.join(result_dir, "json1.txt"), "a", encoding="utf-8") as file:
                file.write(result)
                file.write("\n\n")  # 在批次之间额外添加换行
                print(f"批量处理的 {len(batch_records)} 条数据已保存到json1.txt")
            
            # 更新检查点到最后处理的索引
            with open(CHECKPOINT_FILE, 'w', encoding='utf-8') as f:
                json.dump({
                    'last_processed_index': batch_indices[-1],
                    'timestamp': time.time(),
                    'total_records': len(records)
                }, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            print(f"处理批次时出错: {e}")
            # 即使出错也保存检查点，下次从这里继续
            with open(CHECKPOINT_FILE, 'w', encoding='utf-8') as f:
                json.dump({
                    'last_processed_index': batch_indices[0] - 1,  # 出错的批次不算作已处理
                    'error_at_batch_start': batch_indices[0],
                    'error_message': str(e),
                    'timestamp': time.time(),
                    'total_records': len(records)
                }, f, ensure_ascii=False, indent=2)
            # 如果是严重错误，可以选择退出循环
            # break
        
        # 清空批处理列表，准备下一批
        batch_records = []
        batch_indices = []
