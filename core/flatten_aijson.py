from typing import Dict, List, Tuple
import json
import os
import re
from collections import OrderedDict
from pathlib import Path


class JsonFlattener:
    """JSON数据展平处理器
    
    属性:
        input_pattern (str): JSON文件匹配模式
        output_dir (str): 输出目录路径
    """
    
    def __init__(self, input_pattern: str = 'ai_note2ca_card/data/input/json/ai/*.json',
                 output_dir: str = 'ai_note2ca_card/data/input/json/temp/'):
        self.input_pattern = input_pattern
        self.output_dir = output_dir
        Path(self.output_dir).mkdir(parents=True, exist_ok=True)

    def flatten_json(self, nested_json: Dict) -> List[str]:
        """递归展平JSON嵌套结构
        Args:
            nested_json: 要处理的嵌套JSON字典
        Returns:
            展平后的字符串列表，格式为"路径##组件^值"
        """
        def process_value(value, path_components: List[str], result: List[str]):
            if isinstance(value, dict):
                for k, v in value.items():
                    process_value(v, path_components + [k], result)
            else:
                full_path = '##'.join(path_components)
                result.append(f"{full_path}^{value}")
        
        result = []
        for main_key, main_value in nested_json.items():
            process_value(main_value, [main_key], result)
        return result

    def generate_kd_sets(self, flattened_data: List[str]) -> Tuple[OrderedDict, OrderedDict, OrderedDict]:
        """生成三个中间数据集
        Args:
            flattened_data: 展平后的字符串列表
        Returns:
            包含三个有序字典的元组 (kd1, kd2, kd3)
        """
        # 生成kd1（L前缀）
        l_keys = OrderedDict()
        for line in flattened_data:
            main_key = line.split('##')[0]
            if main_key not in l_keys:
                l_keys[main_key] = f"L{len(l_keys)+1}"
        
        # 生成kd2（H前缀）
        unique_h = set()
        for line in flattened_data:
            parts = line.split('^')[0].split('##')[1:]  # 去掉第一个元素
            # 截取前两位路径组件
            path_key = '##'.join(parts)
            unique_h.add(path_key)
        
        # 字典排序并创建有序映射
        def custom_sort_key(s):
            parts = s.split('##')
            # 第一部分字典序
            part1 = parts[0]
            
            # 第二部分提取数字
            part2_num = int(re.search(r'\d+', parts[1]).group()) if len(parts) > 1 else 0
            
            # 第三部分字典序
            part3 = parts[2] if len(parts) > 2 else ''
            
            return (part1, part2_num, part3)

        sorted_h = sorted(unique_h, key=custom_sort_key)
        h_keys = OrderedDict()
        for idx, key in enumerate(sorted_h, 1):
            h_keys[key] = f"H{idx}"
        
        # 生成kd3
        kd3 = OrderedDict()
        for line in flattened_data:
            parts = line.split('^')
            path_part = parts[0]
            value = parts[1]
            
            l_part = path_part.split('##')[0]
            h_part = '##'.join(path_part.split('##')[1:])
            
            l_id = l_keys[l_part]
            h_id = h_keys[h_part]
            
            if l_id not in kd3:
                kd3[l_id] = OrderedDict()
            kd3[l_id][h_id] = value
        
        return l_keys, h_keys, kd3

    def process_files(self) -> Dict:
        """处理多个JSON文件并返回最终结果
        Returns:
            处理后的有序字典结果
        """
        import glob
        try:
            json_files = glob.glob(self.input_pattern)
            if not json_files:
                raise FileNotFoundError(f"未找到匹配的JSON文件：{self.input_pattern}")

            all_flattened = []
            
            for file_path in json_files:
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f, object_pairs_hook=OrderedDict)
                        all_flattened.extend(self.flatten_json(data))
                except Exception as e:
                    print(f"处理文件 {os.path.basename(file_path)} 时出错：{str(e)}")
                    continue

            kd1, kd2, kd3 = self.generate_kd_sets(all_flattened)
            
            final_output = OrderedDict()
            final_output["str0"] = {v:k for k,v in kd1.items()}
            final_output["str1"] = {v:k.split('##') for k,v in kd2.items()}
            final_output["str2"] = kd3

            # 计算str1Num
            str1_num = 0
            if final_output.get("str1"):
                first_key = next(iter(final_output["str1"].keys()), None)
                if first_key and isinstance(final_output["str1"][first_key], list):
                    str1_num = len(final_output["str1"][first_key])

            # 添加baseinfo并保持顺序
            new_output = OrderedDict()
            new_output["baseinfo"] = {"str1Num": str1_num}
            new_output.update(final_output)
            
            # 保存结果
            output_path = os.path.join(self.output_dir, 'flattened_output.json')
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(new_output, f, ensure_ascii=False, indent=2)
            
            return new_output

        except Exception as e:
            print(f"处理过程中发生错误：{str(e)}")
            raise

def main():
    try:
        processor = JsonFlattener(
            input_pattern='ai_note2ca_card/data/input/json/ai/*.json',
            output_dir='ai_note2ca_card/data/input/json/temp/'
        )
        result = processor.process_files()
        print(f"转换完成，结果已保存到：{processor.output_dir}")
        return result
    except Exception as e:
        print(f"程序运行出错：{str(e)}")
        raise

if __name__ == "__main__":
    main()
