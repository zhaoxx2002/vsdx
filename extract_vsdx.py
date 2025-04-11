import os
import json
import argparse
from parse_vsdx_json import analyze_vsdx_structure_with_geometry
from draw import visualize_vsdx_structure, visualize_connector_diagram

def extract_vsdx_to_json(vsdx_file, output_json, debug=True):
    """从VSDX文件提取信息并导出为JSON"""
    try:
        # 检查文件是否存在
        if not os.path.exists(vsdx_file):
            raise FileNotFoundError(f"找不到VSDX文件：{vsdx_file}")
        
        print(f"正在分析VSDX文件：{vsdx_file}")
        # 分析VSDX文件结构
        result = analyze_vsdx_structure_with_geometry(vsdx_file, debug=debug)
        
        if not result:
            print("警告：未找到任何核心组件")
            return False
            
        # 保存为JSON文件
        with open(output_json, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        print(f"✅ 完成：VSDX结构已提取并保存到 {output_json}")
        return True
        
    except Exception as e:
        print(f"错误：{str(e)}")
        return False

def visualize_from_json(json_file, output_dir=None):
    """从JSON文件生成可视化图像"""
    try:
        if not os.path.exists(json_file):
            raise FileNotFoundError(f"找不到JSON文件：{json_file}")
        
        # 确定输出目录
        if output_dir is None:
            output_dir = os.path.dirname(json_file)
        
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 生成输出文件路径
        base_name = os.path.splitext(os.path.basename(json_file))[0]
        structure_output = os.path.join(output_dir, f"{base_name}_visualization.png")
        connector_output = os.path.join(output_dir, f"{base_name}_connectors.png")
        
        # 生成可视化图像
        print(f"正在生成结构可视化图像...")
        visualize_vsdx_structure(json_file, structure_output)
        
        print(f"正在生成连接器可视化图像...")
        visualize_connector_diagram(json_file, connector_output)
        
        print(f"✅ 完成：可视化图像已保存到 {output_dir}")
        return True
        
    except Exception as e:
        print(f"错误：{str(e)}")
        return False

def main():
    # 解析命令行参数
    parser = argparse.ArgumentParser(description="从VSDX文件提取信息并生成可视化图像")
    parser.add_argument("--vsdx", help="输入的VSDX文件路径")
    parser.add_argument("--json", help="输出的JSON文件路径或输入的JSON文件路径（如果不提取VSDX）")
    parser.add_argument("--output-dir", help="可视化图像的输出目录")
    parser.add_argument("--extract-only", action="store_true", help="仅提取VSDX到JSON，不生成可视化")
    parser.add_argument("--visualize-only", action="store_true", help="仅从JSON生成可视化，不提取VSDX")
    parser.add_argument("--debug", action="store_true", help="启用调试模式")
    
    args = parser.parse_args()
    
    # 默认JSON文件路径
    if args.json is None and args.vsdx is not None:
        args.json = os.path.splitext(args.vsdx)[0] + "_structure.json"
    
    # 执行操作
    if args.visualize_only:
        # 仅可视化
        if args.json is None:
            print("错误：需要指定JSON文件路径")
            return False
        return visualize_from_json(args.json, args.output_dir)
    elif args.extract_only:
        # 仅提取
        if args.vsdx is None:
            print("错误：需要指定VSDX文件路径")
            return False
        return extract_vsdx_to_json(args.vsdx, args.json, args.debug)
    else:
        # 提取并可视化
        if args.vsdx is None:
            print("错误：需要指定VSDX文件路径")
            return False
        
        # 先提取
        if extract_vsdx_to_json(args.vsdx, args.json, args.debug):
            # 再可视化
            return visualize_from_json(args.json, args.output_dir)
        return False

if __name__ == "__main__":
    main()