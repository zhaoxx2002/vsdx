import json
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

def load_vsdx_data(json_file):
    """加载VSDX结构数据"""
    with open(json_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def visualize_shape(ax, shape, parent_pos=(0, 0), level=0, color_map=None):
    """递归绘制形状及其子形状"""
    if color_map is None:
        color_map = {
            'Group': 'lightblue',
            'Shape': 'lightgreen'
        }
    
    # 获取位置信息
    try:
        x = float(shape['position']['x'])
        y = float(shape['position']['y'])
        width = float(shape['position']['width'])
        height = float(shape['position']['height'])
    except (ValueError, KeyError):
        print(f"警告: 形状 {shape.get('id', '未知')} 的位置信息无效")
        return
    
    # 确定颜色
    shape_type = shape.get('type', 'Shape')
    color = color_map.get(shape_type, 'gray')
    
    # 绘制矩形
    rect = patches.Rectangle(
        (x, -y),  # 注意Y轴反转
        width, 
        height,
        linewidth=1,
        edgecolor='black',
        facecolor=color,
        alpha=0.7
    )
    ax.add_patch(rect)
    
    # 添加标签
    name = shape.get('name', f"ID:{shape.get('id', '?')}")
    ax.text(x + width/2, -(y + height/2), name, 
            ha='center', va='center', fontsize=8, 
            color='black', fontweight='bold')
    
    # 递归处理子形状
    for child in shape.get('children', []):
        visualize_shape(ax, child, (x, y), level+1, color_map)

def visualize_vsdx_structure(json_file, output_file=None):
    """可视化VSDX结构"""
    data = load_vsdx_data(json_file)
    
    if not data:
        print("错误: 数据为空")
        return
    
    # 创建图形
    fig, ax = plt.subplots(figsize=(12, 10))
    
    # 设置颜色映射
    color_map = {
        'Group': 'lightblue',
        'Shape': 'lightgreen'
    }
    
    # 处理每个页面
    for page_data in data:
        page_index = page_data.get('page_index', '?')
        page_file = page_data.get('page_file', '?')
        
        print(f"处理页面 {page_index}: {page_file}")
        
        # 绘制每个形状
        for shape in page_data.get('shapes', []):
            visualize_shape(ax, shape, color_map=color_map)
    
    # 设置坐标轴
    ax.set_aspect('equal')
    ax.autoscale()
    ax.set_title(f"VSDX结构可视化")
    ax.set_xlabel('X坐标')
    ax.set_ylabel('Y坐标')
    
    # 添加图例
    legend_elements = [
        patches.Patch(facecolor='lightblue', edgecolor='black', label='Group'),
        patches.Patch(facecolor='lightgreen', edgecolor='black', label='Shape')
    ]
    ax.legend(handles=legend_elements, loc='upper right')
    
    # 保存或显示
    if output_file:
        plt.savefig(output_file, dpi=300, bbox_inches='tight')
        print(f"图像已保存至: {output_file}")
    else:
        plt.tight_layout()
        plt.show()

def visualize_connector_diagram(json_file, output_file=None):
    """绘制连接器示意图，专注于A1-A5连接器"""
    data = load_vsdx_data(json_file)
    
    if not data or not data[0].get('shapes'):
        print("错误: 数据为空或没有形状")
        return
    
    # 创建图形
    fig, ax = plt.subplots(figsize=(10, 8))
    
    # 查找主要连接器组
    main_group = None
    for shape in data[0].get('shapes', []):
        if shape.get('name') == 'A1' and shape.get('type') == 'Group':
            main_group = shape
            break
    
    if not main_group:
        print("错误: 未找到主连接器组")
        return
    
    # 提取A1-A5连接器
    connectors = []
    for child in main_group.get('children', []):
        if child.get('name') in ['A1', 'A2', 'A3', 'A4', 'A5'] and child.get('type') == 'Group':
            # 提取连接器信息
            connector_info = {
                'name': child.get('name'),
                'position': child.get('position'),
                'terminal': None
            }
            
            # 查找端子标签
            for terminal in child.get('children', []):
                if terminal.get('name') not in [child.get('name')]:
                    connector_info['terminal'] = terminal.get('name')
                    break
            
            connectors.append(connector_info)
    
    # 绘制连接器示意图
    connector_height = 0.5
    spacing = 0.2
    y_pos = 0
    
    for connector in connectors:
        name = connector.get('name')
        terminal = connector.get('terminal', 'N/A')
        
        # 绘制连接器
        rect = patches.Rectangle(
            (0, y_pos), 
            2, 
            connector_height,
            linewidth=1,
            edgecolor='black',
            facecolor='lightblue',
            alpha=0.7
        )
        ax.add_patch(rect)
        
        # 添加标签
        ax.text(0.5, y_pos + connector_height/2, name, 
                ha='center', va='center', fontsize=10, fontweight='bold')
        ax.text(1.5, y_pos + connector_height/2, terminal, 
                ha='center', va='center', fontsize=10)
        
        y_pos += connector_height + spacing
    
    # 添加组件信息
    component_info = []
    for child in main_group.get('children', []):
        if child.get('name') in ['Innenleuchte hinten links', '1K0_947_291', '2A1', 'W47 .1']:
            component_info.append(f"{child.get('name')}")
    
    if component_info:
        ax.text(1, y_pos + 0.5, '\n'.join(component_info), 
                ha='center', va='center', fontsize=10, 
                bbox=dict(facecolor='white', alpha=0.7, boxstyle='round'))
    
    # 设置坐标轴
    ax.set_xlim(0, 3)
    ax.set_ylim(0, y_pos + 1.5)
    ax.set_title("连接器示意图")
    ax.axis('off')
    
    # 保存或显示
    if output_file:
        plt.savefig(output_file, dpi=300, bbox_inches='tight')
        print(f"连接器示意图已保存至: {output_file}")
    else:
        plt.tight_layout()
        plt.show()

if __name__ == "__main__":
    json_file = "/Volumes/Data/vsdx/vsdx_structure_with_master_geometry.json"
    
    # 绘制完整结构图
    visualize_vsdx_structure(json_file, "/Volumes/Data/vsdx/vsdx_visualization.png")
    
    # 绘制连接器示意图
    visualize_connector_diagram(json_file, "/Volumes/Data/vsdx/connector_diagram.png")
    
    print("可视化完成！")