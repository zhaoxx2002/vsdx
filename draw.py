import json
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.lines as mlines
import numpy as np
from matplotlib.path import Path


def load_vsdx_data(json_file):
    """加载VSDX结构数据"""
    with open(json_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def visualize_shape(ax, shape, parent_pos=(0, 0), level=0, color_map=None):
    """递归绘制形状及其子形状"""
    if color_map is None:
        color_map = {
            'Group': 'lightblue',
            'Shape': 'lightgreen',
            'Line': 'red',
            'Dynamic connector': 'orange'
        }
    
    # 获取位置信息
    try:
        x = float(shape['position']['x'])
        y = float(shape['position']['y'])
        width = float(shape['position']['width'])
        height = float(shape['position']['height'])
        angle = float(shape['position'].get('angle', '0'))
        flip_x = shape['position'].get('flip_x', '0') == '1'
        flip_y = shape['position'].get('flip_y', '0') == '1'
    except (ValueError, KeyError):
        print(f"警告: 形状 {shape.get('id', '未知')} 的位置信息无效")
        return
    
    # 确定颜色
    shape_type = shape.get('type', 'Shape')
    color = color_map.get(shape_type, 'gray')
    
    # 检查是否有路径数据
    if 'path_data' in shape['position']:
        # 使用路径数据绘制自定义形状
        try:
            path_data = shape['position']['path_data']
            vertices = []
            codes = []
            
            for section in path_data:
                for i, point in enumerate(section):
                    if point['type'] == 'MoveTo':
                        vertices.append((float(point['X']), -float(point['Y'])))
                        codes.append(Path.MOVETO)
                    elif point['type'] == 'LineTo':
                        vertices.append((float(point['X']), -float(point['Y'])))
                        codes.append(Path.LINETO)
                    elif point['type'] in ['ArcTo', 'EllipticalArcTo']:
                        vertices.append((float(point['X']), -float(point['Y'])))
                        codes.append(Path.CURVE4)
            
            if vertices:
                # 闭合路径
                if len(vertices) > 1 and vertices[0] != vertices[-1]:
                    vertices.append(vertices[0])
                    codes.append(Path.CLOSEPOLY)
                
                path = Path(vertices, codes)
                patch = patches.PathPatch(path, facecolor=color, edgecolor='black', alpha=0.7)
                ax.add_patch(patch)
        except Exception as e:
            print(f"警告: 无法绘制自定义路径 {shape.get('id')}: {e}")
            # 回退到矩形
            rect = patches.Rectangle(
                (x, -y),  # 注意Y轴反转
                width, 
                height,
                linewidth=1,
                edgecolor='black',
                facecolor=color,
                alpha=0.7,
                angle=angle
            )
            ax.add_patch(rect)
    else:
        # 绘制矩形
        rect = patches.Rectangle(
            (x, -y),  # 注意Y轴反转
            width, 
            height,
            linewidth=1,
            edgecolor='black',
            facecolor=color,
            alpha=0.7,
            angle=angle
        )
        ax.add_patch(rect)
    
    # 添加标签
    name = shape.get('name', f"ID:{shape.get('id', '?')}")
    text = shape.get('text', '')
    display_text = text if text else name
    
    ax.text(x + width/2, -(y + height/2), display_text, 
            ha='center', va='center', fontsize=8, 
            color='black', fontweight='bold')
    
    # 绘制连接点
    for cp in shape.get('connection_points', []):
        try:
            cp_x = float(cp['x']) + x
            cp_y = -(float(cp['y']) + y)
            ax.plot(cp_x, cp_y, 'ro', markersize=3)
        except (ValueError, KeyError):
            pass
    
    # 递归处理子形状
    for child in shape.get('children', []):
        visualize_shape(ax, child, (x, y), level+1, color_map)

def visualize_connector(ax, connector):
    """绘制连接线"""
    points = connector.get('geometry_points', [])
    if len(points) < 2:
        return
    
    # 提取坐标点
    try:
        x_points = []
        y_points = []
        
        for point in points:
            if 'X' in point and 'Y' in point:
                x_points.append(float(point['X']))
                y_points.append(-float(point['Y']))  # 注意Y轴反转
            elif 'x' in point and 'y' in point:
                x_points.append(float(point['x']))
                y_points.append(-float(point['y']))  # 注意Y轴反转
        
        if len(x_points) < 2:
            return
        
        # 确定线条样式
        line_pattern = connector.get('line_pattern', '1')
        line_style = '-' if line_pattern == '1' else '--'
        
        # 确定线条颜色
        line_color = 'red'
        
        # 确定线条宽度
        try:
            line_weight = float(connector.get('line_weight', '1')) * 1.5
        except ValueError:
            line_weight = 1.5
        
        # 绘制线条
        line = mlines.Line2D(x_points, y_points, 
                            color=line_color, 
                            linestyle=line_style, 
                            linewidth=line_weight, 
                            marker='', 
                            alpha=0.8)
        ax.add_line(line)
        
        # 添加箭头（如果有）
        if connector.get('begin_arrow', '0') != '0':
            ax.arrow(x_points[0], y_points[0], 
                    (x_points[1] - x_points[0])*0.8, (y_points[1] - y_points[0])*0.8,
                    head_width=0.1, head_length=0.2, fc=line_color, ec=line_color)
        
        if connector.get('end_arrow', '0') != '0':
            ax.arrow(x_points[-2], y_points[-2], 
                    (x_points[-1] - x_points[-2])*0.8, (y_points[-1] - y_points[-2])*0.8,
                    head_width=0.1, head_length=0.2, fc=line_color, ec=line_color)
        
        # 添加连接线标签
        if connector.get('name'):
            mid_idx = len(x_points) // 2
            ax.text(x_points[mid_idx], y_points[mid_idx], 
                    connector.get('name'), 
                    ha='center', va='bottom', fontsize=7, 
                    bbox=dict(facecolor='white', alpha=0.7, boxstyle='round'))
    
    except Exception as e:
        print(f"警告: 无法绘制连接线 {connector.get('id', '?')}: {e}")

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
        'Shape': 'lightgreen',
        'Line': 'red',
        'Dynamic connector': 'orange'
    }
    
    # 处理每个页面
    for page_data in data:
        page_index = page_data.get('page_index', '?')
        page_file = page_data.get('page_file', '?')
        
        print(f"处理页面 {page_index}: {page_file}")
        
        # 绘制每个形状
        for shape in page_data.get('shapes', []):
            visualize_shape(ax, shape, color_map=color_map)
        
        # 绘制连接线
        for connector in page_data.get('connectors', []):
            visualize_connector(ax, connector)
    
    # 设置坐标轴
    ax.set_aspect('equal')
    ax.autoscale()
    ax.set_title(f"VSDX结构可视化")
    ax.set_xlabel('X坐标')
    ax.set_ylabel('Y坐标')
    
    # 添加图例
    legend_elements = [
        patches.Patch(facecolor='lightblue', edgecolor='black', label='Group'),
        patches.Patch(facecolor='lightgreen', edgecolor='black', label='Shape'),
        mlines.Line2D([], [], color='red', linestyle='-', linewidth=1.5, label='Connector')
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

def export_vsdx_to_json(vsdx_file, output_json):
    """从VSDX文件提取信息并导出为JSON"""
    try:
        from parse_vsdx_json import analyze_vsdx_structure_with_geometry
        
        # 分析VSDX文件结构
        result = analyze_vsdx_structure_with_geometry(vsdx_file, debug=True)
        
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

if __name__ == "__main__":
    # 文件路径
    vsdx_file = "/Volumes/Data/vsdx/10_SYS_SYS_XXX_1_1.vsdx"
    json_file = "/Volumes/Data/vsdx/vsdx_structure_with_master_geometry.json"
    
    # 如果需要从VSDX文件重新提取数据
    # export_vsdx_to_json(vsdx_file, json_file)
    
    # 绘制完整结构图
    visualize_vsdx_structure(json_file, "/Volumes/Data/vsdx/vsdx_visualization.png")
    
    # 绘制连接器示意图
    visualize_connector_diagram(json_file, "/Volumes/Data/vsdx/connector_diagram.png")
    
    print("可视化完成！")