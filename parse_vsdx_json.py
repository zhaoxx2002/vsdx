import json
import zipfile
import os
from xml.etree import ElementTree as ET

namespaces = {
    'v': 'http://schemas.microsoft.com/office/visio/2012/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}
ET.register_namespace('', namespaces['v'])

# 缓存 Master 结构
master_shape_cache = {}
# 缓存连接信息
connection_cache = {}

# ========== 读取 Shape 文本映射 ==========
def extract_shape_texts(vsdx_path):
    shape_texts = {}
    with zipfile.ZipFile(vsdx_path, 'r') as zf:
        for page_file in [f for f in zf.namelist() if f.startswith('visio/pages/page') and f.endswith('.xml')]:
            with zf.open(page_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                for shape in root.findall('.//v:Shape', namespaces):
                    sid = shape.get('ID')
                    text_elem = shape.find('.//v:Text', namespaces)
                    text = ''.join(text_elem.itertext()).strip() if text_elem is not None else ''
                    shape_texts[sid] = text if text else shape.get('NameU', f"ID:{sid}")
    return shape_texts

# ========== 读取连接信息 ==========
def extract_connections(vsdx_path):
    connections = {}
    with zipfile.ZipFile(vsdx_path, 'r') as zf:
        for page_file in [f for f in zf.namelist() if f.startswith('visio/pages/page') and f.endswith('.xml')]:
            page_id = page_file.split('page')[-1].split('.')[0]
            connections[page_id] = []
            
            with zf.open(page_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                
                # 查找所有连接线
                for connect in root.findall('.//v:Connect', namespaces):
                    from_sheet = connect.get('FromSheet')
                    to_sheet = connect.get('ToSheet')
                    from_cell = connect.get('FromCell', '')
                    to_cell = connect.get('ToCell', '')
                    
                    if from_sheet and to_sheet:
                        connections[page_id].append({
                            'from_id': from_sheet,
                            'to_id': to_sheet,
                            'from_cell': from_cell,
                            'to_cell': to_cell
                        })
    return connections

# ========== 读取 Master 坐标 ==========
def load_master_geometry(zf, master_id):
    # 现有代码保持不变
    master_file = f"visio/masters/master{master_id}.xml"
    if master_file in zf.namelist():
        if master_id in master_shape_cache:
            return master_shape_cache[master_id]

        with zf.open(master_file) as f:
            tree = ET.parse(f)
            root = tree.getroot()
            shape_elem = root.find('.//v:Shape', namespaces)
            if shape_elem is not None:
                # 从 Cell 元素获取几何信息
                cells = {}
                for cell in shape_elem.findall('.//v:Cell', namespaces):
                    name = cell.get('N')
                    if name in ['PinX', 'PinY', 'Width', 'Height']:
                        cells[name] = cell.get('V', '?')
                
                if 'PinX' in cells and 'PinY' in cells:
                    geometry = {
                        "x": cells.get('PinX', '?'),
                        "y": cells.get('PinY', '?'),
                        "width": cells.get('Width', '?'),
                        "height": cells.get('Height', '?')
                    }
                    master_shape_cache[master_id] = geometry
                    return geometry
    return {"x": "?", "y": "?", "width": "?", "height": "?"}

# ========== 提取 Shape 属性 ==========
def extract_shape_properties(shape_elem):
    properties = {}
    
    # 提取所有单元格属性
    for cell in shape_elem.findall('.//v:Cell', namespaces):
        name = cell.get('N')
        value = cell.get('V', '')
        if value:  # 只保存有值的属性
            properties[name] = value
    
    # 提取自定义属性
    for prop in shape_elem.findall('.//v:Prop', namespaces):
        name = prop.get('Name', '')
        label = prop.get('Label', '')
        value = ''
        
        # 查找属性值
        value_cell = prop.find('.//v:Value', namespaces)
        if value_cell is not None:
            value = value_cell.text or ''
        
        if name and value:
            properties[name] = value
        if label and value and label != name:
            properties[label] = value
    
    return properties

# ========== 提取 Shape 坐标（含 Master fallback） ==========
# 修改extract_shape_geometry函数以提取更多几何信息
def extract_shape_geometry(shape_elem, zf):
    sid = shape_elem.get('ID')
    
    # 从Cell元素获取几何信息
    cells = {}
    for cell in shape_elem.findall('.//v:Cell', namespaces):
        name = cell.get('N')
        # 扩展提取的几何属性
        if name in ['PinX', 'PinY', 'Width', 'Height', 'Angle', 'FlipX', 'FlipY', 'ResizeMode']:
            cells[name] = cell.get('V', '?')
    
    # 提取几何路径数据
    geometry_sections = []
    for section in shape_elem.findall('.//v:Section', namespaces):
        if section.get('N', '').startswith('Geometry'):
            geo_rows = []
            for row in section.findall('.//v:Row', namespaces):
                row_type = row.get('T', '')
                row_data = {'type': row_type}
                
                for cell in row.findall('.//v:Cell', namespaces):
                    cell_name = cell.get('N', '')
                    cell_value = cell.get('V', '?')
                    row_data[cell_name] = cell_value
                
                geo_rows.append(row_data)
            
            if geo_rows:
                geometry_sections.append(geo_rows)
    
    # 构建完整的几何信息
    geometry = {
        "x": cells.get('PinX', '?'),
        "y": cells.get('PinY', '?'),
        "width": cells.get('Width', '?'),
        "height": cells.get('Height', '?'),
        "angle": cells.get('Angle', '0'),
        "flip_x": cells.get('FlipX', '0'),
        "flip_y": cells.get('FlipY', '0')
    }
    
    # 如果有几何路径数据，添加到结果中
    if geometry_sections:
        geometry["path_data"] = geometry_sections
    
    # 如果没有找到基本几何信息，尝试从master获取
    if geometry["x"] == '?' or geometry["y"] == '?':
        master_id = shape_elem.get("Master")
        if master_id:
            master_geometry = load_master_geometry(zf, master_id)
            # 合并master几何信息，但保留已有的有效值
            for key, value in master_geometry.items():
                if geometry.get(key) == '?' or key not in geometry:
                    geometry[key] = value
    
    return geometry

# 修改parse_shape_recursive函数以提取更多形状信息
def parse_shape_recursive(shape_elem, shape_texts, zf, debug=False):
    sid = shape_elem.get('ID')
    name = shape_texts.get(sid, f"ID:{sid}")
    shape_type = shape_elem.get('Type', 'Shape')
    geometry = extract_shape_geometry(shape_elem, zf)
    properties = extract_shape_properties(shape_elem)
    
    # 提取连接点信息
    connection_points = []
    for cp in shape_elem.findall('.//v:ConnectionPoint', namespaces):
        cp_info = {
            "id": cp.get('ID', '?'),
            "x": cp.get('X', '?'),
            "y": cp.get('Y', '?'),
            "dir": cp.get('Dir', '?'),
            "type": cp.get('Type', '?')
        }
        connection_points.append(cp_info)
    
    # 提取文本信息
    text_content = ""
    text_elem = shape_elem.find('.//v:Text', namespaces)
    if text_elem is not None:
        text_content = ''.join(text_elem.itertext()).strip()
    
    # 提取Master信息
    master_id = shape_elem.get("Master", "")
    master_name = ""
    if master_id:
        # 尝试获取Master名称
        master_file = f"visio/masters/master{master_id}.xml"
        if master_file in zf.namelist():
            try:
                with zf.open(master_file) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    master_elem = root.find('.//v:Master', namespaces)
                    if master_elem is not None:
                        master_name = master_elem.get('Name', '')
            except Exception as e:
                if debug:
                    print(f"[DEBUG] 无法读取Master名称: {e}")

    if debug:
        print(f"[DEBUG] Shape: {name} (ID:{sid}) => XYWH: {geometry}")

    shape_dict = {
        "id": sid,
        "name": name,
        "type": shape_type,
        "position": geometry,
        "properties": properties,
        "master_id": master_id,
        "master_name": master_name,
        "text": text_content,
        "connection_points": connection_points,
        "children": []
    }

    # 递归子shape
    shapes_container = shape_elem.find('v:Shapes', namespaces)
    if shapes_container is not None:
        for sub_shape in shapes_container.findall('v:Shape', namespaces):
            shape_dict["children"].append(parse_shape_recursive(sub_shape, shape_texts, zf, debug))

    return shape_dict

# 修改extract_connectors函数以提取更详细的连接线信息
def extract_connectors(vsdx_path):
    connectors = {}
    with zipfile.ZipFile(vsdx_path, 'r') as zf:
        for page_file in [f for f in zf.namelist() if f.startswith('visio/pages/page') and f.endswith('.xml')]:
            page_id = page_file.split('page')[-1].split('.')[0]
            connectors[page_id] = []
            
            with zf.open(page_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                
                # 查找所有形状
                for shape in root.findall('.//v:Shape', namespaces):
                    # 检查是否为连接线
                    is_connector = False
                    
                    # 方法1: 检查Type属性
                    shape_type = shape.get('Type', '')
                    if shape_type == 'Line' or shape_type == 'Dynamic connector':
                        is_connector = True
                    
                    # 方法2: 检查特定单元格
                    for cell in shape.findall('.//v:Cell', namespaces):
                        if cell.get('N') == 'LinePattern' and cell.get('V', '0') != '0':
                            is_connector = True
                        if cell.get('N') == 'BeginArrow' or cell.get('N') == 'EndArrow':
                            if cell.get('V', '0') != '0':
                                is_connector = True
                    
                    if is_connector:
                        # 提取连接线属性
                        connector_id = shape.get('ID')
                        connector_name = shape.get('NameU', f"ID:{connector_id}")
                        
                        # 提取箭头信息
                        begin_arrow = "0"
                        end_arrow = "0"
                        line_pattern = "0"
                        line_color = "0"
                        line_weight = "0"
                        geometry_points = []
                        
                        # 提取更多线条属性
                        for cell in shape.findall('.//v:Cell', namespaces):
                            cell_name = cell.get('N')
                            if cell_name == 'BeginArrow':
                                begin_arrow = cell.get('V', '0')
                            elif cell_name == 'EndArrow':
                                end_arrow = cell.get('V', '0')
                            elif cell_name == 'LinePattern':
                                line_pattern = cell.get('V', '0')
                            elif cell_name == 'LineColor':
                                line_color = cell.get('V', '0')
                            elif cell_name == 'LineWeight':
                                line_weight = cell.get('V', '0')
                        
                        # 提取几何点
                        for row in shape.findall('.//v:Row', namespaces):
                            row_type = row.get('T')
                            if row_type in ['LineTo', 'MoveTo', 'ArcTo', 'EllipticalArcTo']:
                                point_data = {"type": row_type}
                                
                                for cell in row.findall('.//v:Cell', namespaces):
                                    cell_name = cell.get('N')
                                    cell_value = cell.get('V', '0')
                                    point_data[cell_name] = cell_value
                                
                                # 确保至少有X和Y坐标
                                if 'X' in point_data and 'Y' in point_data:
                                    geometry_points.append(point_data)
                        
                        # 提取连接信息
                        connects = []
                        for connect in root.findall(f".//v:Connect[@FromSheet='{connector_id}']|.//v:Connect[@ToSheet='{connector_id}']", namespaces):
                            from_sheet = connect.get('FromSheet')
                            to_sheet = connect.get('ToSheet')
                            from_cell = connect.get('FromCell', '')
                            to_cell = connect.get('ToCell', '')
                            
                            connects.append({
                                'from_id': from_sheet,
                                'to_id': to_sheet,
                                'from_cell': from_cell,
                                'to_cell': to_cell
                            })
                        
                        connectors[page_id].append({
                            'id': connector_id,
                            'name': connector_name,
                            'type': shape_type,
                            'begin_arrow': begin_arrow,
                            'end_arrow': end_arrow,
                            'line_pattern': line_pattern,
                            'line_color': line_color,
                            'line_weight': line_weight,
                            'geometry_points': geometry_points,
                            'connects': connects
                        })
    
    return connectors

# ========== 主函数 ==========
def analyze_vsdx_structure_with_geometry(vsdx_path, debug=False):
    shape_texts = extract_shape_texts(vsdx_path)
    pages_data = []

    with zipfile.ZipFile(vsdx_path, 'r') as zf:
                connects_root = root.find('.//v:Connects', namespaces)  # 修改：使用.//确保能找到嵌套的Connects
                if connects_root is not None:
                    for connect in connects_root.findall('v:Connect', namespaces):
                        page_connects.append({
                            "from_shape_id": connect.get('FromSheet'),
                            "to_shape_id": connect.get('ToSheet'),
                            "from_cell": connect.get('FromCell'),
                # ------------------------- 新增：解析页面级连接信息 -------------------------
                page_connects = []
                connects_root = root.find('.//v:Connects', namespaces)  # 修改：使用.//确保能找到嵌套的Connects

                page_shapes = []
                        page_connects.append({
                            "from_shape_id": connect.get('FromSheet'),
                            "to_shape_id": connect.get('ToSheet'),
                            shape_info = parse_shape_recursive(shape, shape_texts, zf, debug)
                            page_shapes.append(shape_info)

                        if debug and page_connects:
                            print(f"[DEBUG] 页面 {page_file} 找到 {len(page_connects)} 个连接")
                shape_id_map = {}
                
                def add_shapes_to_map(shapes_list):
                    for shape in shapes_list:
                        shape_id_map[shape["id"]] = shape
                        if shape["children"]:
                            add_shapes_to_map(shape["children"])
                
                # ------------------------- 将连接信息关联到对应形状 -------------------------
                
                # 处理页面级连接
                for conn in page_connects:
                    from_shape_id = conn["from_shape_id"]
                    to_shape_id = conn["to_shape_id"]
                    
                    if from_shape_id in shape_id_map:
                        shape_id_map[from_shape_id]["connections"].append({
                            "from_shape": from_shape_id,
                            "to_shape": to_shape_id,
                            "from_cell": conn["from_cell"],
                            "to_cell": conn["to_cell"]
                for conn in page_connects:
                    from_shape_id = conn["from_shape_id"]
                    to_shape_id = conn["to_shape_id"]

                    if from_shape_id in shape_id_map:
                        shape_id_map[from_shape_id]["connections"].append({
                            "from_shape": from_shape_id,
                            "to_shape": to_shape_id,
                            "from_cell": conn["from_cell"],
                            "to_cell": conn["to_cell"]
                        })
                        "page_index": idx,
                            print(f"[DEBUG] 添加连接: {from_shape_id} -> {to_shape_id}")
                        "shapes": page_shapes
                    })

    return pages_data

# ========== 判断是否为连接线 ==========
def is_connector(shape_elem):
    # 检查是否为连接线类型
    shape_type = shape_elem.get('Type', '')
    if shape_type in ['Line', 'Dynamic connector']:
        return True
    
    # 检查是否有线条相关属性
    for cell in shape_elem.findall('.//v:Cell', namespaces):
        if cell.get('N') in ['LinePattern', 'BeginArrow', 'EndArrow'] and cell.get('V', '0') != '0':
            return True
    
    # 检查是否有连接点
    if shape_elem.findall('.//v:ConnectionPoint', namespaces):
        return True
    
    return False

# ========== 判断是否为核心组件 ==========
def is_core_component(shape_elem, shape_texts, debug=False):
    sid = shape_elem.get('ID')
    name = shape_texts.get(sid, '')
    shape_type = shape_elem.get('Type', '')
    
    # 1. 排除特定类型
    if shape_type in ['Guide', 'ThemeShape', 'Documentation', 'Annotation']:
        return False
    
    # 2. 排除特定名称模式
    exclude_patterns = [
        'ID:', '©', 'VOLKSWAGEN', 'Frame', 'Border', 'Sheet', 'Title',
        'Page', 'Background', '0-45', 'Tel.', 'Format', 'Blatt', 'Scale'
    ]
    if any(pattern in name for pattern in exclude_patterns):
        return False
    
    # 3. 检查组件特征
    has_connection_points = bool(shape_elem.findall('.//v:ConnectionPoint', namespaces))
    has_children = bool(shape_elem.find('v:Shapes', namespaces))
    has_text = bool(shape_elem.find('.//v:Text', namespaces))
    
    # 4. 检查几何尺寸
    cells = {}
    for cell in shape_elem.findall('.//v:Cell', namespaces):
        name = cell.get('N')
        if name in ['PinX', 'PinY', 'Width', 'Height']:
            try:
                value = float(cell.get('V', '0'))
                cells[name] = value
            except ValueError:
                continue
    
    # 放宽几何条件
    has_valid_geometry = (
        cells.get('Width', 0) > 0.05 and 
        cells.get('Height', 0) > 0.05 and
        cells.get('PinX', 0) > 0.1 and 
        cells.get('PinY', 0) > 0.1
    )
    
    if debug:
        print(f"[DEBUG] Shape {name} (ID:{sid}):")
        print(f"  - Type: {shape_type}")
        print(f"  - Has connection points: {has_connection_points}")
        print(f"  - Has children: {has_children}")
        print(f"  - Has text: {has_text}")
        print(f"  - Has valid geometry: {has_valid_geometry}")
        if has_valid_geometry:
            print(f"  - Geometry: W={cells.get('Width', '?')}, H={cells.get('Height', '?')}")
    
    # 5. 组合判断：放宽条件
    # Group 类型需要有子元素
    if shape_type == 'Group':
        return has_valid_geometry and has_children
    # 普通形状需要有连接点或文本
    return has_valid_geometry and (has_connection_points or has_text)

    if debug:  # 现在可以安全使用 debug 参数
        print(f"检查形状: ID={sid}, Name={name}")
        print(f"Type={shape_elem.get('Type', 'N/A')}")
        print(f"Has ConnectionPoints: {bool(shape_elem.findall('.//v:ConnectionPoint', namespaces))}")
        print(f"Has Children: {bool(shape_elem.find('v:Shapes', namespaces))}")
        print(f"Has Text: {bool(shape_elem.find('.//v:Text', namespaces))}")
    
    # 检查是否为边框或表格（通常这些元素的 Type 属性会有特定值）
    shape_type = shape_elem.get('Type', '')
    if shape_type in ['ThemeShape', 'Documentation', 'Annotation']:
        return False
    
    # 检查名称特征（排除明显的非核心组件）
    if any(keyword in name for keyword in ['Sheet', 'Title', 'Page', 'Border', 'Background']):
        return False
    
    # 检查是否有连接点（电气元件通常有连接点）
    has_connection_points = bool(shape_elem.findall('.//v:ConnectionPoint', namespaces))
    
    # 检查是否有子形状
    has_children = bool(shape_elem.find('v:Shapes', namespaces))
    
    # 检查是否有文本内容（电气元件通常有标识文本）
    has_text = bool(shape_elem.find('.//v:Text', namespaces))
    
    # 检查几何信息
    cells = {}
    for cell in shape_elem.findall('.//v:Cell', namespaces):
        name = cell.get('N')
        if name in ['PinX', 'PinY']:
            try:
                cells[name] = float(cell.get('V', '0'))
            except ValueError:
                cells[name] = 0
    
    # 放宽位置限制
    is_in_valid_area = True
    if 'PinX' in cells and 'PinY' in cells:
        # 排除明显在边缘的元素
        is_in_valid_area = cells['PinX'] > 0.5 and cells['PinY'] > 0.5
    
    # 组合判断条件（放宽条件）
    return is_in_valid_area and (has_connection_points or has_children or has_text)
    
    # 增加更严格的过滤条件
    if shape_type in ['Guide'] or name.startswith('ID:'):
        return False
        
    # 过滤掉版权信息
    if any(keyword in name for keyword in ['©', 'copyright', 'VOLKSWAGEN AG']):
        return False

# ========== 执行 ==========
if __name__ == "__main__":
    try:
        vsdx_file = "10_SYS_SYS_XXX_1_1.vsdx"
        output_json = "vsdx_structure_with_master_geometry.json"
        debug_mode = True  # 开启调试
    
        # if not os.path.exists(vsdx_file):
        #     raise FileNotFoundError(f"找不到 VSDX 文件：{vsdx_file}")
    
        result = analyze_vsdx_structure_with_geometry(vsdx_file, debug=debug_mode)
        
        if not result:
            print("警告：未找到任何核心组件")
            
        with open(output_json, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        print(f"\n✅ 完成：含 Master fallback 的结构写入 {output_json}")
        
    except Exception as e:
        print(f"错误：{str(e)}")
