import json
import zipfile
from xml.etree import ElementTree as ET

namespaces = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}
ET.register_namespace('', namespaces['v'])

# 缓存 Master 结构
master_shape_cache = {}

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


# ========== 读取 Master 坐标 ==========
def load_master_geometry(zf, master_id):
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


# ========== 提取 Shape 坐标（含 Master fallback） ==========
def get_shape_geometry(shape_elem, zf):
    sid = shape_elem.get('ID')

    # 从 Cell 元素获取几何信息
    cells = {}
    for cell in shape_elem.findall('.//v:Cell', namespaces):
        name = cell.get('N')
        if name in ['PinX', 'PinY', 'Width', 'Height']:
            cells[name] = cell.get('V', '?')

    if 'PinX' in cells and 'PinY' in cells:
        return {
            "x": cells.get('PinX', '?'),
            "y": cells.get('PinY', '?'),
            "width": cells.get('Width', '?'),
            "height": cells.get('Height', '?')
        }

    # fallback 到 master 形状坐标
    master_id = shape_elem.get("Master")
    if master_id:
        print(f"[DEBUG] 形状 ID:{sid} 尝试从 Master:{master_id} 获取几何信息")
        return load_master_geometry(zf, master_id)

    print(f"[DEBUG] 警告: 形状 ID:{sid} 没有找到几何信息")
    return {"x": "?", "y": "?", "width": "?", "height": "?"}


# ========== 提取连接关系 ==========
def extract_connections(shape_elem):
    """提取形状的连接关系"""
    connections = []
    for connect in shape_elem.findall('.//v:Connect', namespaces):
        conn_data = {
            "from_shape": connect.get('FromSheet'),
            "to_shape": connect.get('ToSheet'),
            "from_cell": connect.get('FromCell'),
            "to_cell": connect.get('ToCell')
        }
        connections.append(conn_data)
    return connections


# ========== 递归解析 Shape ==========
def parse_shape_recursive(shape_elem, shape_texts, zf, debug=False):
    sid = shape_elem.get('ID')
    name = shape_texts.get(sid, f"ID:{sid}")
    shape_type = shape_elem.get('Type', 'N/A')
    geometry = get_shape_geometry(shape_elem, zf)

    if debug:
        print(f"[DEBUG] Shape: {name} (ID:{sid}) => XYWH: {geometry}")

    shape_dict = {
        "id": sid,
        "name": name,
        "type": shape_type,
        "position": geometry,
        "children": [],
        "connections": []  # 新增字段存储连接信息
    }

    # 提取连接关系
    connections = extract_connections(shape_elem)
    if connections:
        shape_dict["connections"] = connections
        if debug:
            print(f"[DEBUG] Shape {name} (ID:{sid}) has connections: {connections}")

    # 递归子 shape
    shapes_container = shape_elem.find('v:Shapes', namespaces)
    if shapes_container is not None:
        for sub_shape in shapes_container.findall('v:Shape', namespaces):
            shape_dict["children"].append(parse_shape_recursive(sub_shape, shape_texts, zf, debug))

    return shape_dict


# ========== 主函数 ==========
def analyze_vsdx_structure_with_geometry(vsdx_path, debug=False):
    shape_texts = extract_shape_texts(vsdx_path)
    pages_data = []

    with zipfile.ZipFile(vsdx_path, 'r') as zf:
        # 检查文件中是否有连接信息
        if debug:
            for file_name in zf.namelist():
                if file_name.startswith('visio/pages/') and file_name.endswith('.xml'):
                    with zf.open(file_name) as f:
                        content = f.read().decode('utf-8')
                        if '<Connect ' in content:
                            print(f"[DEBUG] 文件 {file_name} 中找到连接信息")
        
        page_files = [f for f in zf.namelist() if f.startswith('visio/pages/page') and f.endswith('.xml')]
        for idx, page_file in enumerate(page_files, start=1):
            with zf.open(page_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                
                # 获取页面ID
                page_id = root.get('ID', str(idx))
                
                # ------------------------- 修改：先收集所有连接信息 -------------------------
                # 使用 .// 确保能找到嵌套的 Connects 元素
                connects = []
                connected_shape_ids = set()  # 存储所有连接中涉及的形状ID
                
                connects_root = root.find('.//v:Connects', namespaces)
                if connects_root is not None:
                    for connect in connects_root.findall('v:Connect', namespaces):
                        from_shape = connect.get('FromSheet')
                        to_shape = connect.get('ToSheet')
                        
                        connects.append({
                            "from_shape": from_shape,
                            "to_shape": to_shape,
                            "from_cell": connect.get('FromCell'),
                            "to_cell": connect.get('ToCell')
                        })
                        
                        # 收集连接中涉及的形状ID
                        if from_shape:
                            connected_shape_ids.add(from_shape)
                        if to_shape:
                            connected_shape_ids.add(to_shape)
                    
                    if debug and connects:
                        print(f"[DEBUG] 页面 {page_file} 找到 {len(connects)} 个连接")
                        print(f"[DEBUG] 连接中涉及的形状ID: {connected_shape_ids}")
                        # 打印前几个连接信息以便调试
                        for i, conn in enumerate(connects[:3]):
                            print(f"[DEBUG] 连接 {i+1}: {conn['from_shape']} -> {conn['to_shape']}")
                
                # ------------------------- 修改：处理所有形状，包括连接中涉及的形状 -------------------------
                shapes_root = root.find('v:Shapes', namespaces)
                page_shapes = []
                
                # 创建一个自定义的is_core_component函数，考虑连接中的形状ID
                def is_core_with_connections(shape_elem):
                    sid = shape_elem.get('ID')
                    # 如果形状ID在连接中出现，则保留
                    if sid in connected_shape_ids:
                        if debug:
                            print(f"[DEBUG] 保留形状 ID:{sid} (在连接中出现)")
                        return True
                    # 否则使用原始判断逻辑
                    return is_core_component(shape_elem, shape_texts, debug)
                
                if shapes_root is not None:
                    for shape in shapes_root.findall('v:Shape', namespaces):
                        if is_core_with_connections(shape):
                            shape_info = parse_shape_recursive(shape, shape_texts, zf, debug)
                            page_shapes.append(shape_info)
                
                # 建立形状ID到形状字典的映射（包括所有子形状）
                shape_id_map = {}
                
                def add_shapes_to_map(shapes_list):
                    for shape in shapes_list:
                        shape_id_map[shape["id"]] = shape
                        if shape["children"]:
                            add_shapes_to_map(shape["children"])
                
                add_shapes_to_map(page_shapes)
                
                # 处理页面级连接
                connections_added = 0
                for conn in connects:
                    from_shape = conn["from_shape"]
                    to_shape = conn["to_shape"]
                    
                    if from_shape in shape_id_map:
                        shape_id_map[from_shape]["connections"].append(conn)
                        connections_added += 1
                        if debug:
                            print(f"[DEBUG] 添加连接: {from_shape} -> {to_shape}")
                
                if debug and connects:
                    print(f"[DEBUG] 总共添加了 {connections_added}/{len(connects)} 个连接")
                    if connections_added == 0:
                        print("[DEBUG] 警告: 没有连接被添加到任何形状")
                        print("[DEBUG] 形状ID列表:", list(shape_id_map.keys())[:10], "...")
                        print("[DEBUG] 连接中的形状ID:", list(connected_shape_ids)[:10], "...")

                if page_shapes:
                    pages_data.append({
                        "page_index": idx,
                        "page_id": page_id,
                        "page_file": page_file,
                        "shapes": page_shapes
                    })

    return pages_data


# ========== 判断是否为核心组件 ==========
def is_core_component(shape_elem, shape_texts, debug=False):
    sid = shape_elem.get('ID')
    name = shape_texts.get(sid, '')
    shape_type = shape_elem.get('Type', '')

    # 1. 排除特定类型
    if shape_type in ['Guide', 'ThemeShape', 'Documentation', 'Annotation']:
        return False

    # 2. 特殊处理：如果形状ID在连接中出现，则保留
    # 这部分需要在主函数中实现，因为这里无法访问连接信息
    
    # 3. 特殊处理：Dynamic Connector 类型总是保留
    if shape_type == 'Dynamic Connector':
        return True
        
    # 4. 排除特定名称模式
    exclude_patterns = [
        'ID:', '©', 'VOLKSWAGEN', 'Frame', 'Border', 'Sheet', 'Title',
        'Page', 'Background', '0-45', 'Tel.', 'Format', 'Blatt', 'Scale'
    ]
    if any(pattern in name for pattern in exclude_patterns):
        return False

    # 5. 检查组件特征
    has_connection_points = bool(shape_elem.findall('.//v:ConnectionPoint', namespaces))
    has_children = bool(shape_elem.find('v:Shapes', namespaces))
    has_text = bool(shape_elem.find('.//v:Text', namespaces))

    # 6. 检查几何尺寸
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

    # 7. 组合判断：放宽条件
    # Group 类型需要有子元素
    if shape_type == 'Group':
        return has_children  # 移除几何条件限制
    # 普通形状需要有连接点或文本
    return has_connection_points or has_text  # 移除几何条件限制


# ========== 执行 ==========
if __name__ == "__main__":
    try:
        vsdx_file = "10_SYS_SYS_XXX_1_1.vsdx"
        output_json = "vsdx_structure_with_master_geometry140.json"
        debug_mode = True  # 开启调试

        result = analyze_vsdx_structure_with_geometry(vsdx_file, debug=debug_mode)

        if not result:
            print("警告：未找到任何核心组件")

        with open(output_json, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        print(f"\n✅ 完成：含 Master fallback 的结构写入 {output_json}")

    except Exception as e:
        print(f"错误：{str(e)}")
        