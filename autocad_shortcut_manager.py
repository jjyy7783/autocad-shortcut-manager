#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AutoCAD 快捷键管理工具 v2.3

开源项目 | 作者: jjyy7783
支持 AutoCAD 2004-2026 全版本
支持用户模式：保存/加载/切换个人快捷键配置
新增：解决插件快捷键冲突问题

License: MIT / Open Source
GitHub: (待补充)
"""

import os
import sys
import re
import shutil
import json
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass, asdict
from datetime import datetime

# 修复 Windows 控制台 Unicode 输出问题
if sys.platform == 'win32' and not getattr(sys, 'frozen', False):
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


@dataclass
class Shortcut:
    """快捷键定义"""
    alias: str          # 快捷命令（如：L）
    command: str        # 完整命令（如：LINE）
    description: str    # 描述/注释
    line_num: int       # 在文件中的行号
    is_builtin: bool    # 是否为内置快捷键


@dataclass
class UserProfile:
    """用户配置"""
    name: str
    description: str
    created_at: str
    modified_at: str
    shortcuts: List[Dict]  # 自定义快捷键列表


# AutoCAD 常用命令数据库（命令名 -> 中文描述）
AUTOCAD_COMMANDS = {
    # 绘图命令
    "LINE": "直线", "L": "直线(别名)", "LWLINE": "轻量多段线", "PLINE": "多段线", "PL": "多段线(别名)",
    "CIRCLE": "圆", "C": "圆(别名)", "ARC": "圆弧", "A": "圆弧(别名)", 
    "RECTANG": "矩形", "REC": "矩形(别名)", "POLYGON": "正多边形", "POL": "正多边形(别名)",
    "ELLIPSE": "椭圆", "EL": "椭圆(别名)", "HATCH": "图案填充", "H": "图案填充(别名)",
    "POINT": "点", "PO": "点(别名)", "DIVIDE": "定数等分", "DIV": "定数等分(别名)",
    "MEASURE": "定距等分", "ME": "定距等分(别名)", "BOUNDARY": "边界创建", "BO": "边界创建(别名)",
    "REGION": "面域", "REG": "面域(别名)", "SOLID": "二维填充", "SO": "二维填充(别名)",
    "REVCLOUD": "修订云线", "WIPEOUT": "遮罩对象", "SPLINE": "样条曲线", "SPL": "样条曲线(别名)",
    "XLINE": "构造线", "XL": "构造线(别名)", "RAY": "射线", "MLINE": "多线", "ML": "多线(别名)",
    "SKETCH": "徒手画",
    
    # 文字命令
    "TEXT": "单行文字", "DT": "单行文字(别名)", "MTEXT": "多行文字", "T": "多行文字(别名)",
    "TABLE": "表格", "TB": "表格(别名)", "TABLESTYLE": "表格样式",
    
    # 标注命令
    "DIMLINEAR": "线性标注", "DLI": "线性标注(别名)", "DIMALIGNED": "对齐标注", "DAL": "对齐标注(别名)",
    "DIMRADIUS": "半径标注", "DRA": "半径标注(别名)", "DIMDIAMETER": "直径标注", "DDI": "直径标注(别名)",
    "DIMANGULAR": "角度标注", "DAN": "角度标注(别名)", "DIMCENTER": "中心标记", "DCE": "中心标记(别名)",
    "DIMORDINATE": "坐标标注", "DOR": "坐标标注(别名)", "DIMCONTINUE": "连续标注", "DCO": "连续标注(别名)",
    "DIMBASELINE": "基线标注", "DBA": "基线标注(别名)", "DIMBREAK": "标注打断", "DIMSPACE": "标注间距",
    "DIMJOGLINE": "折弯标注", "QDIM": "快速标注", "MLEADER": "多重引线", "MLD": "多重引线(别名)",
    "QLEADER": "快速引线", "LE": "快速引线(别名)", "TOLERANCE": "形位公差", "TOL": "形位公差(别名)",
    "DIMEDIT": "标注编辑", "DED": "标注编辑(别名)", "DIMTEDIT": "标注文字编辑", "DIMSTYLE": "标注样式", "D": "标注样式(别名)",
    
    # 编辑命令
    "MOVE": "移动", "M": "移动(别名)", "COPY": "复制", "CO": "复制(别名)", "CP": "复制(别名)",
    "ROTATE": "旋转", "RO": "旋转(别名)", "SCALE": "缩放", "SC": "缩放(别名)",
    "MIRROR": "镜像", "MI": "镜像(别名)", "OFFSET": "偏移", "O": "偏移(别名)",
    "ARRAY": "阵列", "AR": "阵列(别名)", "TRIM": "修剪", "TR": "修剪(别名)",
    "EXTEND": "延伸", "EX": "延伸(别名)", "FILLET": "圆角", "F": "圆角(别名)",
    "CHAMFER": "倒角", "CHA": "倒角(别名)", "BREAK": "打断", "BR": "打断(别名)",
    "JOIN": "合并", "J": "合并(别名)", "EXPLODE": "分解", "X": "分解(别名)",
    "ERASE": "删除", "E": "删除(别名)", "OOPS": "恢复删除", "STRETCH": "拉伸", "S": "拉伸(别名)",
    "LENGTHEN": "拉长", "LEN": "拉长(别名)", "ALIGN": "对齐", "AL": "对齐(别名)",
    "PEDIT": "编辑多段线", "PE": "编辑多段线(别名)", "SPLINEDIT": "编辑样条曲线", "SPE": "编辑样条曲线(别名)",
    "HATCHEDIT": "编辑图案填充", "HE": "编辑图案填充(别名)", "TEXTEDIT": "编辑文字", "ED": "编辑文字(别名)",
    "MIRROR3D": "三维镜像", "3DARRAY": "三维阵列", "3DMIRROR": "三维镜像",
    
    # 视图命令
    "ZOOM": "缩放视图", "Z": "缩放视图(别名)", "PAN": "平移视图", "P": "平移视图(别名)",
    "3DORBIT": "三维动态观察", "3DO": "三维动态观察(别名)", "VPOINT": "视点设置", "VP": "视点设置(别名)",
    "VIEW": "视图管理", "V": "视图管理(别名)", "SHADEMODE": "着色模式", "REGEN": "重生成", "RE": "重生成(别名)",
    "REGENALL": "全部重生成", "REA": "全部重生成(别名)", "REDRAW": "重画", "REDRAWALL": "全部重画",
    
    # 图层命令
    "LAYER": "图层管理", "LA": "图层管理(别名)", "LAYOFF": "关闭图层", "LAYON": "打开所有图层",
    "LAYFRZ": "冻结图层", "LAYTHW": "解冻所有图层", "LAYLCK": "锁定图层", "LAYULK": "解锁图层",
    "VPLAYER": "视口图层控制", "LAYISO": "隔离图层", "LAYUNISO": "取消隔离",
    
    # 特性命令
    "PROPERTIES": "特性面板", "PR": "特性面板(别名)", "CH": "特性(别名)", "MO": "特性(别名)",
    "MATCHPROP": "特性匹配", "MA": "特性匹配(别名)", "PAINTER": "特性匹配",
    "PROPERTIESCLOSE": "关闭特性面板", "PRC": "关闭特性面板(别名)",
    
    # 块与属性
    "BLOCK": "创建块", "B": "创建块(别名)", "INSERT": "插入块", "I": "插入块(别名)",
    "WBLOCK": "写块", "W": "写块(别名)", "ATTDEF": "属性定义", "ATT": "属性定义(别名)",
    "ATTEDIT": "编辑属性", "ATE": "编辑属性(别名)", "ATTDISP": "属性显示", "ATTEXT": "属性提取",
    "BEDIT": "块编辑器", "BE": "块编辑器(别名)", "BCLOSE": "关闭块编辑器", "BC": "关闭块编辑器(别名)",
    
    # 查询命令
    "AREA": "面积查询", "AA": "面积查询(别名)", "DIST": "距离查询", "DI": "距离查询(别名)",
    "ID": "点坐标查询", "LIST": "列表显示", "LI": "列表显示(别名)", "LS": "列表显示(别名)",
    "DBLIST": "数据库列表", "STATUS": "状态查询", "TIME": "时间查询", "SETVAR": "设置变量", "SET": "设置变量(别名)",
    
    # 对象捕捉与设置
    "OSNAP": "对象捕捉设置", "OS": "对象捕捉设置(别名)", "DDOSNAP": "对象捕捉设置",
    "GRID": "栅格显示", "SNAP": "捕捉设置", "ORTHO": "正交模式", 
    "POLAR": "极轴追踪", "OTRACK": "对象捕捉追踪", "DSETTINGS": "草图设置", "DS": "草图设置(别名)", "SE": "草图设置(别名)",
    
    # 文件命令
    "SAVE": "保存", "QSAVE": "快速保存", "SAVEAS": "另存为", "OPEN": "打开文件",
    "NEW": "新建文件", "CLOSE": "关闭文件", "QUIT": "退出程序", "EXIT": "退出程序",
    "RECOVER": "修复图形", "AUDIT": "核查图形", "PURGE": "清理图形", "PU": "清理图形(别名)",
    
    # 打印命令
    "PLOT": "打印", "PREVIEW": "打印预览", "PRE": "打印预览(别名)",
    "PAGESETUP": "页面设置", "PSETUP": "页面设置(别名)",
    
    # 视口命令
    "VPORTS": "视口设置", "VPORT": "视口设置(别名)",
    
    # 外部参照
    "XREF": "外部参照管理器", "XR": "外部参照管理器(别名)", 
    "XATTACH": "附着外部参照", "XA": "附着外部参照(别名)",
    "XCLIP": "外部参照剪裁", "XC": "外部参照剪裁(别名)",
    "IMAGE": "图像管理器", "IM": "图像管理器(别名)", 
    "IMAGEATTACH": "附着图像", "IAT": "附着图像(别名)",
    "PDFATTACH": "附着PDF", "DWFATTACH": "附着DWF",
    
    # 其他常用
    "UNDO": "撤销", "U": "撤销(别名)", "REDO": "重做", "COPYCLIP": "复制到剪贴板",
    "CUTCLIP": "剪切到剪贴板", "PASTECLIP": "粘贴", "PASTEORIG": "粘贴到原坐标",
    "OPTIONS": "选项设置", "OP": "选项设置(别名)", "GR": "选项设置(别名)",
    "TOOLPALETTES": "工具选项板", "TP": "工具选项板(别名)",
    "ADCENTER": "设计中心", "ADC": "设计中心(别名)", "DC": "设计中心(别名)",
    "ADCNAVIGATE": "设计中心导航", "QCCLOSE": "关闭快速计算器",
    "CALCULATOR": "快速计算器", "QC": "快速计算器(别名)", "CAL": "计算器",
    "APPLOAD": "加载应用程序", "AP": "加载应用程序(别名)",
    "SCRIPT": "运行脚本", "SCR": "运行脚本(别名)",
    "DIMALIGNED": "对齐标注", "QSELECT": "快速选择", "QSE": "快速选择(别名)",
    "FILTER": "对象选择过滤器", "FI": "对象选择过滤器(别名)",
    "GROUP": "对象编组", "G": "对象编组(别名)", 
    "COPYTOLAYER": "复制到图层", "COPYM": "多重复制", "LAYMRG": "图层合并",
    "LAYOUT": "布局管理", "LAYOUTWIZARD": "布局向导",
    "MVSETUP": "模型空间设置", "MVIEW": "创建视口", "MV": "创建视口(别名)",
    "VPMAX": "视口最大化", "VPMIN": "视口最小化",
    "CHSPACE": "更改空间", "FLATTEN": "展平",
    
    # 三维建模
    "BOX": "长方体", "SPHERE": "球体", "CYLINDER": "圆柱体", "CONE": "圆锥体",
    "WEDGE": "楔体", "TORUS": "圆环体", "PYRAMID": "棱锥体", "POLYSOLID": "多段体",
    "EXTRUDE": "拉伸", "EXT": "拉伸(别名)", "REVOLVE": "旋转", "REV": "旋转(别名)",
    "SWEEP": "扫掠", "LOFT": "放样", "UNION": "并集", "UNI": "并集(别名)",
    "SUBTRACT": "差集", "SU": "差集(别名)", "INTERSECT": "交集", "IN": "交集(别名)",
    "INTERFERE": "干涉检查", "SLICE": "剖切", "SECTION": "截面", "SEC": "截面(别名)",
    "PRESSPULL": "按住并拖动", "SOLIDEDIT": "实体编辑", "THICKEN": "加厚",
    "SURFSCULPT": "曲面造型", "CONVTOSURFACE": "转换为曲面", "CONVTOSOLID": "转换为实体",
    
    # 渲染
    "RENDER": "渲染", "RR": "渲染(别名)", "RENDERENVIRONMENT": "渲染环境",
    "RENDEREXPOSURE": "渲染曝光", "MATERIALS": "材质浏览器", "MAT": "材质浏览器(别名)",
    "LIGHT": "光源", "LIGHTLIST": "光源列表", "SUNPROPERTIES": "阳光特性",
    
    # 约束与参数化
    "GEOMCONSTRAINT": "几何约束", "DIMCONSTRAINT": "标注约束",
    "PARAMETERS": "参数管理器", "DELCONSTRAINT": "删除约束",
    "CONSTRAINTBAR": "约束栏", "CONSTRAINTSETTINGS": "约束设置",
    
    # 表达式与动态块
    "PARAMETERCOPY": "复制参数", "PARAMETERMANAGER": "参数管理器",
    "BACTION": "块动作", "BAPARAMETER": "块参数", "BACTIONSET": "块动作集",
    "BGRIPSET": "块夹点集", "BPARAMETER": "块参数",
}

def search_commands(keyword: str, limit: int = 20) -> List[Tuple[str, str]]:
    """
    搜索 AutoCAD 命令
    :param keyword: 搜索关键词（支持中文或英文）
    :param limit: 返回数量限制
    :return: [(命令名, 描述), ...]
    """
    keyword = keyword.upper()
    results = []
    
    for cmd, desc in AUTOCAD_COMMANDS.items():
        # 匹配命令名或描述
        if keyword in cmd or keyword in desc.upper():
            # 去掉"(别名)"标记的优先级降低
            is_alias = "(别名)" in desc
            results.append((cmd, desc, is_alias))
    
    # 排序：非别名优先
    results.sort(key=lambda x: (x[2], x[0]))
    
    # 返回前 limit 个
    return [(r[0], r[1]) for r in results[:limit]]


class AutoCADShortcutManager:
    """AutoCAD 快捷键管理器"""
    
    # AutoCAD 版本和对应的 acad.pgp 路径模板
    AUTOCAD_VERSIONS = list(range(2004, 2027))  # 2004-2026
    
    def __init__(self, pgp_path: Optional[str] = None):
        """
        初始化管理器
        :param pgp_path: 指定 acad.pgp 文件路径，None 则自动查找
        """
        self.pgp_path = pgp_path or self._find_acad_pgp()
        self.shortcuts: List[Shortcut] = []
        self.file_content: List[str] = []
        self.header_lines: List[str] = []
        self.builtin_shortcuts: List[Shortcut] = []
        self.custom_shortcuts: List[Shortcut] = []
        
        # 用户配置目录
        self.profiles_dir = Path.home() / ".autocad_shortcuts"
        self.profiles_dir.mkdir(exist_ok=True)
        
        if self.pgp_path and os.path.exists(self.pgp_path):
            self._load_file()
            self._parse_shortcuts()
    
    def _get_autocad_paths(self) -> List[str]:
        """生成所有可能的 AutoCAD 路径"""
        paths = []
        
        for year in self.AUTOCAD_VERSIONS:
            # 标准安装路径
            paths.extend([
                rf"C:\Program Files\Autodesk\AutoCAD {year}\Support\acad.pgp",
                rf"C:\Program Files\Autodesk\AutoCAD {year}\UserDataCache\Support\acad.pgp",
                rf"C:\Program Files (x86)\Autodesk\AutoCAD {year}\Support\acad.pgp",
                rf"C:\Program Files (x86)\Autodesk\AutoCAD {year}\UserDataCache\Support\acad.pgp",
            ])
        
        # 通用路径（不区分版本）
        paths.extend([
            r"C:\Program Files\Autodesk\AutoCAD\Support\acad.pgp",
            r"C:\Program Files (x86)\Autodesk\AutoCAD\Support\acad.pgp",
        ])
        
        # 用户配置路径（可能包含版本信息）
        appdata = os.environ.get('APPDATA', '')
        if appdata:
            paths.extend([
                os.path.join(appdata, 'Autodesk', 'AutoCAD', 'acad.pgp'),
                os.path.join(appdata, 'Autodesk', 'AutoCAD', 'R24.3', 'acad.pgp'),
                os.path.join(appdata, 'Autodesk', 'AutoCAD', 'R24.2', 'acad.pgp'),
                os.path.join(appdata, 'Autodesk', 'AutoCAD', 'R24.1', 'acad.pgp'),
                os.path.join(appdata, 'Autodesk', 'AutoCAD', 'R24.0', 'acad.pgp'),
                os.path.join(appdata, 'Autodesk', 'AutoCAD', 'R23.1', 'acad.pgp'),
                os.path.join(appdata, 'Autodesk', 'AutoCAD', 'R23.0', 'acad.pgp'),
                os.path.join(appdata, 'Autodesk', 'AutoCAD', 'R22.0', 'acad.pgp'),
            ])
            # 扫描所有可能的版本目录（递归搜索）
            autocad_appdata = Path(appdata) / 'Autodesk'
            if autocad_appdata.exists():
                # 使用 rglob 递归查找所有 acad.pgp 文件
                for pgp_file in autocad_appdata.rglob('acad.pgp'):
                    paths.append(str(pgp_file))
                # 也扫描带语言名称的目录（如 "AutoCAD 2012 - Simplified Chinese"）
                for subdir in autocad_appdata.iterdir():
                    if subdir.is_dir() and 'autocad' in subdir.name.lower():
                        paths.append(str(subdir / 'acad.pgp'))
                        paths.append(str(subdir / 'Support' / 'acad.pgp'))
        
        # 本地应用数据路径
        localappdata = os.environ.get('LOCALAPPDATA', '')
        if localappdata:
            paths.extend([
                os.path.join(localappdata, 'Autodesk', 'AutoCAD', 'acad.pgp'),
            ])
            # 扫描本地应用数据中的 AutoCAD 目录
            local_autocad = Path(localappdata) / 'Autodesk'
            if local_autocad.exists():
                for subdir in local_autocad.rglob('acad.pgp'):
                    paths.append(str(subdir))
        
        return paths
    
    def _find_acad_pgp(self) -> Optional[str]:
        """自动查找 acad.pgp 文件"""
        paths = self._get_autocad_paths()
        
        for path in paths:
            if os.path.exists(path):
                return path
        
        return None
    
    def find_all_acad_pgp(self) -> List[Tuple[str, str]]:
        """查找系统中所有 acad.pgp 文件"""
        found = []
        paths = self._get_autocad_paths()
        
        for path in paths:
            if os.path.exists(path):
                # 提取版本信息
                version = self._extract_version_from_path(path)
                found.append((path, version))
        
        # 去重
        seen = set()
        unique = []
        for path, version in found:
            if path not in seen:
                seen.add(path)
                unique.append((path, version))
        
        return unique
    
    def _extract_version_from_path(self, path: str) -> str:
        """从路径提取 AutoCAD 版本"""
        # 尝试匹配年份
        match = re.search(r'AutoCAD\s*(\d{4})', path)
        if match:
            return f"AutoCAD {match.group(1)}"
        
        # 尝试匹配 Rxx.x 版本号
        match = re.search(r'R(\d+\.\d+)', path)
        if match:
            r_version = match.group(1)
            version_map = {
                '24.3': '2025', '24.2': '2024', '24.1': '2023', '24.0': '2022',
                '23.1': '2021', '23.0': '2020', '22.0': '2019',
                '21.0': '2018', '20.1': '2017', '20.0': '2016',
                '19.1': '2015', '19.0': '2014', '18.1': '2013', '18.0': '2012',
                '17.2': '2011', '17.1': '2010', '17.0': '2009',
                '16.2': '2008', '16.1': '2007', '16.0': '2006',
                '15.2': '2005', '15.0': '2004'
            }
            year = version_map.get(r_version, r_version)
            return f"AutoCAD {year} (R{r_version})"
        
        return "未知版本"
    
    def _load_file(self):
        """加载 pgp 文件内容"""
        try:
            with open(self.pgp_path, 'r', encoding='utf-8') as f:
                self.file_content = f.readlines()
        except UnicodeDecodeError:
            try:
                with open(self.pgp_path, 'r', encoding='gbk') as f:
                    self.file_content = f.readlines()
            except UnicodeDecodeError:
                with open(self.pgp_path, 'r', encoding='latin-1') as f:
                    self.file_content = f.readlines()
    
    def _parse_shortcuts(self):
        """解析快捷键定义"""
        self.shortcuts = []
        self.builtin_shortcuts = []
        self.custom_shortcuts = []
        in_alias_section = False
        
        # 内置快捷键列表（常见命令）
        builtin_commands = {
            'L', 'C', 'REC', 'PL', 'POL', 'ARC', 'EL', 'T', 'MT', 'DT',
            'CO', 'CP', 'MI', 'AR', 'O', 'TR', 'EX', 'F', 'CHA', 'RO',
            'SC', 'ST', 'LA', 'CH', 'MA', 'DI', 'E', 'M', 'H', 'B',
            'I', 'ATT', 'W', 'X', 'Z', 'P', 'D', 'G', 'J', 'K',
            'N', 'Q', 'R', 'S', 'U', 'V', 'A', 'ED', 'REGEN', 'RE',
            'UNDO', 'REDO', 'SAVE', 'QUIT', 'OPEN', 'NEW', 'PRINT',
            'LAYER', 'COLOR', 'LINETYPE', 'STYLE', 'DIM', 'BLOCK',
            'INSERT', 'ATTDEF', 'TEXT', 'MTEXT', 'DIMLINEAR', 'DIMALIGNED',
            'DIMRADIUS', 'DIMDIAMETER', 'DIMANGULAR', 'HATCH', 'BHATCH',
            'BOUNDARY', 'REGION', 'UNION', 'SUBTRACT', 'INTERSECT',
            'EXPLODE', 'JOIN', 'PEDIT', 'MLEDIT', 'SPLINEDIT', 'GROUP',
            'FILTER', 'QUICKSELECT', 'PROPERTIES', 'MATCHPROP', 'LIST',
            'DIST', 'AREA', 'ID', 'TIME', 'STATUS', 'SETVAR', 'OPTIONS'
        }
        
        for i, line in enumerate(self.file_content, 1):
            line_stripped = line.strip()
            
            # 检测快捷键定义区域
            if 'Command alias format' in line or '命令别名格式' in line or 'alias' in line.lower():
                in_alias_section = True
                continue
            
            # 收集头部注释
            if not in_alias_section and (line_stripped.startswith('*') or line_stripped.startswith(';')):
                self.header_lines.append(line.rstrip())
                continue
            
            # 跳过空行和纯注释行
            if not line_stripped or line_stripped.startswith(';'):
                continue
            
            # 解析快捷键定义行
            # 支持格式：别名, *命令  或  别名,*命令  或  别名, *命令 ; 注释
            match = re.match(r'^(\w+)\s*,\s*\*([^\s;]+)(?:\s*;\s*(.*))?$', line_stripped)
            if match:
                alias = match.group(1).upper()
                command = match.group(2).upper()
                description = match.group(3).strip() if match.group(3) else ""
                
                is_builtin = alias in builtin_commands or command in builtin_commands
                
                shortcut = Shortcut(
                    alias=alias,
                    command=command,
                    description=description,
                    line_num=i,
                    is_builtin=is_builtin
                )
                self.shortcuts.append(shortcut)
                
                if is_builtin:
                    self.builtin_shortcuts.append(shortcut)
                else:
                    self.custom_shortcuts.append(shortcut)
    
    def get_shortcuts(self, search_term: str = "", only_custom: bool = False, 
                     only_builtin: bool = False) -> List[Shortcut]:
        """获取快捷键列表"""
        if only_custom:
            result = self.custom_shortcuts.copy()
        elif only_builtin:
            result = self.builtin_shortcuts.copy()
        else:
            result = self.shortcuts.copy()
        
        if search_term:
            search_term = search_term.upper()
            result = [s for s in result if search_term in s.alias or search_term in s.command]
        
        return result
    
    def get_shortcut_by_alias(self, alias: str) -> Optional[Shortcut]:
        """根据别名获取快捷键"""
        alias = alias.upper().strip()
        for s in self.shortcuts:
            if s.alias == alias:
                return s
        return None
    
    def add_shortcut(self, alias: str, command: str, description: str = "") -> Tuple[bool, str]:
        """添加快捷键"""
        alias = alias.upper().strip()
        command = command.upper().strip()
        
        if not alias or not command:
            return False, "别名和命令不能为空"
        
        if not re.match(r'^[A-Z0-9_]+$', alias):
            return False, "别名只能包含字母、数字和下划线"
        
        existing = self.get_shortcut_by_alias(alias)
        if existing:
            return False, f"别名 '{alias}' 已存在，对应命令为 '{existing.command}'"
        
        self._backup()
        
        new_line = f"{alias},*{command}"
        if description:
            new_line += f"  ; {description}"
        new_line += "\n"
        
        # 找到用户自定义区域或文件末尾
        insert_pos = len(self.file_content)
        user_section_found = False
        
        for i, line in enumerate(self.file_content):
            if 'user defined' in line.lower() or '用户定义' in line or '自定义' in line:
                user_section_found = True
                # 找到该区域的最后一行
                for j in range(i + 1, len(self.file_content)):
                    next_line = self.file_content[j].strip()
                    if next_line and not next_line.startswith(';'):
                        if re.match(r'^(\w+)\s*,\s*\*', next_line):
                            insert_pos = j + 1
                break
        
        # 如果没有找到用户定义区域，创建一个
        if not user_section_found:
            self.file_content.append("\n; -- User defined shortcuts --\n")
            insert_pos = len(self.file_content)
        
        self.file_content.insert(insert_pos, new_line)
        
        success = self._save_file()
        if success:
            self._parse_shortcuts()
            return True, f"成功添加快捷键：{alias} → {command}"
        else:
            return False, "保存文件失败"
    
    def update_shortcut(self, alias: str, new_command: str = None, 
                       new_description: str = None) -> Tuple[bool, str]:
        """修改快捷键"""
        alias = alias.upper().strip()
        shortcut = self.get_shortcut_by_alias(alias)
        
        if not shortcut:
            return False, f"找不到别名 '{alias}'"
        
        self._backup()
        
        command = new_command.upper().strip() if new_command else shortcut.command
        description = new_description if new_description is not None else shortcut.description
        
        new_line = f"{alias},*{command}"
        if description:
            new_line += f"  ; {description}"
        new_line += "\n"
        
        self.file_content[shortcut.line_num - 1] = new_line
        
        success = self._save_file()
        if success:
            self._parse_shortcuts()
            return True, f"成功更新快捷键：{alias} → {command}"
        else:
            return False, "保存文件失败"
    
    def delete_shortcut(self, alias: str) -> Tuple[bool, str]:
        """删除快捷键"""
        alias = alias.upper().strip()
        shortcut = self.get_shortcut_by_alias(alias)
        
        if not shortcut:
            return False, f"找不到别名 '{alias}'"
        
        self._backup()
        
        del self.file_content[shortcut.line_num - 1]
        
        success = self._save_file()
        if success:
            self._parse_shortcuts()
            return True, f"成功删除快捷键：{alias}"
        else:
            return False, "保存文件失败"
    
    def _backup(self):
        """创建备份"""
        if self.pgp_path:
            backup_path = f"{self.pgp_path}.backup.{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            shutil.copy2(self.pgp_path, backup_path)
            return backup_path
        return None
    
    def _save_file(self) -> bool:
        """保存文件"""
        try:
            with open(self.pgp_path, 'w', encoding='utf-8') as f:
                f.writelines(self.file_content)
            return True
        except Exception as e:
            print(f"保存失败: {e}")
            return False
    
    # ==================== 用户配置管理 ====================
    
    def create_profile(self, name: str, description: str = "") -> Tuple[bool, str]:
        """创建新的用户配置"""
        profile_path = self.profiles_dir / f"{name}.json"
        
        if profile_path.exists():
            return False, f"用户配置 '{name}' 已存在"
        
        # 收集当前所有自定义快捷键
        custom_shortcuts = [
            {
                "alias": s.alias,
                "command": s.command,
                "description": s.description
            }
            for s in self.custom_shortcuts
        ]
        
        profile = {
            "name": name,
            "description": description,
            "created_at": datetime.now().isoformat(),
            "modified_at": datetime.now().isoformat(),
            "shortcuts": custom_shortcuts
        }
        
        try:
            with open(profile_path, 'w', encoding='utf-8') as f:
                json.dump(profile, f, ensure_ascii=False, indent=2)
            return True, f"成功创建用户配置：{name}"
        except Exception as e:
            return False, f"创建失败: {e}"
    
    def load_profile(self, name: str) -> Tuple[bool, str]:
        """加载用户配置到当前 AutoCAD"""
        profile_path = self.profiles_dir / f"{name}.json"
        
        if not profile_path.exists():
            return False, f"找不到用户配置 '{name}'"
        
        try:
            with open(profile_path, 'r', encoding='utf-8') as f:
                profile = json.load(f)
        except Exception as e:
            return False, f"读取配置失败: {e}"
        
        # 备份当前配置
        self._backup()
        
        # 删除所有现有自定义快捷键
        shortcuts_to_remove = [s.alias for s in self.custom_shortcuts]
        for alias in shortcuts_to_remove:
            shortcut = self.get_shortcut_by_alias(alias)
            if shortcut and not shortcut.is_builtin:
                del self.file_content[shortcut.line_num - 1]
        
        # 添加配置中的快捷键
        user_section_added = False
        for shortcut_data in profile.get("shortcuts", []):
            alias = shortcut_data["alias"]
            command = shortcut_data["command"]
            description = shortcut_data.get("description", "")
            
            # 检查是否已存在（内置）
            existing = self.get_shortcut_by_alias(alias)
            if existing and existing.is_builtin:
                continue  # 跳过内置快捷键
            
            new_line = f"{alias},*{command}"
            if description:
                new_line += f"  ; {description}"
            new_line += "\n"
            
            if not user_section_added:
                self.file_content.append("\n; -- User defined shortcuts --\n")
                user_section_added = True
            
            self.file_content.append(new_line)
        
        # 更新修改时间
        profile["modified_at"] = datetime.now().isoformat()
        with open(profile_path, 'w', encoding='utf-8') as f:
            json.dump(profile, f, ensure_ascii=False, indent=2)
        
        success = self._save_file()
        if success:
            self._parse_shortcuts()
            return True, f"成功加载用户配置 '{name}'，已应用 {len(profile.get('shortcuts', []))} 个自定义快捷键"
        else:
            return False, "保存文件失败"
    
    def update_profile(self, name: str) -> Tuple[bool, str]:
        """更新用户配置（保存当前自定义快捷键）"""
        profile_path = self.profiles_dir / f"{name}.json"
        
        if not profile_path.exists():
            return False, f"用户配置 '{name}' 不存在，请先创建"
        
        try:
            with open(profile_path, 'r', encoding='utf-8') as f:
                profile = json.load(f)
        except Exception as e:
            return False, f"读取配置失败: {e}"
        
        # 更新快捷键列表
        custom_shortcuts = [
            {
                "alias": s.alias,
                "command": s.command,
                "description": s.description
            }
            for s in self.custom_shortcuts
        ]
        
        profile["shortcuts"] = custom_shortcuts
        profile["modified_at"] = datetime.now().isoformat()
        
        try:
            with open(profile_path, 'w', encoding='utf-8') as f:
                json.dump(profile, f, ensure_ascii=False, indent=2)
            return True, f"成功更新用户配置 '{name}'，已保存 {len(custom_shortcuts)} 个自定义快捷键"
        except Exception as e:
            return False, f"保存失败: {e}"
    
    def delete_profile(self, name: str) -> Tuple[bool, str]:
        """删除用户配置"""
        profile_path = self.profiles_dir / f"{name}.json"
        
        if not profile_path.exists():
            return False, f"找不到用户配置 '{name}'"
        
        try:
            profile_path.unlink()
            return True, f"成功删除用户配置 '{name}'"
        except Exception as e:
            return False, f"删除失败: {e}"
    
    def list_profiles(self) -> List[Dict]:
        """列出所有用户配置"""
        profiles = []
        
        for profile_file in self.profiles_dir.glob("*.json"):
            try:
                with open(profile_file, 'r', encoding='utf-8') as f:
                    profile = json.load(f)
                    profiles.append({
                        "name": profile.get("name", profile_file.stem),
                        "description": profile.get("description", ""),
                        "created_at": profile.get("created_at", ""),
                        "modified_at": profile.get("modified_at", ""),
                        "shortcut_count": len(profile.get("shortcuts", []))
                    })
            except:
                pass
        
        return sorted(profiles, key=lambda x: x["name"])
    
    def export_profile(self, name: str, output_path: str) -> Tuple[bool, str]:
        """导出用户配置到指定路径"""
        profile_path = self.profiles_dir / f"{name}.json"
        
        if not profile_path.exists():
            return False, f"找不到用户配置 '{name}'"
        
        try:
            shutil.copy2(profile_path, output_path)
            return True, f"成功导出到: {output_path}"
        except Exception as e:
            return False, f"导出失败: {e}"
    
    def import_profile(self, file_path: str) -> Tuple[bool, str]:
        """从文件导入用户配置"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                profile = json.load(f)
            
            name = profile.get("name", Path(file_path).stem)
            profile_path = self.profiles_dir / f"{name}.json"
            
            # 如果已存在，添加数字后缀
            counter = 1
            original_name = name
            while profile_path.exists():
                name = f"{original_name}_{counter}"
                profile_path = self.profiles_dir / f"{name}.json"
                counter += 1
            
            profile["name"] = name
            
            with open(profile_path, 'w', encoding='utf-8') as f:
                json.dump(profile, f, ensure_ascii=False, indent=2)
            
            return True, f"成功导入用户配置: {name}"
        except Exception as e:
            return False, f"导入失败: {e}"
    
    def get_stats(self) -> Dict:
        """获取统计信息"""
        total = len(self.shortcuts)
        builtin = len(self.builtin_shortcuts)
        custom = len(self.custom_shortcuts)
        return {
            'total': total,
            'builtin': builtin,
            'custom': custom,
            'file_path': self.pgp_path,
            'version': self._extract_version_from_path(self.pgp_path) if self.pgp_path else "未找到"
        }


def clear_screen():
    """清屏"""
    os.system('cls' if os.name == 'nt' else 'clear')


def print_header():
    """打印标题"""
    print("=" * 60)
    print("       AutoCAD 快捷键管理工具 v2.3")
    print("       开源项目 | 作者: jjyy7783")
    print("       支持 AutoCAD 2004-2026 | 用户配置系统")
    print("       新增：解决插件快捷键冲突")
    print("=" * 60)


def print_main_menu():
    """打印主菜单"""
    print("\n【主菜单】")
    print("-" * 40)
    print("  1. 查看快捷键")
    print("  2. 管理快捷键 (添加/修改/删除)")
    print("  3. 用户配置管理")
    print("  4. 切换 AutoCAD 版本")
    print("  5. 统计信息")
    print("  6. 重新加载快捷键 (解决插件冲突)")
    print("  0. 退出")
    print("-" * 40)


def print_shortcuts(shortcuts: List[Shortcut], title: str = "快捷键列表"):
    """打印快捷键列表"""
    print(f"\n{'='*70}")
    print(f"  {title} (共 {len(shortcuts)} 个)")
    print(f"{'='*70}")
    print(f"{'序号':<6}{'别名':<12}{'命令':<25}{'类型':<8}{'描述'}")
    print("-" * 70)
    
    for i, s in enumerate(shortcuts, 1):
        type_str = "内置" if s.is_builtin else "自定义"
        desc = s.description[:15] if s.description else ""
        cmd = s.command[:23] if len(s.command) > 23 else s.command
        print(f"{i:<6}{s.alias:<12}{cmd:<25}{type_str:<8}{desc}")
    
    print("=" * 70)


def select_autocad_version(manager: AutoCADShortcutManager) -> Optional[str]:
    """选择 AutoCAD 版本"""
    versions = manager.find_all_acad_pgp()
    
    if not versions:
        print("\n⚠️ 未找到任何 AutoCAD 安装！")
        return None
    
    print(f"\n找到 {len(versions)} 个 AutoCAD 版本：")
    print("-" * 80)
    print(f"{'序号':<6}{'版本':<30}{'路径'}")
    print("-" * 80)
    
    for i, (path, version) in enumerate(versions, 1):
        path_display = path if len(path) < 45 else "..." + path[-42:]
        print(f"{i:<6}{version:<30}{path_display}")
    
    print("-" * 80)
    
    try:
        choice = int(input(f"请选择 [1-{len(versions)}]: ").strip())
        if 1 <= choice <= len(versions):
            return versions[choice - 1][0]
    except ValueError:
        pass
    
    return None


def select_command(manager: AutoCADShortcutManager) -> Optional[Tuple[str, str]]:
    """
    智能命令选择器
    :return: (命令名, 描述) 或 None
    """
    print("\n【选择命令】")
    print("输入命令名或中文描述关键词进行搜索，直接回车显示常用命令")
    
    while True:
        keyword = input("\n搜索关键词 [回车显示常用命令，输入 ? 帮助，q 返回]: ").strip()
        
        if keyword.lower() == 'q':
            return None
        
        if keyword.lower() == '?':
            print("\n帮助说明：")
            print("  - 输入英文关键词（如 COPY）搜索命令")
            print("  - 输入中文关键词（如 复制）搜索命令")
            print("  - 直接回车显示所有常用命令")
            print("  - 输入数字选择对应命令")
            continue
        
        # 搜索命令
        if keyword:
            results = search_commands(keyword, limit=30)
        else:
            # 显示所有常用命令（按类别分组）
            results = list(AUTOCAD_COMMANDS.items())[:50]
        
        if not results:
            print("未找到匹配的命令，请换一个关键词")
            continue
        
        # 显示搜索结果
        print(f"\n找到 {len(results)} 个命令：")
        print("-" * 50)
        
        # 分两列显示
        half = (len(results) + 1) // 2
        for i in range(half):
            left = f"{i+1}. {results[i][0]:<15} {results[i][1]}"
            right_idx = i + half
            if right_idx < len(results):
                right = f"{right_idx+1}. {results[right_idx][0]:<15} {results[right_idx][1]}"
                print(f"{left:<30} {right}")
            else:
                print(left)
        
        print("-" * 50)
        
        # 选择命令
        try:
            choice = input("\n请选择命令序号 [回车重新搜索]: ").strip()
            if not choice:
                continue
            
            idx = int(choice) - 1
            if 0 <= idx < len(results):
                cmd, desc = results[idx]
                # 如果是别名，尝试获取原始命令名
                if "(别名)" in desc:
                    # 找到对应的非别名命令
                    base_desc = desc.replace("(别名)", "")
                    for c, d in AUTOCAD_COMMANDS.items():
                        if d == base_desc and "(别名)" not in d:
                            cmd = c
                            desc = base_desc
                            break
                return (cmd, desc)
            else:
                print("序号无效，请重新选择")
        except ValueError:
            print("请输入数字序号")


def manage_shortcuts(manager: AutoCADShortcutManager):
    """管理快捷键子菜单"""
    while True:
        print("\n【快捷键管理】")
        print("-" * 40)
        print("  1. 添加快捷键（智能选择命令）")
        print("  2. 添加快捷键（手动输入命令）")
        print("  3. 修改快捷键")
        print("  4. 删除快捷键")
        print("  5. 批量添加（从文本）")
        print("  0. 返回主菜单")
        print("-" * 40)
        
        choice = input("请选择 [0-5]: ").strip()
        
        if choice == '0':
            break
        
        elif choice == '1':
            # 智能选择命令
            print("\n--- 添加快捷键（智能选择） ---")
            alias = input("快捷命令 (如: C): ").strip().upper()
            if not alias:
                continue
            
            # 检查是否已存在
            existing = manager.get_shortcut_by_alias(alias)
            if existing:
                # 自动处理：旧别名改为其命令全名，腾出别名给新命令
                old_cmd = existing.command
                print(f"⚠️ 别名 '{alias}' 已被占用: {alias} → {old_cmd}")
                success_swap, msg_swap = manager.add_shortcut(old_cmd, old_cmd, existing.description or f"原 {alias} 命令")
                if success_swap:
                    success_del, _ = manager.delete_shortcut(alias)
                    if success_del:
                        print(f"  ↳ 已自动处理: 原命令 '{old_cmd}' 已获得新别名 '{old_cmd}'，别名 '{alias}' 已释放")
                    else:
                        print(f"  ↳ ⚠️ 释放别名失败，请手动处理")
                        continue
                else:
                    print(f"  ↳ ⚠️ 自动处理失败: {msg_swap}")
                    continue
            
            # 选择命令
            result = select_command(manager)
            if result:
                cmd, desc = result
                print(f"\n将添加: {alias} → {cmd} ({desc})")
                confirm = input("确认添加? [Y/n]: ").strip().lower()
                if confirm in ('', 'y', 'yes'):
                    success, msg = manager.add_shortcut(alias, cmd, desc)
                    print(f"{'✓' if success else '✗'} {msg}")
        
        elif choice == '2':
            # 手动输入命令
            print("\n--- 添加快捷键（手动输入） ---")
            alias = input("快捷命令 (如: XX): ").strip()
            if not alias:
                continue
            command = input("完整命令 (如: EXPLODE): ").strip()
            desc = input("描述 (可选): ").strip()
            
            # 检查是否已存在，自动处理冲突
            existing = manager.get_shortcut_by_alias(alias)
            if existing:
                old_cmd = existing.command
                print(f"⚠️ 别名 '{alias}' 已被占用: {alias} → {old_cmd}")
                success_swap, msg_swap = manager.add_shortcut(old_cmd, old_cmd, existing.description or f"原 {alias} 命令")
                if success_swap:
                    success_del, _ = manager.delete_shortcut(alias)
                    if success_del:
                        print(f"  ↳ 已自动处理: 原命令 '{old_cmd}' 已获得新别名 '{old_cmd}'，别名 '{alias}' 已释放")
                    else:
                        print(f"  ↳ ⚠️ 释放别名失败，请手动处理")
                        continue
                else:
                    print(f"  ↳ ⚠️ 自动处理失败: {msg_swap}")
                    continue
            
            confirm = input(f"确认添加 {alias} → {command}? [Y/n]: ").strip().lower()
            if confirm in ('', 'y', 'yes'):
                success, msg = manager.add_shortcut(alias, command, desc)
                print(f"{'✓' if success else '✗'} {msg}")
        
        elif choice == '3':
            print("\n--- 修改快捷键 ---")
            alias = input("要修改的别名: ").strip()
            shortcut = manager.get_shortcut_by_alias(alias)
            if not shortcut:
                print(f"✗ 找不到别名 '{alias}'")
                continue
            
            print(f"当前: {shortcut.alias} → {shortcut.command}")
            
            # 提供选择：修改命令还是修改别名
            print("\n修改选项:")
            print("  1. 修改命令（智能选择）")
            print("  2. 修改命令（手动输入）")
            print("  3. 修改别名")
            sub_choice = input("请选择 [1-3]: ").strip()
            
            if sub_choice == '1':
                # 智能选择新命令
                result = select_command(manager)
                if result:
                    new_cmd, new_desc = result
                    confirm = input(f"确认修改为 {alias} → {new_cmd}? [Y/n]: ").strip().lower()
                    if confirm in ('', 'y', 'yes'):
                        success, msg = manager.update_shortcut(alias, new_cmd, new_desc)
                        print(f"{'✓' if success else '✗'} {msg}")
            elif sub_choice == '2':
                new_cmd = input(f"新命令 [回车保持 '{shortcut.command}']: ").strip()
                new_desc = input(f"新描述 [回车保持 '{shortcut.description}']: ").strip()
                confirm = input("确认修改? [Y/n]: ").strip().lower()
                if confirm in ('', 'y', 'yes'):
                    success, msg = manager.update_shortcut(
                        alias,
                        new_cmd if new_cmd else None,
                        new_desc if new_desc else None
                    )
                    print(f"{'✓' if success else '✗'} {msg}")
            elif sub_choice == '3':
                new_alias = input(f"新别名 [回车取消]: ").strip().upper()
                if new_alias and new_alias != alias:
                    # 先添加新的，再删除旧的
                    success, msg = manager.add_shortcut(new_alias, shortcut.command, shortcut.description)
                    if success:
                        manager.delete_shortcut(alias)
                        print(f"✓ 已将 {alias} 改为 {new_alias}")
                    else:
                        print(f"✗ {msg}")
        
        elif choice == '4':
            print("\n--- 删除快捷键 ---")
            alias = input("要删除的别名: ").strip()
            shortcut = manager.get_shortcut_by_alias(alias)
            if not shortcut:
                print(f"✗ 找不到别名 '{alias}'")
                continue
            
            if shortcut.is_builtin:
                print(f"⚠️ 警告: '{alias}' 是内置快捷键！")
            
            confirm = input(f"确认删除 {alias}? [y/N]: ").strip().lower()
            if confirm in ('y', 'yes'):
                success, msg = manager.delete_shortcut(alias)
                print(f"{'✓' if success else '✗'} {msg}")
        
        elif choice == '5':
            print("\n--- 批量添加 ---")
            print("格式: 别名,命令,描述（每行一个）")
            print("示例: XX,EXPLODE,分解命令")
            print("输入空行结束")
            print("-" * 40)
            
            lines = []
            while True:
                line = input("> ").strip()
                if not line:
                    break
                lines.append(line)
            
            added = 0
            failed = []
            for line in lines:
                parts = line.split(',')
                if len(parts) >= 2:
                    alias, command = parts[0].strip(), parts[1].strip()
                    desc = parts[2].strip() if len(parts) > 2 else ""
                    success, msg = manager.add_shortcut(alias, command, desc)
                    if success:
                        added += 1
                    else:
                        failed.append(f"{alias}: {msg}")
            
            print(f"\n✓ 成功添加 {added} 个快捷键")
            if failed:
                print(f"✗ 失败 {len(failed)} 个:")
                for f in failed:
                    print(f"  - {f}")
        
        input("\n按回车键继续...")


def manage_profiles(manager: AutoCADShortcutManager):
    """用户配置管理子菜单"""
    while True:
        print("\n【用户配置管理】")
        print("-" * 40)
        print("  1. 查看所有配置")
        print("  2. 创建新配置")
        print("  3. 加载配置到当前 AutoCAD")
        print("  4. 更新配置（保存当前快捷键）")
        print("  5. 删除配置")
        print("  6. 导出配置（用于其他电脑）")
        print("  7. 导入配置")
        print("  0. 返回主菜单")
        print("-" * 40)
        
        choice = input("请选择 [0-7]: ").strip()
        
        if choice == '0':
            break
        
        elif choice == '1':
            profiles = manager.list_profiles()
            if not profiles:
                print("\n暂无用户配置")
            else:
                print(f"\n{'='*70}")
                print(f"  用户配置列表 (共 {len(profiles)} 个)")
                print(f"{'='*70}")
                print(f"{'配置名':<20}{'描述':<25}{'快捷键数':<10}{'修改时间'}")
                print("-" * 70)
                
                for p in profiles:
                    desc = p['description'][:23] if p['description'] else ""
                    modified = p['modified_at'][:10] if p['modified_at'] else ""
                    print(f"{p['name']:<20}{desc:<25}{p['shortcut_count']:<10}{modified}")
                print("=" * 70)
        
        elif choice == '2':
            print("\n--- 创建新配置 ---")
            name = input("配置名称: ").strip()
            if not name:
                print("✗ 名称不能为空")
                continue
            
            desc = input("配置描述 (可选): ").strip()
            
            success, msg = manager.create_profile(name, desc)
            print(f"{'✓' if success else '✗'} {msg}")
        
        elif choice == '3':
            profiles = manager.list_profiles()
            if not profiles:
                print("\n暂无用户配置可加载")
                continue
            
            print("\n可用配置:")
            for i, p in enumerate(profiles, 1):
                print(f"  {i}. {p['name']} ({p['shortcut_count']} 个快捷键)")
            
            try:
                idx = int(input("请选择配置编号: ").strip()) - 1
                if 0 <= idx < len(profiles):
                    name = profiles[idx]['name']
                    print(f"\n正在加载配置 '{name}'...")
                    success, msg = manager.load_profile(name)
                    print(f"{'✓' if success else '✗'} {msg}")
                    if success:
                        print("⚠️ 请重启 AutoCAD 使更改生效")
                else:
                    print("✗ 无效选择")
            except ValueError:
                print("✗ 请输入数字")
        
        elif choice == '4':
            profiles = manager.list_profiles()
            if not profiles:
                print("\n暂无用户配置，请先创建")
                continue
            
            print("\n选择要更新的配置:")
            for i, p in enumerate(profiles, 1):
                print(f"  {i}. {p['name']}")
            
            try:
                idx = int(input("请选择配置编号: ").strip()) - 1
                if 0 <= idx < len(profiles):
                    name = profiles[idx]['name']
                    success, msg = manager.update_profile(name)
                    print(f"{'✓' if success else '✗'} {msg}")
                else:
                    print("✗ 无效选择")
            except ValueError:
                print("✗ 请输入数字")
        
        elif choice == '5':
            profiles = manager.list_profiles()
            if not profiles:
                print("\n暂无用户配置")
                continue
            
            print("\n选择要删除的配置:")
            for i, p in enumerate(profiles, 1):
                print(f"  {i}. {p['name']}")
            
            try:
                idx = int(input("请选择配置编号: ").strip()) - 1
                if 0 <= idx < len(profiles):
                    name = profiles[idx]['name']
                    confirm = input(f"确认删除配置 '{name}'? [y/N]: ").strip().lower()
                    if confirm in ('y', 'yes'):
                        success, msg = manager.delete_profile(name)
                        print(f"{'✓' if success else '✗'} {msg}")
                else:
                    print("✗ 无效选择")
            except ValueError:
                print("✗ 请输入数字")
        
        elif choice == '6':
            profiles = manager.list_profiles()
            if not profiles:
                print("\n暂无用户配置")
                continue
            
            print("\n选择要导出的配置:")
            for i, p in enumerate(profiles, 1):
                print(f"  {i}. {p['name']}")
            
            try:
                idx = int(input("请选择配置编号: ").strip()) - 1
                if 0 <= idx < len(profiles):
                    name = profiles[idx]['name']
                    output = input("导出文件路径 (如: D:\\我的配置.json): ").strip()
                    if output:
                        success, msg = manager.export_profile(name, output)
                        print(f"{'✓' if success else '✗'} {msg}")
                        if success:
                            print(f"💡 提示: 将此文件复制到其他电脑，使用'导入配置'功能即可应用")
                else:
                    print("✗ 无效选择")
            except ValueError:
                print("✗ 请输入数字")
        
        elif choice == '7':
            print("\n--- 导入配置 ---")
            file_path = input("配置文件路径: ").strip()
            if file_path and os.path.exists(file_path):
                success, msg = manager.import_profile(file_path)
                print(f"{'✓' if success else '✗'} {msg}")
                if success:
                    print("💡 提示: 使用'加载配置到当前 AutoCAD'功能应用此配置")
            else:
                print("✗ 文件不存在")
        
        input("\n按回车键继续...")


def view_shortcuts(manager: AutoCADShortcutManager):
    """查看快捷键子菜单"""
    while True:
        print("\n【查看快捷键】")
        print("-" * 40)
        print("  1. 查看全部")
        print("  2. 查看内置快捷键")
        print("  3. 查看自定义快捷键")
        print("  4. 搜索快捷键")
        print("  5. 导出列表到文件")
        print("  0. 返回主菜单")
        print("-" * 40)
        
        choice = input("请选择 [0-5]: ").strip()
        
        if choice == '0':
            break
        
        elif choice == '1':
            shortcuts = manager.get_shortcuts()
            print_shortcuts(shortcuts, "全部快捷键")
        
        elif choice == '2':
            shortcuts = manager.get_shortcuts(only_builtin=True)
            print_shortcuts(shortcuts, "内置快捷键")
        
        elif choice == '3':
            shortcuts = manager.get_shortcuts(only_custom=True)
            print_shortcuts(shortcuts, "自定义快捷键")
        
        elif choice == '4':
            term = input("输入搜索关键词: ").strip()
            shortcuts = manager.get_shortcuts(search_term=term)
            print_shortcuts(shortcuts, f"搜索结果: '{term}'")
        
        elif choice == '5':
            output = input("导出文件路径 [默认: shortcuts_list.txt]: ").strip()
            if not output:
                output = "shortcuts_list.txt"
            
            try:
                with open(output, 'w', encoding='utf-8') as f:
                    f.write("# AutoCAD 快捷键列表\n")
                    f.write(f"# 导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"# 来源: {manager.pgp_path}\n\n")
                    f.write("类型 | 别名 | 命令 | 描述\n")
                    f.write("-" * 60 + "\n")
                    
                    for s in manager.shortcuts:
                        type_str = "内置" if s.is_builtin else "自定义"
                        desc = s.description if s.description else ""
                        f.write(f"{type_str} | {s.alias} | {s.command} | {desc}\n")
                
                print(f"✓ 已导出到: {os.path.abspath(output)}")
            except Exception as e:
                print(f"✗ 导出失败: {e}")
        
        input("\n按回车键继续...")


def main():
    """主程序"""
    clear_screen()
    print_header()
    print("\n正在初始化...")
    
    manager = AutoCADShortcutManager()
    
    # 如果没有找到，让用户选择
    if not manager.pgp_path:
        print("\n⚠️ 未自动找到 AutoCAD 安装")
        path = select_autocad_version(manager)
        if path:
            manager = AutoCADShortcutManager(path)
        else:
            print("\n请手动输入 acad.pgp 文件路径:")
            custom_path = input("> ").strip()
            if custom_path and os.path.exists(custom_path):
                manager = AutoCADShortcutManager(custom_path)
            else:
                print("路径无效，程序退出")
                return
    
    print(f"\n✓ 当前 AutoCAD: {manager.get_stats()['version']}")
    print(f"✓ 配置文件: {manager.pgp_path}")
    print(f"✓ 用户配置目录: {manager.profiles_dir}")
    
    while True:
        print_main_menu()
        choice = input("请选择操作 [0-5]: ").strip()
        
        if choice == '0':
            print("\n感谢使用，再见！")
            break
        
        elif choice == '1':
            view_shortcuts(manager)
        
        elif choice == '2':
            manage_shortcuts(manager)
        
        elif choice == '3':
            manage_profiles(manager)
        
        elif choice == '4':
            print("\n【切换 AutoCAD 版本】")
            path = select_autocad_version(manager)
            if path:
                manager = AutoCADShortcutManager(path)
                print(f"\n✓ 已切换到: {manager.get_stats()['version']}")
                print(f"✓ 配置文件: {manager.pgp_path}")
        
        elif choice == '5':
            stats = manager.get_stats()
            print(f"\n{'='*60}")
            print("              统计信息")
            print(f"{'='*60}")
            print(f"AutoCAD 版本: {stats['version']}")
            print(f"配置文件路径: {stats['file_path']}")
            print(f"总快捷键数:   {stats['total']}")
            print(f"  - 内置:     {stats['builtin']}")
            print(f"  - 自定义:   {stats['custom']}")
            print(f"{'='*60}")
            input("\n按回车键继续...")
        
        elif choice == '6':
            reload_shortcuts_in_autocad(manager)
        
        else:
            print("无效选择，请重试")


def reload_shortcuts_in_autocad(manager: AutoCADShortcutManager):
    """
    重新加载快捷键到 AutoCAD
    解决天正等插件冲突的问题
    """
    print("\n【重新加载快捷键】")
    print("=" * 60)
    print("说明：当天正等插件覆盖了你的快捷键时，使用此功能重新加载")
    print("=" * 60)
    
    print("\n选择加载方式：")
    print("  1. 快速重新加载（生成 SCR 脚本）")
    print("  2. 创建自动加载脚本（启动时自动重新加载）")
    print("  3. 查看详细教程")
    print("  0. 返回")
    
    choice = input("\n请选择 [0-3]: ").strip()
    
    if choice == '0':
        return
    
    elif choice == '1':
        # 方法1：生成 SCR 脚本
        print("\n正在生成脚本...")
        
        scr_path = os.path.join(os.path.dirname(manager.pgp_path), "reload_pgp.scr")
        try:
            with open(scr_path, 'w', encoding='utf-8') as f:
                f.write("(command \"_.REINIT\" 16)\n")
                f.write("(princ)\n")
            
            print(f"\n✓ 已生成脚本文件:")
            print(f"  {scr_path}")
            print("\n在 AutoCAD 中执行:")
            print("  1. 输入命令: SCRIPT")
            print(f"  2. 选择文件: reload_pgp.scr")
            print("\n或直接在命令行输入: (command \"REINIT\" 16)")
            
        except Exception as e:
            print(f"✗ 生成脚本失败: {e}")
    
    elif choice == '2':
        # 方法2：创建自动加载 LISP 脚本
        print("\n【创建自动加载脚本】")
        print("此脚本会在 AutoCAD 启动时自动重新加载你的快捷键")
        
        lisp_path = os.path.join(os.path.dirname(manager.pgp_path), "acad_reload_pgp.lsp")
        try:
            with open(lisp_path, 'w', encoding='utf-8') as f:
                f.write(";; AutoCAD 快捷键自动重新加载脚本\n")
                f.write(";; 解决天正等插件覆盖快捷键的问题\n")
                f.write(";; 生成时间: " + datetime.now().strftime('%Y-%m-%d %H:%M:%S') + "\n\n")
                f.write("(defun S::STARTUP ()\n")
                f.write("  ;; 延迟执行，确保所有插件加载完成\n")
                f.write("  (command \"_.DELAY\" 1000)\n")  # 延迟1秒
                f.write("  ;; 重新初始化 PGP 文件\n")
                f.write("  (command \"_.REINIT\" 16)\n")
                f.write("  (princ \"\\n[快捷键管理工具] 已重新加载快捷键配置\\n\")\n")
                f.write("  (princ)\n")
                f.write(")\n")
            
            print(f"\n✓ 已生成自动加载脚本:")
            print(f"  {lisp_path}")
            print("\n【安装步骤】")
            print("  方法A - 添加到启动组:")
            print("    1. 在 AutoCAD 中输入命令: APPLOAD")
            print("    2. 点击 \"启动组\" -> \"内容\"")
            print("    3. 点击 \"添加\"，选择此文件")
            print("    4. 以后每次启动 AutoCAD 都会自动重新加载快捷键")
            print("\n  方法B - 复制到支持目录:")
            print(f"    将文件复制到: {os.path.dirname(manager.pgp_path)}")
            print("    文件名改为: acad.lsp (如果没有此文件)")
            
        except Exception as e:
            print(f"✗ 生成脚本失败: {e}")
    
    elif choice == '3':
        # 详细教程
        print("\n" + "=" * 60)
        print("手动重新加载快捷键教程")
        print("=" * 60)
        print("""
【方法一】使用 REINIT 命令（最简单）
  1. 在 AutoCAD 命令行输入: REINIT
  2. 勾选 "PGP 文件" 选项
  3. 点击 "确定"
  快捷键立即生效！

【方法二】使用命令表达式
  在命令行直接输入: (command "REINIT" 16)
  
【方法三】关闭重开 AutoCAD
  关闭 AutoCAD 后重新打开，快捷键会自动加载

【解决天正插件冲突的终极方案】
  天正等插件在启动时会加载自己的快捷键覆盖你的设置。
  解决方法：
  
  1. 找到天正的 PGP 文件并重命名
     位置通常在: C:\\TArch\\TArch20XX\\Sys\\*.pgp
     重命名为: *.pgp.bak
     
  2. 在天正设置中禁用快捷键
     打开天正 -> 设置 -> 快捷键 -> 禁用

  3. 使用本工具的"自动加载脚本"功能
     让你的快捷键在插件加载后重新覆盖回来
""")
        print("=" * 60)
    
    input("\n按回车键继续...")


if __name__ == '__main__':
    main()
