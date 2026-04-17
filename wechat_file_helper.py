# -*- coding: utf-8 -*-
"""
PC版微信文件消息路径提取工具
功能：监控微信聊天窗口，当用户选中或右键点击文件类型消息时，
      自动获取文件的绝对保存路径并复制到剪贴板
"""

import uiautomation as auto
import pyperclip
import time
import re
import os
import sys

# =============================================================================
# 问题1修复：正确导入和检测 pywin32
# =============================================================================
WIN32_AVAILABLE = False
try:
    import win32api
    import win32con
    import win32gui
    WIN32_AVAILABLE = True
    print("[系统] pywin32 已成功加载")
except ImportError as e:
    print(f"[系统] pywin32 导入失败: {e}")
except Exception as e:
    print(f"[系统] pywin32 加载出错: {e}")

# 微信窗口的类名和标题
WECHAT_CLASS_NAME = 'WeChatMainWndForPC'
WECHAT_TITLE = '微信'

# 文件消息的特征关键词 - 更严格的匹配
FILE_KEYWORDS = ['文件', '文件:', '已接收文件', '接收文件', '另存为', '打开文件夹', '在文件夹中显示']

# 文件大小模式 - 用于识别文件消息
FILE_SIZE_PATTERN = r'\(\d+\.?\d*\s*[KMGT]?B\)'

# 微信默认文件保存路径
DEFAULT_WECHAT_FILE_PATH = os.path.join(
    os.environ.get('USERPROFILE', ''),
    'Documents', 'WeChat Files'
)

# 调试模式
DEBUG_MODE = True


class WeChatFileMonitor:
    """微信文件消息监控类"""
    
    def __init__(self):
        self.wechat_window = None
        self.last_clicked_file = None
        self.monitoring = False
        self.debug_mode = DEBUG_MODE
        self.wechat_file_dir = None  # 缓存微信文件目录
        
    def find_wechat_window(self):
        """查找微信主窗口"""
        try:
            print("正在查找微信窗口...")
            
            # 方法1: 通过类名查找
            try:
                window = auto.WindowControl(ClassName=WECHAT_CLASS_NAME)
                if window.Exists(2):
                    self.wechat_window = window
                    print(f"✓ 通过类名找到微信窗口: {WECHAT_CLASS_NAME}")
                    return True
            except Exception as e:
                if self.debug_mode:
                    print(f"  [调试] 类名查找失败: {e}")
            
            # 方法2: 通过标题查找
            try:
                window = auto.WindowControl(Name=WECHAT_TITLE)
                if window.Exists(2):
                    self.wechat_window = window
                    print(f"✓ 通过标题找到微信窗口: {WECHAT_TITLE}")
                    return True
            except Exception as e:
                if self.debug_mode:
                    print(f"  [调试] 标题查找失败: {e}")
            
            # 方法3: 枚举所有顶层窗口（如果有 win32gui）
            if WIN32_AVAILABLE:
                print("尝试枚举所有窗口查找微信...")
                try:
                    def callback(hwnd, windows):
                        if win32gui.IsWindowVisible(hwnd):
                            title = win32gui.GetWindowText(hwnd)
                            class_name = win32gui.GetClassName(hwnd)
                            if WECHAT_TITLE in title or WECHAT_CLASS_NAME in class_name:
                                windows.append((hwnd, title, class_name))
                        return True
                    
                    windows = []
                    win32gui.EnumWindows(callback, windows)
                    if windows:
                        hwnd, title, class_name = windows[0]
                        print(f"✓ 找到微信窗口: HWND={hwnd}, 标题='{title}', 类名='{class_name}'")
                        try:
                            self.wechat_window = auto.ControlFromHandle(hwnd)
                            return True
                        except Exception as e:
                            if self.debug_mode:
                                print(f"  [调试] 从句柄获取元素失败: {e}")
                except Exception as e:
                    if self.debug_mode:
                        print(f"  [调试] 枚举窗口失败: {e}")
            
            print("❌ 未找到微信窗口")
            print("请确保:")
            print("  1. 微信已登录并运行")
            print("  2. 微信窗口未最小化到系统托盘")
            return False
            
        except Exception as e:
            print(f"查找微信窗口时出错: {e}")
            return False
    
    def _get_element_name(self, element):
        """安全获取元素名称"""
        try:
            if not element:
                return ""
            name = element.Name
            return name if name else ""
        except Exception:
            return ""
    
    def _get_element_class(self, element):
        """安全获取元素类名"""
        try:
            if not element:
                return ""
            class_name = element.ClassName
            return class_name if class_name else ""
        except Exception:
            return ""
    
    def _get_children(self, element):
        """安全获取子元素"""
        try:
            if not element:
                return []
            children = element.GetChildren()
            return children if children else []
        except Exception:
            return []
    
    def _get_parent(self, element):
        """安全获取父元素"""
        try:
            if not element:
                return None
            parent = element.GetParentControl()
            return parent
        except Exception:
            return None
    
    def print_element_tree(self, element, max_depth=4, indent=0):
        """打印元素树结构（用于调试）"""
        if not element:
            return
        
        try:
            name = self._get_element_name(element)
            class_name = self._get_element_class(element)
            
            prefix = "  " * indent
            info = []
            if class_name:
                info.append(f"类:{class_name}")
            if name:
                name_short = name[:60] + ('...' if len(name) > 60 else '')
                info.append(f"名:'{name_short}'")
            
            print(f"{prefix}├── {' | '.join(info)}")
            
            if indent < max_depth - 1:
                children = self._get_children(element)
                for child in children[:8]:
                    self.print_element_tree(child, max_depth, indent + 1)
                    
        except Exception:
            pass
    
    # =============================================================================
    # 问题2修复：增强文件元素识别逻辑
    # =============================================================================
    def _is_file_element(self, element):
        """
        检查元素是否为文件类型的消息元素
        改进：更严格的匹配逻辑，避免误识别
        """
        try:
            if not element:
                return False
            
            name = self._get_element_name(element)
            class_name = self._get_element_class(element)
            
            if self.debug_mode:
                print(f"  [调试] 检查是否为文件元素: 名称='{name[:50] if name else ''}', 类名='{class_name}'")
            
            # 如果没有名称，直接返回 False
            if not name:
                return False
            
            # 规则1：检查名称中是否包含文件关键词（高优先级）
            for keyword in FILE_KEYWORDS:
                if keyword in name:
                    if self.debug_mode:
                        print(f"  [调试] ✓ 匹配文件关键词: '{keyword}'")
                    return True
            
            # 规则2：检查类名是否包含文件相关特征
            if 'File' in class_name or 'file' in class_name:
                if self.debug_mode:
                    print(f"  [调试] ✓ 类名包含 File")
                return True
            
            # =============================================================================
            # 关键改进：更严格的文件扩展名匹配
            # 微信文件消息通常格式为: "文件名.ext (大小)" 或 "文件: 文件名.ext"
            # =============================================================================
            
            # 规则3：检查是否包含文件大小格式（如 "(2.5 MB)", "(100 KB)"）
            has_file_size = bool(re.search(FILE_SIZE_PATTERN, name))
            
            # 规则4：检查是否包含文件扩展名（更严格）
            # 匹配: 文件名.ext 或 文件名.ext (大小)
            # 不匹配: 普通文本中的 .txt 等
            file_ext_pattern = r'(^|\s|文件:|文件)([^\s\\/:*?"<>|\r\n]+\.[a-zA-Z0-9]{2,4})'
            ext_match = re.search(file_ext_pattern, name)
            
            if self.debug_mode:
                print(f"  [调试]   - 包含文件大小: {has_file_size}")
                print(f"  [调试]   - 包含文件扩展名: {bool(ext_match)}")
            
            # 组合条件：
            # A. 有文件大小 + 有扩展名 = 几乎肯定是文件消息
            # B. 有扩展名 + 类名是 TextBlock/TextEdit 等常见控件
            if has_file_size and ext_match:
                if self.debug_mode:
                    print(f"  [调试] ✓ 文件大小 + 扩展名匹配，判定为文件消息")
                return True
            
            # 额外检查：排除明显不是文件的情况
            # 如果名称包含 URL，排除
            if 'http://' in name or 'https://' in name:
                if self.debug_mode:
                    print(f"  [调试] ✗ 包含 URL，排除")
                return False
            
            # 如果只是纯文本中的扩展名（没有文件大小，也不是文件消息格式），排除
            # 例如："请查看 requirements.txt" 这样的文本消息
            if ext_match and not has_file_size:
                # 进一步检查：微信文件消息的扩展名前后通常有特定格式
                # 检查是否有 "文件" 前缀或其他特征
                if '文件' in name[:10] or class_name in ['TextBlock', 'Text']:
                    # 这可能是文件消息
                    if self.debug_mode:
                        print(f"  [调试] ✓ 扩展名匹配且格式合理，可能是文件消息")
                    return True
                else:
                    if self.debug_mode:
                        print(f"  [调试] ✗ 扩展名匹配但格式不合理，可能是普通文本")
                    return False
            
            return False
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 检查文件元素时出错: {e}")
            return False
    
    def _is_valid_path(self, path_str):
        """检查字符串是否为有效的文件路径"""
        if not path_str or not isinstance(path_str, str):
            return False
        
        path_str = path_str.strip()
        if not path_str:
            return False
        
        # 检查是否包含常见的路径分隔符
        if '\\' not in path_str and '/' not in path_str:
            return False
        
        # 检查是否包含驱动器号（Windows路径）
        if re.match(r'^[A-Za-z]:', path_str):
            return True
        
        # 检查是否包含 WeChat Files 等微信相关路径
        if 'WeChat Files' in path_str or '微信文件' in path_str:
            return True
        
        # 检查是否是网络路径
        if path_str.startswith('\\\\'):
            return True
        
        return False
    
    def _normalize_path(self, path_str):
        """规范化文件路径"""
        if not path_str:
            return ""
        
        path_str = path_str.strip()
        path_str = path_str.replace('/', '\\')
        path_str = re.sub(r'\s+', ' ', path_str).strip()
        path_str = path_str.strip('"\'')
        return path_str
    
    # =============================================================================
    # 关键改进：缓存微信文件目录，避免重复搜索
    # =============================================================================
    def _get_wechat_file_dir(self):
        """获取并缓存微信文件保存目录"""
        if self.wechat_file_dir:
            return self.wechat_file_dir
        
        self.wechat_file_dir = self._find_wechat_file_directory()
        return self.wechat_file_dir
    
    def _find_wechat_file_directory(self):
        """查找微信文件保存目录"""
        try:
            # 首先检查默认路径
            if os.path.exists(DEFAULT_WECHAT_FILE_PATH):
                # 查找用户文件夹
                for item in os.listdir(DEFAULT_WECHAT_FILE_PATH):
                    item_path = os.path.join(DEFAULT_WECHAT_FILE_PATH, item)
                    if os.path.isdir(item_path):
                        # 检查是否包含 FileStorage 或 Msg 目录
                        file_storage = os.path.join(item_path, 'FileStorage')
                        if os.path.exists(file_storage):
                            if self.debug_mode:
                                print(f"  [调试] 找到微信文件目录: {file_storage}")
                            return file_storage
                        msg_dir = os.path.join(item_path, 'Msg')
                        if os.path.exists(msg_dir):
                            if self.debug_mode:
                                print(f"  [调试] 找到微信文件目录: {msg_dir}")
                            return msg_dir
                
                return DEFAULT_WECHAT_FILE_PATH
            
            # 尝试查找其他可能的微信路径
            possible_paths = [
                os.path.join(os.environ.get('USERPROFILE', ''), 'WeChat Files'),
                os.path.join(os.environ.get('APPDATA', ''), 'Tencent', 'WeChat'),
                os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Tencent', 'WeChat'),
            ]
            
            for path in possible_paths:
                if os.path.exists(path):
                    if self.debug_mode:
                        print(f"  [调试] 找到微信文件目录: {path}")
                    return path
            
            if self.debug_mode:
                print("  [调试] 未找到微信文件目录")
            return None
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 查找微信文件目录时出错: {e}")
            return None
    
    # =============================================================================
    # 问题2修复：改进文件搜索逻辑
    # =============================================================================
    def _find_file_in_directory(self, directory, file_name):
        """
        在目录中查找文件（递归）
        改进：只搜索微信相关目录，避免搜索到项目目录下的 requirements.txt
        """
        try:
            if not directory or not os.path.exists(directory):
                return None
            
            if self.debug_mode:
                print(f"  [调试] 在微信目录中搜索文件: {file_name}")
                print(f"  [调试] 搜索根目录: {directory}")
            
            # =============================================================================
            # 关键改进：优先检查微信特有的文件接收目录
            # 这些目录下的文件才是微信接收的文件
            # =============================================================================
            wechat_specific_dirs = [
                'File',           # 接收的文件
                'FileStorage',    # 文件存储
                'Video',          # 视频文件
                'Image',          # 图片文件
                'Attachment',     # 附件
            ]
            
            for specific_dir in wechat_specific_dirs:
                specific_path = os.path.join(directory, specific_dir)
                if os.path.exists(specific_path):
                    if self.debug_mode:
                        print(f"  [调试] 检查微信特定目录: {specific_path}")
                    
                    # 在该目录下查找（限制深度）
                    for root, dirs, files in os.walk(specific_path):
                        # 限制深度，避免搜索过深
                        depth = root[len(specific_path):].count(os.sep)
                        if depth > 4:
                            continue
                        
                        if file_name in files:
                            full_path = os.path.join(root, file_name)
                            
                            # =============================================================================
                            # 额外验证：确保找到的文件确实在微信目录中
                            # 排除项目目录或其他非微信目录
                            # =============================================================================
                            if 'WeChat Files' in full_path or '微信文件' in full_path:
                                if self.debug_mode:
                                    print(f"  [调试] ✓ 在微信目录中找到文件: {full_path}")
                                return full_path
                            else:
                                if self.debug_mode:
                                    print(f"  [调试] ✗ 找到的文件不在微信目录中，跳过: {full_path}")
            
            # 如果在特定目录没找到，再检查整个微信目录（但有更严格的验证）
            if self.debug_mode:
                print(f"  [调试] 在特定目录未找到，检查整个微信目录...")
            
            for root, dirs, files in os.walk(directory):
                depth = root[len(directory):].count(os.sep)
                if depth > 5:
                    continue
                
                if file_name in files:
                    full_path = os.path.join(root, file_name)
                    
                    # 严格验证：必须在微信相关目录中
                    if 'WeChat Files' in full_path or '微信文件' in full_path:
                        if self.debug_mode:
                            print(f"  [调试] ✓ 在微信目录中找到文件: {full_path}")
                        return full_path
                    else:
                        if self.debug_mode:
                            print(f"  [调试] ✗ 找到的文件不在微信目录中，跳过: {full_path}")
            
            if self.debug_mode:
                print(f"  [调试] ✗ 未在微信目录中找到文件: {file_name}")
            return None
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 在目录中查找文件时出错: {e}")
            return None
    
    def _extract_file_name(self, name):
        """
        从元素名称中提取文件名
        改进：更严格的提取逻辑
        """
        if not name:
            return None
        
        # 微信文件消息的常见格式：
        # 1. "文件名.ext (大小)"  如: "document.pdf (2.5 MB)"
        # 2. "文件: 文件名.ext"   如: "文件: 照片.jpg"
        # 3. "文件名.ext"         如: "report.docx"
        
        # 首先尝试匹配带文件大小的格式（最可靠）
        pattern_with_size = r'([^\s\\/:*?"<>|\r\n]+\.[a-zA-Z0-9]{2,4})\s*\(\d+\.?\d*\s*[KMGT]?B\)'
        match = re.search(pattern_with_size, name)
        if match:
            file_name = match.group(1)
            if self.debug_mode:
                print(f"  [调试] 从带大小的格式提取文件名: {file_name}")
            return file_name
        
        # 然后尝试匹配 "文件: 文件名.ext" 格式
        pattern_with_prefix = r'文件[:：]\s*([^\s\\/:*?"<>|\r\n]+\.[a-zA-Z0-9]{2,4})'
        match = re.search(pattern_with_prefix, name)
        if match:
            file_name = match.group(1)
            if self.debug_mode:
                print(f"  [调试] 从带前缀的格式提取文件名: {file_name}")
            return file_name
        
        # 最后尝试匹配简单的 "文件名.ext" 格式
        # 但这种方式最不可靠，需要额外验证
        pattern_simple = r'([^\s\\/:*?"<>|\r\n]+\.[a-zA-Z0-9]{2,4})'
        matches = re.findall(pattern_simple, name)
        
        if matches:
            # 过滤掉明显不是微信文件的扩展名
            # 例如：.txt 可能是普通文本，.py 是代码文件
            excluded_exts = ['.txt', '.py', '.js', '.html', '.css', '.json', '.xml', '.md']
            
            for file_name in matches:
                ext = os.path.splitext(file_name)[1].lower()
                if ext not in excluded_exts:
                    if self.debug_mode:
                        print(f"  [调试] 从简单格式提取文件名: {file_name}")
                    return file_name
            
            # 如果所有扩展名都被排除了，返回 None
            if self.debug_mode:
                print(f"  [调试] 提取的文件扩展名被排除: {matches}")
            return None
        
        return None
    
    def _extract_path_from_file_element(self, element):
        """从文件元素中提取路径"""
        try:
            if not element:
                return None
            
            name = self._get_element_name(element)
            
            # 首先检查名称是否已经是完整路径
            if self._is_valid_path(name):
                path = self._normalize_path(name)
                if self.debug_mode:
                    print(f"  [调试] 名称已经是完整路径: {path}")
                return path
            
            # 从名称中提取文件名
            file_name = self._extract_file_name(name)
            if not file_name:
                if self.debug_mode:
                    print(f"  [调试] 无法从名称中提取文件名")
                return None
            
            if self.debug_mode:
                print(f"  [调试] 提取到文件名: {file_name}")
            
            # 获取微信文件目录
            wechat_path = self._get_wechat_file_dir()
            if not wechat_path:
                if self.debug_mode:
                    print(f"  [调试] 未找到微信文件目录")
                return None
            
            # 在微信目录中查找文件
            full_path = self._find_file_in_directory(wechat_path, file_name)
            if full_path:
                return full_path
            
            return None
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 从文件元素提取路径时出错: {e}")
            return None
    
    def _find_file_path_in_children(self, element):
        """在子元素中查找文件路径"""
        try:
            if not element:
                return None
            
            children = self._get_children(element)
            
            for child in children:
                child_name = self._get_element_name(child)
                
                # 检查子元素名称是否为有效路径
                if self._is_valid_path(child_name):
                    return self._normalize_path(child_name)
                
                # 递归检查更深层的子元素（限制深度）
                nested_path = self._find_file_path_in_children(child)
                if nested_path:
                    return nested_path
                
                # 检查是否为文件元素
                if self._is_file_element(child):
                    path = self._extract_path_from_file_element(child)
                    if path:
                        return path
            
            return None
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 在子元素中查找路径时出错: {e}")
            return None
    
    def _find_file_path_in_context(self, element):
        """在元素的上下文（父元素、兄弟元素）中查找文件路径"""
        try:
            if not element:
                return None
            
            # 检查父元素（最多检查3层）
            max_levels = 3
            current = element
            
            for level in range(max_levels):
                parent = self._get_parent(current)
                if not parent:
                    break
                
                parent_name = self._get_element_name(parent)
                
                if self._is_valid_path(parent_name):
                    return self._normalize_path(parent_name)
                
                # 检查父元素的其他子元素
                siblings = self._get_children(parent)
                for sibling in siblings:
                    if sibling == current:
                        continue
                    
                    sibling_name = self._get_element_name(sibling)
                    
                    if self._is_valid_path(sibling_name):
                        return self._normalize_path(sibling_name)
                    
                    # 检查是否为文件元素
                    if self._is_file_element(sibling):
                        path = self._extract_path_from_file_element(sibling)
                        if path:
                            return path
                
                current = parent
            
            return None
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 在上下文中查找路径时出错: {e}")
            return None
    
    def _construct_path_from_features(self, element):
        """根据元素特征构造文件路径"""
        try:
            if not element:
                return None
            
            name = self._get_element_name(element)
            
            # 首先检查是否为文件元素
            if not self._is_file_element(element):
                return None
            
            # 提取文件名
            file_name = self._extract_file_name(name)
            if not file_name:
                return None
            
            if self.debug_mode:
                print(f"  [调试] 特征构造: 提取到文件名 {file_name}")
            
            # 获取微信文件目录
            wechat_path = self._get_wechat_file_dir()
            if not wechat_path:
                return None
            
            # 在微信目录中查找文件
            full_path = self._find_file_in_directory(wechat_path, file_name)
            if full_path:
                return full_path
            
            return None
            
        except Exception as e:
            # 不要打印 COM 错误
            error_str = str(e)
            if '-2147220991' not in error_str and '0x80040201' not in error_str:
                if self.debug_mode:
                    print(f"  [调试] 根据特征构造路径时出错: {e}")
            return None
    
    def get_file_path_from_element(self, element):
        """从UI元素中提取文件路径"""
        try:
            if not element:
                return None
            
            if self.debug_mode:
                print("\n" + "="*60)
                print("[调试] 开始分析元素...")
                print("="*60)
                self.print_element_tree(element, max_depth=3)
            
            # 获取元素的基本属性
            name = self._get_element_name(element)
            class_name = self._get_element_class(element)
            
            if self.debug_mode:
                print(f"\n[调试] 当前元素信息:")
                print(f"  名称: '{name[:80] if name else ''}'")
                print(f"  类名: '{class_name}'")
            
            # 方法1: 直接从名称中提取路径
            if self._is_valid_path(name):
                path = self._normalize_path(name)
                if self.debug_mode:
                    print(f"[调试] 方法1成功: 从名称提取路径: {path}")
                return path
            
            # 方法2: 检查子元素
            file_path = self._find_file_path_in_children(element)
            if file_path:
                if self.debug_mode:
                    print(f"[调试] 方法2成功: 从子元素提取路径: {file_path}")
                return file_path
            
            # 方法3: 检查父元素和兄弟元素
            file_path = self._find_file_path_in_context(element)
            if file_path:
                if self.debug_mode:
                    print(f"[调试] 方法3成功: 从上下文提取路径: {file_path}")
                return file_path
            
            # 方法4: 根据文件名特征构造路径
            file_path = self._construct_path_from_features(element)
            if file_path:
                if self.debug_mode:
                    print(f"[调试] 方法4成功: 根据特征构造路径: {file_path}")
                return file_path
            
            if self.debug_mode:
                print("[调试] 所有方法均未能提取路径")
            
            return None
            
        except Exception as e:
            print(f"提取文件路径时出错: {e}")
            return None
    
    def is_wechat_element(self, element):
        """检查元素是否属于微信窗口"""
        try:
            if not element:
                return False
            
            # 向上遍历直到找到顶层窗口
            current = element
            max_levels = 20
            level = 0
            
            while current and level < max_levels:
                try:
                    class_name = self._get_element_class(current)
                    name = self._get_element_name(current)
                    
                    if WECHAT_CLASS_NAME in class_name:
                        if self.debug_mode:
                            print(f"  [调试] 找到微信窗口 (类名匹配): {class_name}")
                        return True
                    
                    if WECHAT_TITLE in name and 'Window' in str(type(current)):
                        if self.debug_mode:
                            print(f"  [调试] 找到微信窗口 (标题匹配): {name}")
                        return True
                    
                    # 获取父元素
                    current = self._get_parent(current)
                    level += 1
                    
                except Exception:
                    break
            
            return False
            
        except Exception:
            return False
    
    def get_element_at_mouse(self):
        """获取鼠标位置的UI元素"""
        try:
            # 获取鼠标位置
            x, y = auto.GetCursorPos()
            
            if self.debug_mode:
                print(f"\n[调试] 鼠标位置: ({x}, {y})")
            
            # 从鼠标位置获取元素
            try:
                element = auto.ControlFromPoint(x, y)
                return element
            except Exception as e:
                if self.debug_mode:
                    print(f"  [调试] 获取鼠标位置元素时出错: {e}")
                return None
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 获取鼠标位置元素时出错: {e}")
            return None
    
    def check_and_copy_file_path(self, element):
        """检查元素是否为文件消息，如果是则复制路径到剪贴板"""
        try:
            if not element:
                return False
            
            # 检查是否属于微信窗口
            if not self.is_wechat_element(element):
                if self.debug_mode:
                    print("[调试] 元素不属于微信窗口，跳过")
                return False
            
            # 获取文件路径
            file_path = self.get_file_path_from_element(element)
            
            if file_path:
                # 检查是否与上次复制的相同（避免重复复制）
                if file_path == self.last_clicked_file:
                    if self.debug_mode:
                        print(f"[调试] 路径与上次相同，跳过: {file_path}")
                    return False
                
                # 复制到剪贴板
                pyperclip.copy(file_path)
                
                # 记录上次复制的文件
                self.last_clicked_file = file_path
                
                # 打印提示
                print(f"\n{'='*60}")
                print(f"✅ 已复制文件路径到剪贴板:")
                print(f"   {file_path}")
                print(f"{'='*60}\n")
                
                return True
            
            return False
            
        except Exception as e:
            print(f"检查和复制文件路径时出错: {e}")
            return False
    
    # =============================================================================
    # 问题1修复：使用 win32api 检测鼠标点击
    # =============================================================================
    def check_left_mouse_click(self):
        """检测鼠标左键点击（使用 win32api）"""
        if not WIN32_AVAILABLE:
            return False
        
        try:
            # 检测左键是否被按下
            state = win32api.GetAsyncKeyState(win32con.VK_LBUTTON)
            # 最高位为1表示按下
            return state < 0
        except Exception:
            return False
    
    def check_right_mouse_click(self):
        """检测鼠标右键点击"""
        if not WIN32_AVAILABLE:
            return False
        
        try:
            state = win32api.GetAsyncKeyState(win32con.VK_RBUTTON)
            return state < 0
        except Exception:
            return False
    
    def monitor_mouse_clicks(self):
        """监控鼠标点击事件"""
        print("\n开始监控鼠标点击事件...")
        print("="*60)
        print("使用方法:")
        print("  1. 在微信聊天窗口中找到文件消息")
        print("  2. 用鼠标左键或右键点击文件消息")
        print("  3. 文件路径将自动复制到剪贴板")
        print("="*60)
        print(f"调试模式: {'开启' if self.debug_mode else '关闭'}")
        print(f"pywin32: {'✓ 已加载' if WIN32_AVAILABLE else '✗ 未加载'}")
        print("按 Ctrl+C 停止监控\n")
        
        self.monitoring = True
        
        # 记录上次的点击状态
        left_was_pressed = False
        right_was_pressed = False
        
        # 防抖延迟
        DEBOUNCE_DELAY = 0.2
        last_left_click_time = 0
        last_right_click_time = 0
        
        try:
            while self.monitoring:
                current_time = time.time()
                
                # =============================================================================
                # 问题1修复：使用 win32api 检测真实点击
                # =============================================================================
                if WIN32_AVAILABLE:
                    # 检测左键点击
                    left_pressed = self.check_left_mouse_click()
                    
                    # 检测上升沿（按下后释放）
                    if left_was_pressed and not left_pressed:
                        if current_time - last_left_click_time > DEBOUNCE_DELAY:
                            if self.debug_mode:
                                print("\n[调试] 检测到鼠标左键点击")
                            
                            # 获取鼠标位置的元素
                            element = self.get_element_at_mouse()
                            
                            if element:
                                self.check_and_copy_file_path(element)
                            
                            last_left_click_time = current_time
                    
                    left_was_pressed = left_pressed
                    
                    # 检测右键点击
                    right_pressed = self.check_right_mouse_click()
                    
                    if right_was_pressed and not right_pressed:
                        if current_time - last_right_click_time > DEBOUNCE_DELAY:
                            if self.debug_mode:
                                print("\n[调试] 检测到鼠标右键点击")
                            
                            # 获取鼠标位置的元素
                            element = self.get_element_at_mouse()
                            
                            if element:
                                self.check_and_copy_file_path(element)
                            
                            last_right_click_time = current_time
                    
                    right_was_pressed = right_pressed
                
                else:
                    # 如果没有 win32api，使用简单的位置稳定检测
                    current_x, current_y = auto.GetCursorPos()
                    
                    if not hasattr(self, '_last_pos'):
                        self._last_pos = (current_x, current_y)
                        self._last_pos_time = current_time
                    
                    if (current_x, current_y) == self._last_pos:
                        if current_time - self._last_pos_time > 0.5:
                            if current_time - (getattr(self, '_last_simple_check', 0)) > 1.0:
                                if self.debug_mode:
                                    print("\n[调试] 检测到鼠标位置稳定（简单模式）")
                                
                                element = self.get_element_at_mouse()
                                if element:
                                    self.check_and_copy_file_path(element)
                                
                                self._last_simple_check = current_time
                    else:
                        self._last_pos = (current_x, current_y)
                        self._last_pos_time = current_time
                
                # 短暂休眠，减少CPU占用
                time.sleep(0.02)
                
        except KeyboardInterrupt:
            print("\n\n监控已停止")
        finally:
            self.monitoring = False
    
    def start(self):
        """启动监控"""
        print("="*60)
        print("微信文件消息路径提取工具 v2.1")
        print("="*60)
        print()
        
        # =============================================================================
        # 问题1修复：正确显示 pywin32 状态
        # =============================================================================
        print(f"[系统] pywin32 状态: {'✓ 已加载' if WIN32_AVAILABLE else '✗ 未加载'}")
        if not WIN32_AVAILABLE:
            print("⚠️  警告: 未安装 pywin32，点击检测可能不够准确")
            print("   建议安装: pip install pywin32")
        print()
        
        # 查找微信窗口
        if not self.find_wechat_window():
            print("\n继续运行于通用模式（将检测所有窗口的点击）")
            print()
        
        # 预加载微信文件目录
        print("正在初始化微信文件目录...")
        wechat_dir = self._get_wechat_file_dir()
        if wechat_dir:
            print(f"✓ 微信文件目录: {wechat_dir}")
        else:
            print("⚠️  未找到微信文件目录，可能无法自动查找文件路径")
        print()
        
        # 启动监控
        try:
            self.monitor_mouse_clicks()
        except Exception as e:
            print(f"监控过程中出错: {e}")
            import traceback
            traceback.print_exc()
            return False
        
        return True


def main():
    """主函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description='微信文件消息路径提取工具')
    parser.add_argument('--no-debug', action='store_true', help='关闭调试模式')
    parser.add_argument('--debug', action='store_true', help='开启调试模式（默认）')
    
    args = parser.parse_args()
    
    # 设置调试模式
    global DEBUG_MODE
    if args.no_debug:
        DEBUG_MODE = False
    elif args.debug:
        DEBUG_MODE = True
    
    monitor = WeChatFileMonitor()
    monitor.debug_mode = DEBUG_MODE
    
    try:
        monitor.start()
    except KeyboardInterrupt:
        print("\n\n程序已退出")
    except Exception as e:
        print(f"程序运行出错: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
