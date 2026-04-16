# -*- coding: utf-8 -*-
"""
PC版微信文件消息路径提取工具 - 改进版
功能：监控微信聊天窗口，当用户选中或右键点击文件类型消息时，
      自动获取文件的绝对保存路径并复制到剪贴板
      
改进：
1. 修复 COM 错误 (-2147220991) - 添加重试机制和更健壮的错误处理
2. 改进点击检测 - 使用更可靠的检测逻辑
3. 添加调试模式 - 可查看元素结构
4. 优化微信元素识别
"""

import uiautomation as auto
import pyperclip
import time
import re
import os
import sys
import threading
from ctypes import windll, wintypes, byref

try:
    import win32api
    import win32gui
    import win32con
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# 微信窗口的类名和标题
WECHAT_CLASS_NAME = 'WeChatMainWndForPC'
WECHAT_TITLE = '微信'

# 文件消息的特征关键词
FILE_KEYWORDS = ['文件', '文件:', '已接收文件', '接收文件', '另存为', '打开文件夹', '打开']

# 微信默认文件保存路径（用于路径拼接）
DEFAULT_WECHAT_FILE_PATH = os.path.join(
    os.environ.get('USERPROFILE', ''),
    'Documents', 'WeChat Files'
)

# 调试模式
DEBUG_MODE = True


class COMError(Exception):
    """COM 错误异常类"""
    pass


def safe_com_call(func, *args, max_retries=3, **kwargs):
    """
    安全的 COM 调用包装器
    自动重试失败的 COM 调用
    """
    last_error = None
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            error_str = str(e)
            # 检测 COM 错误
            if '-2147220991' in error_str or '0x80040201' in error_str:
                last_error = e
                if attempt < max_retries - 1:
                    time.sleep(0.05 * (attempt + 1))
                    continue
            # 其他错误直接抛出
            raise
    raise last_error if last_error else Exception("COM 调用失败")


def get_element_safe(element, attr_name, default=""):
    """
    安全获取元素属性
    处理可能的 COM 错误
    """
    try:
        if not element:
            return default
        # 使用安全的 COM 调用
        value = safe_com_call(lambda: getattr(element, attr_name, default))
        return value if value else default
    except Exception as e:
        if DEBUG_MODE:
            print(f"  [调试] 获取属性 {attr_name} 时出错: {e}")
        return default


def get_children_safe(element):
    """
    安全获取子元素
    """
    try:
        if not element:
            return []
        children = safe_com_call(lambda: element.GetChildren())
        return children if children else []
    except Exception as e:
        if DEBUG_MODE:
            print(f"  [调试] 获取子元素时出错: {e}")
        return []


def get_parent_safe(element):
    """
    安全获取父元素
    """
    try:
        if not element:
            return None
        parent = safe_com_call(lambda: element.GetParentControl())
        return parent
    except Exception as e:
        if DEBUG_MODE:
            print(f"  [调试] 获取父元素时出错: {e}")
        return None


class WeChatFileMonitor:
    """微信文件消息监控类"""
    
    def __init__(self):
        self.wechat_window = None
        self.last_clicked_file = None
        self.monitoring = False
        self.debug_mode = DEBUG_MODE
        
    def find_wechat_window(self):
        """查找微信主窗口"""
        try:
            print("正在查找微信窗口...")
            
            # 方法1: 通过类名查找
            try:
                window = auto.WindowControl(ClassName=WECHAT_CLASS_NAME)
                if safe_com_call(lambda: window.Exists(2)):
                    self.wechat_window = window
                    print(f"✓ 通过类名找到微信窗口: {WECHAT_CLASS_NAME}")
                    return True
            except Exception as e:
                if self.debug_mode:
                    print(f"  [调试] 类名查找失败: {e}")
            
            # 方法2: 通过标题查找
            try:
                window = auto.WindowControl(Name=WECHAT_TITLE)
                if safe_com_call(lambda: window.Exists(2)):
                    self.wechat_window = window
                    print(f"✓ 通过标题找到微信窗口: {WECHAT_TITLE}")
                    return True
            except Exception as e:
                if self.debug_mode:
                    print(f"  [调试] 标题查找失败: {e}")
            
            # 方法3: 枚举所有顶层窗口
            print("尝试枚举所有窗口查找微信...")
            try:
                def callback(hwnd, windows):
                    if win32gui.IsWindowVisible(hwnd):
                        title = win32gui.GetWindowText(hwnd)
                        class_name = win32gui.GetClassName(hwnd)
                        if WECHAT_TITLE in title or WECHAT_CLASS_NAME in class_name:
                            windows.append((hwnd, title, class_name))
                    return True
                
                if WIN32_AVAILABLE:
                    windows = []
                    win32gui.EnumWindows(callback, windows)
                    if windows:
                        hwnd, title, class_name = windows[0]
                        print(f"✓ 找到微信窗口: HWND={hwnd}, 标题='{title}', 类名='{class_name}'")
                        # 通过句柄获取 uiautomation 元素
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
            print("  3. 微信窗口没有被其他窗口完全遮挡")
            return False
            
        except Exception as e:
            print(f"查找微信窗口时出错: {e}")
            return False
    
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
                    class_name = get_element_safe(current, 'ClassName', '')
                    name = get_element_safe(current, 'Name', '')
                    
                    if WECHAT_CLASS_NAME in class_name:
                        if self.debug_mode:
                            print(f"  [调试] 找到微信窗口 (类名匹配): {class_name}")
                        return True
                    
                    if WECHAT_TITLE in name and 'Window' in str(type(current)):
                        if self.debug_mode:
                            print(f"  [调试] 找到微信窗口 (标题匹配): {name}")
                        return True
                    
                    # 获取父元素
                    current = get_parent_safe(current)
                    level += 1
                    
                except Exception as e:
                    if self.debug_mode:
                        print(f"  [调试] 遍历父元素时出错: {e}")
                    break
            
            return False
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 检查微信元素时出错: {e}")
            return False
    
    def print_element_tree(self, element, max_depth=5, indent=0):
        """打印元素树结构（用于调试）"""
        if not element:
            return
        
        try:
            name = get_element_safe(element, 'Name', '')
            class_name = get_element_safe(element, 'ClassName', '')
            automation_id = get_element_safe(element, 'AutomationId', '')
            
            prefix = "  " * indent
            info = []
            if class_name:
                info.append(f"类:{class_name}")
            if name:
                info.append(f"名:'{name[:50]}{'...' if len(name) > 50 else ''}'")
            if automation_id:
                info.append(f"ID:{automation_id}")
            
            print(f"{prefix}├── {' | '.join(info)}")
            
            if indent < max_depth - 1:
                children = get_children_safe(element)
                for child in children[:10]:  # 最多显示10个子元素
                    self.print_element_tree(child, max_depth, indent + 1)
                    
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 打印元素树时出错: {e}")
    
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
            name = get_element_safe(element, 'Name', '')
            class_name = get_element_safe(element, 'ClassName', '')
            automation_id = get_element_safe(element, 'AutomationId', '')
            
            if self.debug_mode:
                print(f"\n[调试] 当前元素信息:")
                print(f"  名称: '{name}'")
                print(f"  类名: '{class_name}'")
                print(f"  AutomationId: '{automation_id}'")
            
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
        # 替换正斜杠为反斜杠
        path_str = path_str.replace('/', '\\')
        # 移除多余的空格和特殊字符
        path_str = re.sub(r'\s+', ' ', path_str).strip()
        # 移除可能的路径前后的引号
        path_str = path_str.strip('"\'')
        return path_str
    
    def _find_file_path_in_children(self, element):
        """在子元素中查找文件路径"""
        try:
            if not element:
                return None
            
            children = get_children_safe(element)
            
            for child in children:
                child_name = get_element_safe(child, 'Name', '')
                
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
                parent = get_parent_safe(current)
                if not parent:
                    break
                
                parent_name = get_element_safe(parent, 'Name', '')
                
                if self._is_valid_path(parent_name):
                    return self._normalize_path(parent_name)
                
                # 检查父元素的其他子元素
                siblings = get_children_safe(parent)
                for sibling in siblings:
                    if sibling == current:
                        continue
                    
                    sibling_name = get_element_safe(sibling, 'Name', '')
                    
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
    
    def _is_file_element(self, element):
        """检查元素是否为文件类型的消息元素"""
        try:
            if not element:
                return False
            
            name = get_element_safe(element, 'Name', '')
            class_name = get_element_safe(element, 'ClassName', '')
            
            # 检查名称中是否包含文件关键词
            for keyword in FILE_KEYWORDS:
                if keyword in name:
                    if self.debug_mode:
                        print(f"  [调试] 文件元素匹配: 关键词 '{keyword}' 在名称中")
                    return True
            
            # 检查类名是否包含文件相关特征
            if 'File' in class_name or 'file' in class_name:
                if self.debug_mode:
                    print(f"  [调试] 文件元素匹配: 类名包含 File")
                return True
            
            # 检查是否包含文件扩展名（更严格的匹配）
            # 匹配: 文件名.ext (大小) 或 文件名.ext
            file_pattern = r'([^\\/:*?"<>|\r\n]+\.[a-zA-Z0-9]{2,4})\s*(\(\d+[KMGT]?B\))?'
            if re.search(file_pattern, name):
                # 排除纯文本URL中的扩展名
                if 'http://' not in name and 'https://' not in name:
                    if self.debug_mode:
                        print(f"  [调试] 文件元素匹配: 名称包含文件扩展名")
                    return True
            
            return False
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 检查文件元素时出错: {e}")
            return False
    
    def _extract_path_from_file_element(self, element):
        """从文件元素中提取路径"""
        try:
            if not element:
                return None
            
            name = get_element_safe(element, 'Name', '')
            
            # 尝试从名称中提取文件名
            # 微信文件消息的格式可能是："文件名.ext (大小)" 或 "文件: 文件名.ext"
            
            # 提取文件名（包含扩展名）
            file_name_match = re.search(r'([^\\/:*?"<>|\r\n]+\.[a-zA-Z0-9]{2,4})', name)
            if file_name_match:
                file_name = file_name_match.group(1)
                
                if self.debug_mode:
                    print(f"  [调试] 提取到文件名: {file_name}")
                
                # 尝试查找微信文件目录
                wechat_path = self._find_wechat_file_directory()
                if wechat_path:
                    # 尝试在微信目录中查找文件
                    full_path = self._find_file_in_directory(wechat_path, file_name)
                    if full_path:
                        return full_path
            
            # 尝试直接从名称中提取完整路径
            if self._is_valid_path(name):
                return self._normalize_path(name)
            
            return None
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 从文件元素提取路径时出错: {e}")
            return None
    
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
    
    def _find_file_in_directory(self, directory, file_name):
        """在目录中查找文件（递归）"""
        try:
            if not directory or not os.path.exists(directory):
                return None
            
            # 优先检查常见的文件接收目录
            common_dirs = ['File', 'FileStorage', 'Msg', 'Attachment', 'Files', 'Video', 'Image']
            
            for common_dir in common_dirs:
                common_path = os.path.join(directory, common_dir)
                if os.path.exists(common_path):
                    # 在该目录下查找（限制深度）
                    for root, dirs, files in os.walk(common_path):
                        # 限制深度，避免搜索过深
                        depth = root[len(common_path):].count(os.sep)
                        if depth > 3:
                            continue
                        
                        if file_name in files:
                            full_path = os.path.join(root, file_name)
                            if self.debug_mode:
                                print(f"  [调试] 在目录中找到文件: {full_path}")
                            return full_path
            
            # 如果没找到，遍历整个目录（限制深度）
            for root, dirs, files in os.walk(directory):
                depth = root[len(directory):].count(os.sep)
                if depth > 5:
                    continue
                
                if file_name in files:
                    full_path = os.path.join(root, file_name)
                    if self.debug_mode:
                        print(f"  [调试] 在目录中找到文件: {full_path}")
                    return full_path
            
            return None
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 在目录中查找文件时出错: {e}")
            return None
    
    def _construct_path_from_features(self, element):
        """根据元素特征构造文件路径"""
        try:
            if not element:
                return None
            
            name = get_element_safe(element, 'Name', '')
            
            # 提取文件名
            file_name_match = re.search(r'([^\\/:*?"<>|\r\n]+\.[a-zA-Z0-9]{2,4})', name)
            if not file_name_match:
                return None
            
            file_name = file_name_match.group(1)
            
            if self.debug_mode:
                print(f"  [调试] 特征构造: 提取到文件名 {file_name}")
            
            # 查找微信文件目录
            wechat_path = self._find_wechat_file_directory()
            if not wechat_path:
                return None
            
            # 在微信目录中查找文件
            full_path = self._find_file_in_directory(wechat_path, file_name)
            if full_path:
                return full_path
            
            return None
            
        except Exception as e:
            # 不要打印 COM 错误，这是正常的
            error_str = str(e)
            if '-2147220991' not in error_str and '0x80040201' not in error_str:
                if self.debug_mode:
                    print(f"  [调试] 根据特征构造路径时出错: {e}")
            return None
    
    def get_element_at_mouse(self):
        """获取鼠标位置的UI元素"""
        try:
            # 获取鼠标位置
            x, y = auto.GetCursorPos()
            
            if self.debug_mode:
                print(f"\n[调试] 鼠标位置: ({x}, {y})")
            
            # 从鼠标位置获取元素
            try:
                element = safe_com_call(lambda: auto.ControlFromPoint(x, y))
                return element
            except Exception as e:
                if self.debug_mode:
                    print(f"  [调试] 获取鼠标位置元素时 COM 出错: {e}")
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
    
    def check_left_mouse_click(self):
        """检测鼠标左键点击（使用 win32api）"""
        if not WIN32_AVAILABLE:
            return False
        
        try:
            # 检测左键是否被按下
            state = win32api.GetAsyncKeyState(win32con.VK_LBUTTON)
            # 最高位为1表示按下
            return state < 0
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 检测鼠标点击时出错: {e}")
            return False
    
    def check_right_mouse_click(self):
        """检测鼠标右键点击"""
        if not WIN32_AVAILABLE:
            return False
        
        try:
            state = win32api.GetAsyncKeyState(win32con.VK_RBUTTON)
            return state < 0
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 检测右键点击时出错: {e}")
            return False
    
    def monitor_mouse_clicks(self):
        """监控鼠标点击事件（改进版）"""
        print("\n开始监控鼠标点击事件...")
        print("="*60)
        print("使用方法:")
        print("  1. 在微信聊天窗口中找到文件消息")
        print("  2. 用鼠标左键或右键点击文件消息")
        print("  3. 文件路径将自动复制到剪贴板")
        print("="*60)
        print(f"调试模式: {'开启' if self.debug_mode else '关闭'}")
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
                
                # 检测左键点击
                if WIN32_AVAILABLE:
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
                    
                    # 简单的稳定性检测
                    if not hasattr(self, '_last_pos'):
                        self._last_pos = (current_x, current_y)
                        self._last_pos_time = current_time
                    
                    if (current_x, current_y) == self._last_pos:
                        # 位置稳定超过一定时间
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
        print("微信文件消息路径提取工具 v2.0")
        print("="*60)
        print()
        
        # 检查依赖
        if not WIN32_AVAILABLE:
            print("⚠️  警告: 未安装 pywin32，点击检测可能不够准确")
            print("   建议安装: pip install pywin32")
            print()
        
        # 查找微信窗口
        if not self.find_wechat_window():
            print("\n继续运行于通用模式（将检测所有窗口的点击）")
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
    
    def stop(self):
        """停止监控"""
        self.monitoring = False
        print("监控已停止")


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
