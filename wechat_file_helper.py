# -*- coding: utf-8 -*-
"""
PC版微信文件消息路径提取工具 v2.2
修复：
1. pywin32 导入诊断 - 详细显示环境信息
2. 微信元素获取 - 使用窗口范围检测而不是直接从鼠标位置获取
"""

import uiautomation as auto
import pyperclip
import time
import re
import os
import sys

# =============================================================================
# 问题1修复：详细诊断 pywin32 导入问题
# =============================================================================
WIN32_AVAILABLE = False
win32api = None
win32con = None
win32gui = None

print("="*60)
print("系统诊断信息")
print("="*60)
print(f"Python 版本: {sys.version}")
print(f"Python 路径: {sys.executable}")
print(f"当前工作目录: {os.getcwd()}")
print()
print("正在检查 pywin32...")

# 尝试多种方式导入 pywin32
try:
    # 方式1: 直接导入
    import win32api
    import win32con
    import win32gui
    WIN32_AVAILABLE = True
    print(f"✓ pywin32 导入成功 (方式1: 直接导入)")
    print(f"  win32api 路径: {win32api.__file__}")
except ImportError as e1:
    print(f"✗ 方式1失败: {e1}")
    
    try:
        # 方式2: 从 win32 包导入
        from win32 import win32api
        from win32 import win32con
        from win32 import win32gui
        WIN32_AVAILABLE = True
        print(f"✓ pywin32 导入成功 (方式2: 从 win32 包导入)")
    except ImportError as e2:
        print(f"✗ 方式2失败: {e2}")
        
        try:
            # 方式3: 从 pywin32 包导入
            from pywin32 import win32api
            from pywin32 import win32con
            from pywin32 import win32gui
            WIN32_AVAILABLE = True
            print(f"✓ pywin32 导入成功 (方式3: 从 pywin32 包导入)")
        except ImportError as e3:
            print(f"✗ 方式3失败: {e3}")
            
            # 检查 sys.path
            print()
            print("Python 搜索路径:")
            for i, p in enumerate(sys.path):
                print(f"  {i+1}. {p}")
            
            # 检查是否有 win32 相关的包
            print()
            print("检查已安装的包...")
            try:
                import pkg_resources
                installed_packages = [pkg.key for pkg in pkg_resources.working_set]
                win32_packages = [p for p in installed_packages if 'win32' in p.lower() or 'pywin' in p.lower()]
                if win32_packages:
                    print(f"找到 win32 相关包: {win32_packages}")
                    print("提示: 可能需要运行 'python Scripts/pywin32_postinstall.py -install'")
                else:
                    print("未找到 win32 相关包")
                    print("提示: 请运行 'pip install pywin32' 安装")
            except:
                pass

if WIN32_AVAILABLE:
    print(f"✓ pywin32 状态: 已加载")
else:
    print(f"✗ pywin32 状态: 未加载")
print()

# 微信窗口的类名和标题
WECHAT_CLASS_NAME = 'WeChatMainWndForPC'
WECHAT_TITLE = '微信'

# 文件消息的特征关键词
FILE_KEYWORDS = ['文件', '文件:', '已接收文件', '接收文件', '另存为', '打开文件夹', '在文件夹中显示', '打开']

# 文件大小模式
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
        self.wechat_window_handle = None
        self.wechat_window_rect = None
        self.last_clicked_file = None
        self.monitoring = False
        self.debug_mode = DEBUG_MODE
        self.wechat_file_dir = None
        
    # =============================================================================
    # 问题2修复：使用窗口范围检测而不是直接从鼠标位置获取
    # =============================================================================
    def get_window_rect(self, hwnd):
        """获取窗口矩形（屏幕坐标）"""
        try:
            if WIN32_AVAILABLE and win32gui:
                rect = win32gui.GetWindowRect(hwnd)
                return rect  # (left, top, right, bottom)
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 获取窗口矩形失败: {e}")
        return None
    
    def is_point_in_window(self, x, y, rect):
        """检查点是否在窗口矩形内"""
        if not rect:
            return False
        left, top, right, bottom = rect
        return left <= x <= right and top <= y <= bottom
    
    def screen_to_client(self, hwnd, x, y):
        """屏幕坐标转换为客户端坐标"""
        try:
            if WIN32_AVAILABLE and win32gui:
                point = (x, y)
                client_point = win32gui.ScreenToClient(hwnd, point)
                return client_point
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 坐标转换失败: {e}")
        return None
    
    def find_wechat_window(self):
        """查找微信主窗口（改进版）"""
        try:
            print("正在查找微信窗口...")
            
            # 方法1: 使用 uiautomation 通过类名查找
            try:
                window = auto.WindowControl(ClassName=WECHAT_CLASS_NAME)
                if window.Exists(2):
                    self.wechat_window = window
                    # 尝试获取窗口句柄
                    try:
                        self.wechat_window_handle = window.NativeWindowHandle
                        if self.wechat_window_handle:
                            self.wechat_window_rect = self.get_window_rect(self.wechat_window_handle)
                            print(f"✓ 通过 uiautomation 找到微信窗口")
                            print(f"  句柄: {self.wechat_window_handle}")
                            if self.wechat_window_rect:
                                print(f"  位置: {self.wechat_window_rect}")
                            return True
                    except Exception as e:
                        if self.debug_mode:
                            print(f"  [调试] 获取窗口句柄失败: {e}")
                        print(f"✓ 通过 uiautomation 找到微信窗口（无法获取句柄）")
                        return True
            except Exception as e:
                if self.debug_mode:
                    print(f"  [调试] uiautomation 类名查找失败: {e}")
            
            # 方法2: 使用 win32gui 枚举窗口（更可靠）
            if WIN32_AVAILABLE and win32gui:
                print("尝试使用 win32gui 查找微信窗口...")
                try:
                    def callback(hwnd, windows):
                        if win32gui.IsWindowVisible(hwnd):
                            title = win32gui.GetWindowText(hwnd)
                            class_name = win32gui.GetClassName(hwnd)
                            if WECHAT_CLASS_NAME in class_name:
                                windows.append((hwnd, title, class_name))
                            elif WECHAT_TITLE in title and len(title) < 20:
                                windows.append((hwnd, title, class_name))
                        return True
                    
                    windows = []
                    win32gui.EnumWindows(callback, windows)
                    
                    if windows:
                        # 优先选择类名匹配的
                        for hwnd, title, class_name in windows:
                            if WECHAT_CLASS_NAME in class_name:
                                self.wechat_window_handle = hwnd
                                self.wechat_window_rect = self.get_window_rect(hwnd)
                                # 尝试获取 uiautomation 元素
                                try:
                                    self.wechat_window = auto.ControlFromHandle(hwnd)
                                    print(f"✓ 通过 win32gui 找到微信窗口（类名匹配）")
                                except:
                                    print(f"✓ 通过 win32gui 找到微信窗口（类名匹配，uiautomation 不可用）")
                                print(f"  句柄: {hwnd}")
                                print(f"  标题: '{title}'")
                                print(f"  类名: '{class_name}'")
                                if self.wechat_window_rect:
                                    print(f"  位置: {self.wechat_window_rect}")
                                return True
                        
                        # 如果没有类名匹配的，使用标题匹配的
                        hwnd, title, class_name = windows[0]
                        self.wechat_window_handle = hwnd
                        self.wechat_window_rect = self.get_window_rect(hwnd)
                        try:
                            self.wechat_window = auto.ControlFromHandle(hwnd)
                            print(f"✓ 通过 win32gui 找到微信窗口（标题匹配）")
                        except:
                            print(f"✓ 通过 win32gui 找到微信窗口（标题匹配，uiautomation 不可用）")
                        print(f"  句柄: {hwnd}")
                        print(f"  标题: '{title}'")
                        print(f"  类名: '{class_name}'")
                        return True
                        
                except Exception as e:
                    if self.debug_mode:
                        print(f"  [调试] win32gui 查找失败: {e}")
            
            print("❌ 未找到微信窗口")
            print("请确保:")
            print("  1. 微信已登录并运行")
            print("  2. 微信窗口未最小化到系统托盘")
            return False
            
        except Exception as e:
            print(f"查找微信窗口时出错: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def update_window_rect(self):
        """更新微信窗口位置（窗口可能移动）"""
        if self.wechat_window_handle:
            self.wechat_window_rect = self.get_window_rect(self.wechat_window_handle)
    
    # =============================================================================
    # 关键改进：从微信窗口内部获取元素，而不是从鼠标位置
    # =============================================================================
    def get_element_at_mouse_safe(self):
        """
        安全获取鼠标位置的元素（改进版）
        解决微信窗口堆叠时获取到下方窗口的问题
        """
        try:
            # 获取鼠标屏幕坐标
            x, y = auto.GetCursorPos()
            
            if self.debug_mode:
                print(f"\n[调试] 鼠标屏幕坐标: ({x}, {y})")
            
            # =============================================================================
            # 方法1：如果有微信窗口句柄，先检查鼠标是否在微信窗口内
            # =============================================================================
            if self.wechat_window_handle:
                # 更新窗口位置
                self.update_window_rect()
                
                if self.debug_mode and self.wechat_window_rect:
                    print(f"  [调试] 微信窗口位置: {self.wechat_window_rect}")
                
                # 检查鼠标是否在微信窗口范围内
                if self.wechat_window_rect and self.is_point_in_window(x, y, self.wechat_window_rect):
                    if self.debug_mode:
                        print(f"  [调试] 鼠标在微信窗口范围内 ✓")
                    
                    # =============================================================================
                    # 关键：从微信窗口内部查找元素，而不是从屏幕位置
                    # =============================================================================
                    
                    # 方法A: 尝试从微信窗口句柄获取元素
                    if self.wechat_window:
                        try:
                            # 将屏幕坐标转换为客户端坐标
                            client_x, client_y = x - self.wechat_window_rect[0], y - self.wechat_window_rect[1]
                            
                            if self.debug_mode:
                                print(f"  [调试] 鼠标客户端坐标: ({client_x}, {client_y})")
                            
                            # 从微信窗口元素查找
                            # 注意：uiautomation 没有直接从客户端坐标获取元素的方法
                            # 所以我们尝试两种方式：
                            
                            # 方式1: 直接使用 ControlFromPoint（但可能穿透）
                            # 我们先尝试，然后验证元素是否属于微信
                            element = auto.ControlFromPoint(x, y)
                            
                            if element:
                                # 验证元素是否属于微信窗口
                                if self.is_element_belongs_to_wechat(element):
                                    if self.debug_mode:
                                        print(f"  [调试] ControlFromPoint 获取的元素属于微信 ✓")
                                    return element
                                else:
                                    if self.debug_mode:
                                        print(f"  [调试] ControlFromPoint 获取的元素不属于微信 ✗")
                                        print(f"  [调试] 尝试从微信窗口遍历查找...")
                                    
                                    # 方式2: 遍历微信窗口的子元素查找
                                    element = self.find_element_in_wechat_at_position(client_x, client_y)
                                    if element:
                                        return element
                        except Exception as e:
                            if self.debug_mode:
                                print(f"  [调试] 从微信窗口获取元素失败: {e}")
                    
                    # 方法B: 使用 win32gui 获取窗口内的元素
                    if WIN32_AVAILABLE and win32gui:
                        try:
                            # 使用 ChildWindowFromPoint 获取子窗口
                            # 注意：这只对传统 Win32 控件有效，对 DirectUI 无效
                            pass
                        except:
                            pass
                else:
                    if self.debug_mode:
                        print(f"  [调试] 鼠标不在微信窗口范围内 ✗")
                    return None
            
            # =============================================================================
            # 方法2：如果没有句柄或上述方法失败，使用传统方式但增加验证
            # =============================================================================
            if self.debug_mode:
                print(f"  [调试] 使用传统方式获取元素...")
            
            try:
                element = auto.ControlFromPoint(x, y)
                
                if element:
                    # 验证元素是否属于微信
                    if self.is_element_belongs_to_wechat(element):
                        if self.debug_mode:
                            print(f"  [调试] 元素属于微信 ✓")
                        return element
                    else:
                        if self.debug_mode:
                            print(f"  [调试] 元素不属于微信 ✗，跳过")
                        return None
            except Exception as e:
                if self.debug_mode:
                    print(f"  [调试] ControlFromPoint 失败: {e}")
            
            return None
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 获取鼠标位置元素时出错: {e}")
            return None
    
    def is_element_belongs_to_wechat(self, element):
        """检查元素是否属于微信窗口"""
        try:
            if not element:
                return False
            
            # 方法1: 如果有微信窗口句柄，检查元素的窗口句柄
            try:
                element_handle = element.NativeWindowHandle
                if element_handle:
                    # 检查是否是微信窗口本身
                    if element_handle == self.wechat_window_handle:
                        return True
                    
                    # 检查是否是微信窗口的子窗口
                    if WIN32_AVAILABLE and win32gui and self.wechat_window_handle:
                        try:
                            # 获取祖先窗口
                            ancestor = win32gui.GetAncestor(element_handle, 2)  # GA_ROOT
                            if ancestor == self.wechat_window_handle:
                                return True
                        except:
                            pass
            except:
                pass
            
            # 方法2: 向上遍历元素树查找微信窗口
            current = element
            max_levels = 30
            level = 0
            
            while current and level < max_levels:
                try:
                    class_name = current.ClassName if current.ClassName else ""
                    name = current.Name if current.Name else ""
                    
                    # 检查类名
                    if WECHAT_CLASS_NAME in class_name:
                        if self.debug_mode:
                            print(f"  [调试] 元素祖先匹配微信类名: {class_name}")
                        return True
                    
                    # 检查名称
                    if WECHAT_TITLE in name and 'Window' in str(type(current)):
                        if self.debug_mode:
                            print(f"  [调试] 元素祖先匹配微信标题: {name}")
                        return True
                    
                    # 获取父元素
                    current = current.GetParentControl()
                    level += 1
                    
                except Exception:
                    break
            
            return False
            
        except Exception:
            return False
    
    def find_element_in_wechat_at_position(self, client_x, client_y):
        """
        在微信窗口中查找指定客户端坐标位置的元素
        这是一个备用方法，当 ControlFromPoint 穿透窗口时使用
        """
        try:
            if not self.wechat_window:
                return None
            
            if self.debug_mode:
                print(f"  [调试] 在微信窗口中遍历查找坐标 ({client_x}, {client_y}) 处的元素...")
            
            # 获取微信窗口的子元素
            def find_clickable_element(element, depth=0):
                if depth > 10:
                    return None
                
                try:
                    # 获取元素的边界矩形
                    rect = element.BoundingRectangle
                    if rect:
                        # 检查坐标是否在元素范围内
                        left = rect.left
                        top = rect.top
                        right = rect.right
                        bottom = rect.bottom
                        
                        if left <= client_x <= right and top <= client_y <= bottom:
                            # 检查这个元素是否是文件元素或包含文件元素
                            name = element.Name if element.Name else ""
                            class_name = element.ClassName if element.ClassName else ""
                            
                            if self.debug_mode:
                                print(f"  [调试]   找到匹配元素 (深度 {depth}): 类='{class_name}', 名='{name[:50] if name else ''}'")
                            
                            # 返回这个元素或其父元素
                            return element
                    
                    # 递归检查子元素
                    children = element.GetChildren()
                    for child in children:
                        found = find_clickable_element(child, depth + 1)
                        if found:
                            return found
                            
                except Exception:
                    pass
                
                return None
            
            # 从微信窗口开始查找
            result = find_clickable_element(self.wechat_window)
            return result
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 遍历查找元素失败: {e}")
            return None
    
    # =============================================================================
    # 以下是辅助方法
    # =============================================================================
    def _get_element_name(self, element):
        try:
            if not element:
                return ""
            name = element.Name
            return name if name else ""
        except Exception:
            return ""
    
    def _get_element_class(self, element):
        try:
            if not element:
                return ""
            class_name = element.ClassName
            return class_name if class_name else ""
        except Exception:
            return ""
    
    def _get_children(self, element):
        try:
            if not element:
                return []
            children = element.GetChildren()
            return children if children else []
        except Exception:
            return []
    
    def print_element_tree(self, element, max_depth=4, indent=0):
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
    
    def _is_file_element(self, element):
        """检查元素是否为文件类型的消息元素"""
        try:
            if not element:
                return False
            
            name = self._get_element_name(element)
            class_name = self._get_element_class(element)
            
            if self.debug_mode:
                print(f"  [调试] 检查是否为文件元素: 名称='{name[:50] if name else ''}', 类名='{class_name}'")
            
            if not name:
                return False
            
            # 规则1：包含文件关键词
            for keyword in FILE_KEYWORDS:
                if keyword in name:
                    if self.debug_mode:
                        print(f"  [调试] ✓ 匹配文件关键词: '{keyword}'")
                    return True
            
            # 规则2：类名包含 File
            if 'File' in class_name or 'file' in class_name:
                if self.debug_mode:
                    print(f"  [调试] ✓ 类名包含 File")
                return True
            
            # 规则3：文件大小 + 扩展名
            has_file_size = bool(re.search(FILE_SIZE_PATTERN, name))
            
            # 更严格的扩展名匹配
            file_ext_pattern = r'(^|\s|文件:|文件)([^\s\\/:*?"<>|\r\n]+\.[a-zA-Z0-9]{2,4})'
            ext_match = re.search(file_ext_pattern, name)
            
            if self.debug_mode:
                print(f"  [调试]   - 包含文件大小: {has_file_size}")
                print(f"  [调试]   - 包含文件扩展名: {bool(ext_match)}")
            
            if has_file_size and ext_match:
                if self.debug_mode:
                    print(f"  [调试] ✓ 文件大小 + 扩展名匹配")
                return True
            
            # 排除 URL
            if 'http://' in name or 'https://' in name:
                if self.debug_mode:
                    print(f"  [调试] ✗ 包含 URL，排除")
                return False
            
            return False
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 检查文件元素时出错: {e}")
            return False
    
    def _is_valid_path(self, path_str):
        if not path_str or not isinstance(path_str, str):
            return False
        
        path_str = path_str.strip()
        if not path_str:
            return False
        
        if '\\' not in path_str and '/' not in path_str:
            return False
        
        if re.match(r'^[A-Za-z]:', path_str):
            return True
        
        if 'WeChat Files' in path_str or '微信文件' in path_str:
            return True
        
        if path_str.startswith('\\\\'):
            return True
        
        return False
    
    def _normalize_path(self, path_str):
        if not path_str:
            return ""
        
        path_str = path_str.strip()
        path_str = path_str.replace('/', '\\')
        path_str = re.sub(r'\s+', ' ', path_str).strip()
        path_str = path_str.strip('"\'')
        return path_str
    
    def _get_wechat_file_dir(self):
        if self.wechat_file_dir:
            return self.wechat_file_dir
        
        self.wechat_file_dir = self._find_wechat_file_directory()
        return self.wechat_file_dir
    
    def _find_wechat_file_directory(self):
        try:
            if os.path.exists(DEFAULT_WECHAT_FILE_PATH):
                for item in os.listdir(DEFAULT_WECHAT_FILE_PATH):
                    item_path = os.path.join(DEFAULT_WECHAT_FILE_PATH, item)
                    if os.path.isdir(item_path):
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
            
            return None
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 查找微信文件目录时出错: {e}")
            return None
    
    def _find_file_in_directory(self, directory, file_name):
        try:
            if not directory or not os.path.exists(directory):
                return None
            
            if self.debug_mode:
                print(f"  [调试] 在微信目录中搜索文件: {file_name}")
            
            wechat_specific_dirs = ['File', 'FileStorage', 'Video', 'Image', 'Attachment']
            
            for specific_dir in wechat_specific_dirs:
                specific_path = os.path.join(directory, specific_dir)
                if os.path.exists(specific_path):
                    for root, dirs, files in os.walk(specific_path):
                        depth = root[len(specific_path):].count(os.sep)
                        if depth > 4:
                            continue
                        
                        if file_name in files:
                            full_path = os.path.join(root, file_name)
                            if 'WeChat Files' in full_path or '微信文件' in full_path:
                                if self.debug_mode:
                                    print(f"  [调试] ✓ 找到文件: {full_path}")
                                return full_path
            
            for root, dirs, files in os.walk(directory):
                depth = root[len(directory):].count(os.sep)
                if depth > 5:
                    continue
                
                if file_name in files:
                    full_path = os.path.join(root, file_name)
                    if 'WeChat Files' in full_path or '微信文件' in full_path:
                        if self.debug_mode:
                            print(f"  [调试] ✓ 找到文件: {full_path}")
                        return full_path
            
            return None
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 搜索文件时出错: {e}")
            return None
    
    def _extract_file_name(self, name):
        if not name:
            return None
        
        # 带文件大小的格式
        pattern_with_size = r'([^\s\\/:*?"<>|\r\n]+\.[a-zA-Z0-9]{2,4})\s*\(\d+\.?\d*\s*[KMGT]?B\)'
        match = re.search(pattern_with_size, name)
        if match:
            file_name = match.group(1)
            if self.debug_mode:
                print(f"  [调试] 从带大小格式提取: {file_name}")
            return file_name
        
        # 带"文件:"前缀的格式
        pattern_with_prefix = r'文件[:：]\s*([^\s\\/:*?"<>|\r\n]+\.[a-zA-Z0-9]{2,4})'
        match = re.search(pattern_with_prefix, name)
        if match:
            file_name = match.group(1)
            if self.debug_mode:
                print(f"  [调试] 从带前缀格式提取: {file_name}")
            return file_name
        
        # 简单格式（最后尝试）
        pattern_simple = r'([^\s\\/:*?"<>|\r\n]+\.[a-zA-Z0-9]{2,4})'
        matches = re.findall(pattern_simple, name)
        
        if matches:
            excluded_exts = ['.txt', '.py', '.js', '.html', '.css', '.json', '.xml', '.md']
            for file_name in matches:
                ext = os.path.splitext(file_name)[1].lower()
                if ext not in excluded_exts:
                    if self.debug_mode:
                        print(f"  [调试] 从简单格式提取: {file_name}")
                    return file_name
        
        return None
    
    def _extract_path_from_file_element(self, element):
        try:
            if not element:
                return None
            
            name = self._get_element_name(element)
            
            if self._is_valid_path(name):
                return self._normalize_path(name)
            
            file_name = self._extract_file_name(name)
            if not file_name:
                return None
            
            if self.debug_mode:
                print(f"  [调试] 提取到文件名: {file_name}")
            
            wechat_path = self._get_wechat_file_dir()
            if not wechat_path:
                return None
            
            full_path = self._find_file_in_directory(wechat_path, file_name)
            if full_path:
                return full_path
            
            return None
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 提取路径时出错: {e}")
            return None
    
    def _find_file_path_in_children(self, element):
        try:
            if not element:
                return None
            
            children = self._get_children(element)
            
            for child in children:
                child_name = self._get_element_name(child)
                
                if self._is_valid_path(child_name):
                    return self._normalize_path(child_name)
                
                nested_path = self._find_file_path_in_children(child)
                if nested_path:
                    return nested_path
                
                if self._is_file_element(child):
                    path = self._extract_path_from_file_element(child)
                    if path:
                        return path
            
            return None
            
        except Exception as e:
            if self.debug_mode:
                print(f"  [调试] 在子元素中查找路径时出错: {e}")
            return None
    
    def get_file_path_from_element(self, element):
        try:
            if not element:
                return None
            
            if self.debug_mode:
                print("\n" + "="*60)
                print("[调试] 开始分析元素...")
                print("="*60)
                self.print_element_tree(element, max_depth=3)
            
            name = self._get_element_name(element)
            class_name = self._get_element_class(element)
            
            if self.debug_mode:
                print(f"\n[调试] 当前元素信息:")
                print(f"  名称: '{name[:80] if name else ''}'")
                print(f"  类名: '{class_name}'")
            
            if self._is_valid_path(name):
                path = self._normalize_path(name)
                if self.debug_mode:
                    print(f"[调试] 方法1成功: {path}")
                return path
            
            file_path = self._find_file_path_in_children(element)
            if file_path:
                if self.debug_mode:
                    print(f"[调试] 方法2成功: {file_path}")
                return file_path
            
            if self._is_file_element(element):
                path = self._extract_path_from_file_element(element)
                if path:
                    if self.debug_mode:
                        print(f"[调试] 方法3成功: {path}")
                    return path
            
            if self.debug_mode:
                print("[调试] 未能提取路径")
            
            return None
            
        except Exception as e:
            print(f"提取文件路径时出错: {e}")
            return None
    
    def check_and_copy_file_path(self, element):
        try:
            if not element:
                return False
            
            file_path = self.get_file_path_from_element(element)
            
            if file_path:
                if file_path == self.last_clicked_file:
                    if self.debug_mode:
                        print(f"[调试] 路径与上次相同，跳过")
                    return False
                
                pyperclip.copy(file_path)
                self.last_clicked_file = file_path
                
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
    # 鼠标点击检测
    # =============================================================================
    def check_left_mouse_click(self):
        if not WIN32_AVAILABLE or not win32api:
            return False
        
        try:
            state = win32api.GetAsyncKeyState(win32con.VK_LBUTTON)
            return state < 0
        except Exception:
            return False
    
    def check_right_mouse_click(self):
        if not WIN32_AVAILABLE or not win32api:
            return False
        
        try:
            state = win32api.GetAsyncKeyState(win32con.VK_RBUTTON)
            return state < 0
        except Exception:
            return False
    
    def monitor_mouse_clicks(self):
        print("\n开始监控鼠标点击事件...")
        print("="*60)
        print("使用方法:")
        print("  1. 确保微信窗口在前台")
        print("  2. 在微信聊天窗口中找到文件消息")
        print("  3. 用鼠标左键或右键点击文件消息")
        print("="*60)
        print(f"调试模式: {'开启' if self.debug_mode else '关闭'}")
        print(f"pywin32: {'✓ 已加载' if WIN32_AVAILABLE else '✗ 未加载'}")
        print("按 Ctrl+C 停止监控\n")
        
        self.monitoring = True
        
        left_was_pressed = False
        right_was_pressed = False
        
        DEBOUNCE_DELAY = 0.3
        last_left_click_time = 0
        last_right_click_time = 0
        
        try:
            while self.monitoring:
                current_time = time.time()
                
                if WIN32_AVAILABLE:
                    left_pressed = self.check_left_mouse_click()
                    
                    if left_was_pressed and not left_pressed:
                        if current_time - last_left_click_time > DEBOUNCE_DELAY:
                            if self.debug_mode:
                                print("\n[调试] 检测到鼠标左键点击")
                            
                            # 使用改进的方法获取元素
                            element = self.get_element_at_mouse_safe()
                            
                            if element:
                                self.check_and_copy_file_path(element)
                            
                            last_left_click_time = current_time
                    
                    left_was_pressed = left_pressed
                    
                    right_pressed = self.check_right_mouse_click()
                    
                    if right_was_pressed and not right_pressed:
                        if current_time - last_right_click_time > DEBOUNCE_DELAY:
                            if self.debug_mode:
                                print("\n[调试] 检测到鼠标右键点击")
                            
                            element = self.get_element_at_mouse_safe()
                            
                            if element:
                                self.check_and_copy_file_path(element)
                            
                            last_right_click_time = current_time
                    
                    right_was_pressed = right_pressed
                
                else:
                    # 没有 pywin32 时使用简单检测
                    current_x, current_y = auto.GetCursorPos()
                    
                    if not hasattr(self, '_last_pos'):
                        self._last_pos = (current_x, current_y)
                        self._last_pos_time = current_time
                    
                    if (current_x, current_y) == self._last_pos:
                        if current_time - self._last_pos_time > 0.5:
                            if current_time - (getattr(self, '_last_simple_check', 0)) > 1.0:
                                element = self.get_element_at_mouse_safe()
                                if element:
                                    self.check_and_copy_file_path(element)
                                self._last_simple_check = current_time
                    else:
                        self._last_pos = (current_x, current_y)
                        self._last_pos_time = current_time
                
                time.sleep(0.02)
                
        except KeyboardInterrupt:
            print("\n\n监控已停止")
        finally:
            self.monitoring = False
    
    def start(self):
        print("="*60)
        print("微信文件消息路径提取工具 v2.2")
        print("="*60)
        print()
        
        if not WIN32_AVAILABLE:
            print("="*60)
            print("⚠️  重要提示：pywin32 未成功加载")
            print("="*60)
            print("可能的原因:")
            print("  1. 运行脚本的 Python 环境与安装 pywin32 的环境不同")
            print("  2. pywin32 安装不完整")
            print()
            print("建议操作:")
            print("  1. 检查 Python 路径是否正确")
            print("  2. 尝试重新安装: pip install --upgrade pywin32")
            print("  3. 如果安装后仍有问题，尝试运行:")
            print(f"     {sys.executable} -m pip install pywin32")
            print("="*60)
            print()
        
        if not self.find_wechat_window():
            print("\n继续运行于检测模式...")
            print()
        
        print("正在初始化微信文件目录...")
        wechat_dir = self._get_wechat_file_dir()
        if wechat_dir:
            print(f"✓ 微信文件目录: {wechat_dir}")
        else:
            print("⚠️  未找到微信文件目录")
        print()
        
        try:
            self.monitor_mouse_clicks()
        except Exception as e:
            print(f"监控过程中出错: {e}")
            import traceback
            traceback.print_exc()
            return False
        
        return True


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='微信文件消息路径提取工具')
    parser.add_argument('--no-debug', action='store_true', help='关闭调试模式')
    parser.add_argument('--debug', action='store_true', help='开启调试模式（默认）')
    
    args = parser.parse_args()
    
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
