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
from ctypes import windll

# 微信窗口的类名和标题
WECHAT_CLASS_NAME = 'WeChatMainWndForPC'
WECHAT_TITLE = '微信'

# 文件消息的特征关键词
FILE_KEYWORDS = ['文件', '文件:', '已接收文件', '接收文件']

# 微信默认文件保存路径（用于路径拼接）
DEFAULT_WECHAT_FILE_PATH = os.path.join(
    os.environ.get('USERPROFILE', ''),
    'Documents', 'WeChat Files'
)


class WeChatFileMonitor:
    """微信文件消息监控类"""
    
    def __init__(self):
        self.wechat_window = None
        self.last_clicked_file = None
        self.monitoring = False
        
    def find_wechat_window(self):
        """查找微信主窗口"""
        try:
            # 尝试通过类名查找
            window = auto.WindowControl(ClassName=WECHAT_CLASS_NAME)
            if window.Exists(3):
                self.wechat_window = window
                return True
            
            # 尝试通过标题查找
            window = auto.WindowControl(Name=WECHAT_TITLE)
            if window.Exists(3):
                self.wechat_window = window
                return True
                
            print("未找到微信窗口，请确保微信已登录并打开")
            return False
            
        except Exception as e:
            print(f"查找微信窗口时出错: {e}")
            return False
    
    def get_file_path_from_element(self, element):
        """从UI元素中提取文件路径"""
        try:
            # 获取元素的名称和其他属性
            name = element.Name if element.Name else ""
            automation_id = element.AutomationId if element.AutomationId else ""
            class_name = element.ClassName if element.ClassName else ""
            
            # 打印调试信息（可注释掉）
            # print(f"元素信息: 名称='{name}', ID='{automation_id}', 类名='{class_name}'")
            
            # 方法1: 直接从名称中提取路径（如果元素名称包含完整路径）
            if self._is_valid_path(name):
                return self._normalize_path(name)
            
            # 方法2: 检查子元素是否包含文件路径
            file_path = self._find_file_path_in_children(element)
            if file_path:
                return file_path
            
            # 方法3: 检查父元素和兄弟元素
            file_path = self._find_file_path_in_context(element)
            if file_path:
                return file_path
            
            # 方法4: 尝试获取文件大小或其他特征，然后构造路径
            file_path = self._construct_path_from_features(element)
            if file_path:
                return file_path
            
            return None
            
        except Exception as e:
            print(f"提取文件路径时出错: {e}")
            return None
    
    def _is_valid_path(self, path_str):
        """检查字符串是否为有效的文件路径"""
        if not path_str:
            return False
        
        # 检查是否包含常见的路径分隔符
        if '\\' not in path_str and '/' not in path_str:
            return False
        
        # 检查是否包含驱动器号（Windows路径）
        if re.match(r'^[A-Za-z]:', path_str):
            # 检查路径是否存在（可选，可能会影响性能）
            # try:
            #     if os.path.exists(path_str.strip()):
            #         return True
            # except:
            #     pass
            return True
        
        # 检查是否包含 WeChat Files 等微信相关路径
        if 'WeChat Files' in path_str or '微信文件' in path_str:
            return True
        
        return False
    
    def _normalize_path(self, path_str):
        """规范化文件路径"""
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
            # 获取所有子元素
            children = element.GetChildren()
            
            for child in children:
                child_name = child.Name if child.Name else ""
                
                # 检查子元素名称是否为有效路径
                if self._is_valid_path(child_name):
                    return self._normalize_path(child_name)
                
                # 递归检查更深层的子元素
                nested_path = self._find_file_path_in_children(child)
                if nested_path:
                    return nested_path
                
                # 检查是否包含文件大小等特征
                if self._is_file_element(child):
                    path = self._extract_path_from_file_element(child)
                    if path:
                        return path
            
            return None
            
        except Exception as e:
            print(f"在子元素中查找路径时出错: {e}")
            return None
    
    def _find_file_path_in_context(self, element):
        """在元素的上下文（父元素、兄弟元素）中查找文件路径"""
        try:
            # 检查父元素
            parent = element.GetParentControl()
            if parent:
                parent_name = parent.Name if parent.Name else ""
                
                if self._is_valid_path(parent_name):
                    return self._normalize_path(parent_name)
                
                # 检查父元素的其他子元素
                siblings = parent.GetChildren()
                for sibling in siblings:
                    if sibling == element:
                        continue
                    
                    sibling_name = sibling.Name if sibling.Name else ""
                    
                    if self._is_valid_path(sibling_name):
                        return self._normalize_path(sibling_name)
                    
                    # 检查是否为文件元素
                    if self._is_file_element(sibling):
                        path = self._extract_path_from_file_element(sibling)
                        if path:
                            return path
                
                # 递归向上检查
                return self._find_file_path_in_context(parent)
            
            return None
            
        except Exception as e:
            print(f"在上下文中查找路径时出错: {e}")
            return None
    
    def _is_file_element(self, element):
        """检查元素是否为文件类型的消息元素"""
        try:
            name = element.Name if element.Name else ""
            class_name = element.ClassName if element.ClassName else ""
            automation_id = element.AutomationId if element.AutomationId else ""
            
            # 检查名称中是否包含文件关键词
            for keyword in FILE_KEYWORDS:
                if keyword in name:
                    return True
            
            # 检查类名是否包含文件相关特征
            if 'File' in class_name or 'file' in class_name:
                return True
            
            # 检查是否包含文件扩展名
            if re.search(r'\.[a-zA-Z0-9]{2,4}', name):
                # 排除纯文本中的扩展名（需要结合其他特征）
                pass
            
            return False
            
        except Exception as e:
            print(f"检查文件元素时出错: {e}")
            return False
    
    def _extract_path_from_file_element(self, element):
        """从文件元素中提取路径"""
        try:
            name = element.Name if element.Name else ""
            
            # 尝试从名称中提取路径
            # 微信文件消息的格式可能是："文件名.ext (大小)" 或 "文件: 文件名.ext"
            
            # 提取文件名（包含扩展名）
            file_name_match = re.search(r'([^\\/:*?"<>|\r\n]+\.[a-zA-Z0-9]{2,4})', name)
            if file_name_match:
                file_name = file_name_match.group(1)
                
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
            print(f"从文件元素提取路径时出错: {e}")
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
                            return file_storage
                        msg_dir = os.path.join(item_path, 'Msg')
                        if os.path.exists(msg_dir):
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
                    return path
            
            return None
            
        except Exception as e:
            print(f"查找微信文件目录时出错: {e}")
            return None
    
    def _find_file_in_directory(self, directory, file_name):
        """在目录中查找文件（递归）"""
        try:
            if not directory or not os.path.exists(directory):
                return None
            
            # 优先检查常见的文件接收目录
            common_dirs = ['File', 'FileStorage', 'Msg', 'Attachment', 'Files']
            
            for common_dir in common_dirs:
                common_path = os.path.join(directory, common_dir)
                if os.path.exists(common_path):
                    # 在该目录下查找
                    for root, dirs, files in os.walk(common_path):
                        if file_name in files:
                            return os.path.join(root, file_name)
            
            # 如果没找到，遍历整个目录
            for root, dirs, files in os.walk(directory):
                if file_name in files:
                    return os.path.join(root, file_name)
            
            return None
            
        except Exception as e:
            print(f"在目录中查找文件时出错: {e}")
            return None
    
    def _construct_path_from_features(self, element):
        """根据元素特征构造文件路径"""
        try:
            name = element.Name if element.Name else ""
            
            # 提取文件名
            file_name_match = re.search(r'([^\\/:*?"<>|\r\n]+\.[a-zA-Z0-9]{2,4})', name)
            if not file_name_match:
                return None
            
            file_name = file_name_match.group(1)
            
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
            print(f"根据特征构造路径时出错: {e}")
            return None
    
    def get_element_at_mouse(self):
        """获取鼠标位置的UI元素"""
        try:
            # 获取鼠标位置
            x, y = auto.GetCursorPos()
            
            # 从鼠标位置获取元素
            element = auto.ControlFromPoint(x, y)
            
            if element:
                return element
            
            return None
            
        except Exception as e:
            print(f"获取鼠标位置元素时出错: {e}")
            return None
    
    def check_and_copy_file_path(self, element):
        """检查元素是否为文件消息，如果是则复制路径到剪贴板"""
        try:
            # 获取文件路径
            file_path = self.get_file_path_from_element(element)
            
            if file_path:
                # 检查是否与上次复制的相同（避免重复复制）
                if file_path == self.last_clicked_file:
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
    
    def monitor_mouse_clicks(self):
        """监控鼠标点击事件"""
        print("开始监控鼠标点击事件...")
        print("提示: 当您在微信中点击或右键点击文件消息时，文件路径将自动复制到剪贴板。")
        print("按 Ctrl+C 停止监控\n")
        
        self.monitoring = True
        
        try:
            # 记录上次点击的位置（用于检测点击事件）
            last_x, last_y = auto.GetCursorPos()
            last_click_time = time.time()
            
            while self.monitoring:
                # 获取当前鼠标位置
                current_x, current_y = auto.GetCursorPos()
                current_time = time.time()
                
                # 简单的点击检测（位置短时间内不变可能表示点击）
                # 更准确的检测需要使用鼠标钩子，但为了简单起见，这里使用轮询方式
                
                # 检测鼠标位置是否稳定（可能表示点击）
                if (current_x == last_x and current_y == last_y and 
                    current_time - last_click_time > 0.1):
                    
                    # 获取鼠标位置的元素
                    element = self.get_element_at_mouse()
                    
                    if element:
                        # 检查是否为文件消息
                        self.check_and_copy_file_path(element)
                    
                    # 更新点击时间
                    last_click_time = current_time
                
                # 更新位置
                last_x, last_y = current_x, current_y
                
                # 短暂休眠，减少CPU占用
                time.sleep(0.05)
                
        except KeyboardInterrupt:
            print("\n监控已停止")
        finally:
            self.monitoring = False
    
    def monitor_foreground_window(self):
        """监控前台窗口，当微信在前台时增强监控"""
        print("开始监控前台窗口...")
        
        try:
            while self.monitoring:
                # 获取前台窗口
                foreground_window = auto.GetForegroundControl()
                
                if foreground_window:
                    # 检查是否为微信窗口
                    window_class = foreground_window.ClassName if foreground_window.ClassName else ""
                    window_name = foreground_window.Name if foreground_window.Name else ""
                    
                    if (WECHAT_CLASS_NAME in window_class or 
                        WECHAT_TITLE in window_name):
                        # 微信在前台，可以在这里添加额外的监控逻辑
                        pass
                
                time.sleep(0.5)
                
        except KeyboardInterrupt:
            pass
    
    def start(self):
        """启动监控"""
        print("="*60)
        print("微信文件消息路径提取工具")
        print("="*60)
        print()
        
        # 查找微信窗口
        print("正在查找微信窗口...")
        if not self.find_wechat_window():
            print("错误: 无法找到微信窗口")
            print("请确保:")
            print("  1. 微信已登录并运行")
            print("  2. 微信窗口未最小化到系统托盘")
            return False
        
        print("✓ 找到微信窗口")
        print()
        
        # 启动监控
        try:
            self.monitor_mouse_clicks()
        except Exception as e:
            print(f"监控过程中出错: {e}")
            return False
        
        return True
    
    def stop(self):
        """停止监控"""
        self.monitoring = False
        print("监控已停止")


def main():
    """主函数"""
    monitor = WeChatFileMonitor()
    
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
