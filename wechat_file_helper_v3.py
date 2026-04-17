# -*- coding: utf-8 -*-
"""
PC版微信文件消息路径提取工具 v3.0
针对 Qt 版微信的替代方案

重要发现：
- 新版微信使用 Qt 5.15.14 框架
- UI Automation 只能获取到外层 Qt 窗口，无法获取内部控件
- 类名: Qt51514QWindowIcon

替代方案：
1. 剪贴板监控 - 当用户复制文件时自动获取路径
2. 微信文件目录监控 - 显示最近接收的文件
3. 快捷键触发 - 用户按下快捷键时显示最近文件列表
"""

import pyperclip
import time
import os
import sys
import re
import threading
from datetime import datetime
from pathlib import Path

# =============================================================================
# 尝试导入可选依赖
# =============================================================================
try:
    import keyboard
    KEYBOARD_AVAILABLE = True
except ImportError:
    KEYBOARD_AVAILABLE = False
    print("[提示] 未安装 keyboard 库，快捷键功能不可用")
    print("       可运行: pip install keyboard")

try:
    import win32api
    import win32con
    import win32gui
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("[提示] 未安装 pywin32，部分功能受限")

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    WATCHDOG_AVAILABLE = True
except ImportError:
    WATCHDOG_AVAILABLE = False
    print("[提示] 未安装 watchdog 库，实时文件监控不可用")
    print("       可运行: pip install watchdog")

# 微信默认文件保存路径
DEFAULT_WECHAT_FILE_PATH = os.path.join(
    os.environ.get('USERPROFILE', ''),
    'Documents', 'WeChat Files'
)

# 快捷键
HOTKEY_SHOW_RECENT = 'ctrl+alt+w'  # 显示最近文件
HOTKEY_COPY_LATEST = 'ctrl+alt+e'  # 复制最新文件路径

# 监控的文件扩展名（常见文件类型）
TARGET_EXTENSIONS = [
    '.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx',
    '.txt', '.rtf', '.zip', '.rar', '.7z', '.tar', '.gz',
    '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.mp4', '.avi',
    '.mkv', '.mov', '.mp3', '.wav', '.flac', '.aac',
    '.exe', '.msi', '.apk', '.ipa', '.dmg',
]


class WeChatFileMonitorV3:
    """微信文件监控 v3.0 - Qt 版微信替代方案"""
    
    def __init__(self):
        self.wechat_file_dir = None
        self.recent_files = []
        self.monitoring = False
        self.last_clipboard_content = None
        self.observer = None
        
    def find_wechat_file_directory(self):
        """查找微信文件保存目录"""
        print("="*60)
        print("正在查找微信文件目录...")
        print("="*60)
        
        # 可能的微信文件路径
        possible_paths = [
            DEFAULT_WECHAT_FILE_PATH,
            os.path.join(os.environ.get('USERPROFILE', ''), 'WeChat Files'),
            os.path.join(os.environ.get('APPDATA', ''), 'Tencent', 'WeChat'),
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Tencent', 'WeChat'),
        ]
        
        found_paths = []
        
        for base_path in possible_paths:
            if os.path.exists(base_path):
                print(f"\n检查路径: {base_path}")
                
                # 查找用户文件夹
                try:
                    for item in os.listdir(base_path):
                        item_path = os.path.join(base_path, item)
                        if os.path.isdir(item_path):
                            # 检查是否包含 FileStorage 或 Msg 目录
                            file_storage = os.path.join(item_path, 'FileStorage')
                            if os.path.exists(file_storage):
                                found_paths.append(file_storage)
                                print(f"  ✓ 找到 FileStorage: {file_storage}")
                                
                                # 检查子目录
                                for sub in ['File', 'Video', 'Image', 'Attachment']:
                                    sub_path = os.path.join(file_storage, sub)
                                    if os.path.exists(sub_path):
                                        print(f"    - {sub}: {sub_path}")
                            
                            msg_dir = os.path.join(item_path, 'Msg')
                            if os.path.exists(msg_dir):
                                found_paths.append(msg_dir)
                                print(f"  ✓ 找到 Msg: {msg_dir}")
                except Exception as e:
                    print(f"  ✗ 遍历失败: {e}")
        
        if found_paths:
            # 选择第一个找到的路径作为主要监控目录
            self.wechat_file_dir = found_paths[0]
            print(f"\n✓ 主监控目录: {self.wechat_file_dir}")
            return found_paths
        
        print("\n✗ 未找到微信文件目录")
        print("提示: 请确保微信已登录并接收过文件")
        return []
    
    def get_recent_files(self, count=10, extensions=None):
        """
        获取最近的文件
        
        Args:
            count: 返回文件数量
            extensions: 要筛选的扩展名列表，None 表示所有文件
        
        Returns:
            list: [(文件路径, 修改时间), ...]
        """
        if not self.wechat_file_dir:
            return []
        
        files_with_time = []
        
        try:
            # 遍历目录
            for root, dirs, files in os.walk(self.wechat_file_dir):
                # 限制深度
                depth = root[len(self.wechat_file_dir):].count(os.sep)
                if depth > 4:
                    continue
                
                for file_name in files:
                    # 筛选扩展名
                    if extensions:
                        file_ext = os.path.splitext(file_name)[1].lower()
                        if file_ext not in extensions:
                            continue
                    
                    file_path = os.path.join(root, file_name)
                    
                    try:
                        # 获取修改时间
                        mtime = os.path.getmtime(file_path)
                        files_with_time.append((file_path, mtime))
                    except:
                        continue
        except Exception as e:
            print(f"[错误] 遍历目录失败: {e}")
        
        # 按修改时间排序（最新的在前）
        files_with_time.sort(key=lambda x: x[1], reverse=True)
        
        # 返回前 count 个
        return files_with_time[:count]
    
    def display_recent_files(self, files=None, count=10):
        """显示最近的文件列表"""
        if files is None:
            files = self.get_recent_files(count=count, extensions=TARGET_EXTENSIONS)
        
        if not files:
            print("\n[提示] 未找到最近的文件")
            return []
        
        print("\n" + "="*60)
        print(f"最近 {len(files)} 个文件 (按时间排序):")
        print("="*60)
        
        for i, (file_path, mtime) in enumerate(files, 1):
            file_name = os.path.basename(file_path)
            time_str = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
            
            # 获取文件大小
            try:
                size = os.path.getsize(file_path)
                if size < 1024:
                    size_str = f"{size} B"
                elif size < 1024*1024:
                    size_str = f"{size/1024:.1f} KB"
                elif size < 1024*1024*1024:
                    size_str = f"{size/(1024*1024):.1f} MB"
                else:
                    size_str = f"{size/(1024*1024*1024):.1f} GB"
            except:
                size_str = "未知"
            
            print(f"\n{i}. [{time_str}] [{size_str}]")
            print(f"   文件名: {file_name}")
            print(f"   路径: {file_path}")
        
        print("\n" + "="*60)
        print("提示: 输入数字复制对应文件路径，输入 'q' 退出")
        print("="*60)
        
        return files
    
    def interactive_copy(self, files):
        """交互式复制文件路径"""
        if not files:
            return
        
        while True:
            try:
                choice = input("\n请选择要复制的文件编号 (或 'q' 退出): ").strip()
                
                if choice.lower() == 'q':
                    break
                
                if choice.isdigit():
                    index = int(choice) - 1
                    if 0 <= index < len(files):
                        file_path, _ = files[index]
                        pyperclip.copy(file_path)
                        print(f"\n✅ 已复制到剪贴板:")
                        print(f"   {file_path}")
                    else:
                        print(f"❌ 无效的编号: {choice}")
                else:
                    print("❌ 请输入有效的数字或 'q'")
                    
            except KeyboardInterrupt:
                print("\n\n已退出")
                break
            except EOFError:
                break
    
    # =============================================================================
    # 剪贴板监控
    # =============================================================================
    def monitor_clipboard(self):
        """监控剪贴板变化"""
        print("\n[剪贴板监控] 已启动")
        print("提示: 当您在微信中复制文件时，文件路径将自动显示")
        print("按 Ctrl+C 停止监控\n")
        
        self.last_clipboard_content = pyperclip.paste()
        
        try:
            while self.monitoring:
                try:
                    current_content = pyperclip.paste()
                    
                    if current_content != self.last_clipboard_content:
                        self.last_clipboard_content = current_content
                        
                        # 检查是否是文件路径
                        if self._is_valid_path(current_content):
                            path = self._normalize_path(current_content)
                            print(f"\n{'='*60}")
                            print(f"✅ 检测到剪贴板中的文件路径:")
                            print(f"   {path}")
                            print(f"{'='*60}\n")
                        elif os.path.exists(current_content.strip()):
                            path = current_content.strip()
                            print(f"\n{'='*60}")
                            print(f"✅ 检测到剪贴板中的文件路径:")
                            print(f"   {path}")
                            print(f"{'='*60}\n")
                        
                except Exception:
                    pass
                
                time.sleep(0.1)
                
        except KeyboardInterrupt:
            print("\n剪贴板监控已停止")
    
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
        
        # 检查是否包含驱动器号
        if re.match(r'^[A-Za-z]:', path_str):
            return True
        
        # 检查是否包含 WeChat Files
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
    # 文件系统监控（使用 watchdog）
    # =============================================================================
    def start_file_watcher(self):
        """启动文件系统监控"""
        if not WATCHDOG_AVAILABLE:
            print("[提示] watchdog 未安装，无法启动实时文件监控")
            return False
        
        if not self.wechat_file_dir:
            print("[错误] 未找到微信文件目录")
            return False
        
        class WeChatFileHandler(FileSystemEventHandler):
            def __init__(self, monitor):
                self.monitor = monitor
            
            def on_created(self, event):
                if not event.is_directory:
                    self.monitor._on_file_created(event.src_path)
            
            def on_modified(self, event):
                if not event.is_directory:
                    self.monitor._on_file_modified(event.src_path)
        
        try:
            event_handler = WeChatFileHandler(self)
            self.observer = Observer()
            self.observer.schedule(event_handler, self.wechat_file_dir, recursive=True)
            self.observer.start()
            
            print(f"\n[实时文件监控] 已启动")
            print(f"监控目录: {self.wechat_file_dir}")
            print("当微信接收新文件时，将自动显示路径\n")
            
            return True
            
        except Exception as e:
            print(f"[错误] 启动文件监控失败: {e}")
            return False
    
    def _on_file_created(self, file_path):
        """文件创建时的回调"""
        self._check_and_notify(file_path, "新建")
    
    def _on_file_modified(self, file_path):
        """文件修改时的回调"""
        self._check_and_notify(file_path, "修改")
    
    def _check_and_notify(self, file_path, action):
        """检查并通知用户"""
        # 过滤临时文件
        file_name = os.path.basename(file_path)
        if file_name.startswith('~$') or file_name.startswith('.'):
            return
        
        # 检查扩展名
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in TARGET_EXTENSIONS and ext:
            return
        
        print(f"\n{'='*60}")
        print(f"✅ 检测到文件{action}:")
        print(f"   {file_path}")
        print(f"\n提示: 路径已自动复制到剪贴板！")
        print(f"{'='*60}\n")
        
        # 自动复制到剪贴板
        pyperclip.copy(file_path)
    
    # =============================================================================
    # 快捷键功能
    # =============================================================================
    def setup_hotkeys(self):
        """设置快捷键"""
        if not KEYBOARD_AVAILABLE:
            return False
        
        def on_show_recent():
            print("\n[快捷键] 显示最近文件...")
            files = self.display_recent_files(count=10)
            if files:
                self.interactive_copy(files)
        
        def on_copy_latest():
            print("\n[快捷键] 复制最新文件...")
            files = self.get_recent_files(count=1, extensions=TARGET_EXTENSIONS)
            if files:
                file_path, _ = files[0]
                pyperclip.copy(file_path)
                print(f"✅ 已复制最新文件路径:")
                print(f"   {file_path}")
            else:
                print("❌ 未找到文件")
        
        try:
            keyboard.add_hotkey(HOTKEY_SHOW_RECENT, on_show_recent)
            keyboard.add_hotkey(HOTKEY_COPY_LATEST, on_copy_latest)
            
            print(f"\n[快捷键] 已注册:")
            print(f"  {HOTKEY_SHOW_RECENT} - 显示最近文件")
            print(f"  {HOTKEY_COPY_LATEST} - 复制最新文件路径")
            
            return True
            
        except Exception as e:
            print(f"[错误] 设置快捷键失败: {e}")
            return False
    
    # =============================================================================
    # 主菜单
    # =============================================================================
    def show_menu(self):
        """显示主菜单"""
        while True:
            print("\n" + "="*60)
            print("微信文件助手 v3.0 (Qt 版微信专用)")
            print("="*60)
            print("\n请选择功能:")
            print("  1. 显示最近接收的文件")
            print("  2. 启动剪贴板监控")
            print("  3. 启动实时文件监控")
            print("  4. 注册快捷键")
            print("  5. 综合模式（推荐）")
            print("  0. 退出")
            
            try:
                choice = input("\n请输入选项: ").strip()
                
                if choice == '0':
                    print("再见！")
                    break
                
                elif choice == '1':
                    files = self.display_recent_files(count=15)
                    if files:
                        self.interactive_copy(files)
                
                elif choice == '2':
                    self.monitoring = True
                    try:
                        self.monitor_clipboard()
                    finally:
                        self.monitoring = False
                
                elif choice == '3':
                    if self.start_file_watcher():
                        self.monitoring = True
                        print("按 Ctrl+C 停止监控\n")
                        try:
                            while self.monitoring:
                                time.sleep(1)
                        except KeyboardInterrupt:
                            print("\n正在停止...")
                        finally:
                            if self.observer:
                                self.observer.stop()
                                self.observer.join()
                                self.monitoring = False
                
                elif choice == '4':
                    if self.setup_hotkeys():
                        print("\n快捷键已生效，按 Ctrl+C 退出")
                        try:
                            while True:
                                time.sleep(1)
                        except KeyboardInterrupt:
                            print("\n已退出")
                
                elif choice == '5':
                    print("\n[综合模式] 启动中...")
                    
                    # 1. 启动剪贴板监控线程
                    clipboard_thread = threading.Thread(target=self.monitor_clipboard, daemon=True)
                    self.monitoring = True
                    clipboard_thread.start()
                    
                    # 2. 启动实时文件监控
                    self.start_file_watcher()
                    
                    # 3. 注册快捷键
                    self.setup_hotkeys()
                    
                    # 4. 显示最近文件
                    files = self.display_recent_files(count=10)
                    
                    print("\n" + "="*60)
                    print("综合模式已启动！")
                    print("="*60)
                    print("功能:")
                    print("  ✓ 剪贴板监控 - 复制文件时自动显示路径")
                    print("  ✓ 实时文件监控 - 新文件接收时自动复制路径")
                    print(f"  ✓ 快捷键 {HOTKEY_SHOW_RECENT} - 显示最近文件")
                    print(f"  ✓ 快捷键 {HOTKEY_COPY_LATEST} - 复制最新文件")
                    print("="*60)
                    print("\n按 Ctrl+C 停止所有监控\n")
                    
                    if files:
                        self.interactive_copy(files)
                    
                    try:
                        while self.monitoring:
                            time.sleep(1)
                    except KeyboardInterrupt:
                        print("\n正在停止所有监控...")
                        self.monitoring = False
                        if self.observer:
                            self.observer.stop()
                            self.observer.join()
                        print("已停止")
                
                else:
                    print("无效的选项")
                    
            except KeyboardInterrupt:
                print("\n\n已退出")
                break
    
    def start(self):
        """启动"""
        print("="*60)
        print("微信文件助手 v3.0")
        print("="*60)
        print("\n重要说明:")
        print("  检测到您的微信使用 Qt 框架 (Qt51514QWindowIcon)")
        print("  UI Automation 无法获取微信内部控件")
        print("  本工具使用替代方案实现功能")
        print()
        
        # 查找微信文件目录
        paths = self.find_wechat_file_directory()
        
        if not paths:
            print("\n警告: 未找到微信文件目录")
            print("部分功能可能不可用\n")
        
        # 显示主菜单
        self.show_menu()


def main():
    """主函数"""
    monitor = WeChatFileMonitorV3()
    
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
