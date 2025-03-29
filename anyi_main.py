
#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Anyi自动化脚本

该脚本实现以下功能：
1. 进程管理：检查并结束Anyi.exe和anyi-core.exe进程
2. 文件操作：删除用户.config文件夹
3. 程序运行：运行用户指定的Anyi.exe程序
4. 图像自动化：识别并点击指定图片，输入随机生成的邮箱

作者：AI助手
日期：2023年
"""

import os
import sys
import time
import json
import random
import string
import logging
import re
import subprocess
import shutil
from typing import List, Dict, Tuple, Optional, Union, Any
from pathlib import Path
import psutil
import pyautogui
import pyperclip
from PIL import Image
import keyboard

# 配置日志系统
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("anyi_automation.log"),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger("AnyiAutomation")


class ConfigManager:
    """
    配置管理类，负责读取和保存配置信息
    """
    def __init__(self, config_file: str = "anyi_config.json"):
        """
        初始化配置管理器
        
        Args:
            config_file: 配置文件路径
        """
        self.config_file = config_file
        self.config = self._load_config()
        
    def _load_config(self) -> Dict[str, Any]:
        """
        加载配置文件，如果不存在则创建默认配置
        
        Returns:
            配置字典
        """
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                return {"anyi_path": ""}
        except Exception as e:
            logger.error(f"加载配置文件失败: {e}")
            return {"anyi_path": ""}
    
    def save_config(self) -> bool:
        """
        保存配置到文件
        
        Returns:
            保存是否成功
        """
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
            return True
        except Exception as e:
            logger.error(f"保存配置文件失败: {e}")
            return False
    
    def get_anyi_path(self) -> str:
        """
        获取Anyi.exe的路径
        
        Returns:
            Anyi.exe的路径
        """
        return self.config.get("anyi_path", "")
    
    def set_anyi_path(self, path: str) -> None:
        """
        设置Anyi.exe的路径
        
        Args:
            path: Anyi.exe的路径
        """
        self.config["anyi_path"] = path
        self.save_config()


class ProcessManager:
    """
    进程管理类，负责检查和终止进程
    """
    @staticmethod
    def is_process_running(process_name: str) -> bool:
        """
        检查指定名称的进程是否正在运行
        
        Args:
            process_name: 进程名称
            
        Returns:
            进程是否正在运行
        """
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                if process_name.lower() in proc.info['name'].lower():
                    return True
            return False
        except Exception as e:
            logger.error(f"检查进程 {process_name} 失败: {e}")
            return False
    
    @staticmethod
    def terminate_process(process_name: str) -> bool:
        """
        终止指定名称的进程
        
        Args:
            process_name: 进程名称
            
        Returns:
            是否成功终止进程
        """
        try:
            terminated = False
            for proc in psutil.process_iter(['pid', 'name']):
                if process_name.lower() in proc.info['name'].lower():
                    try:
                        process = psutil.Process(proc.info['pid'])
                        process.terminate()
                        terminated = True
                        logger.info(f"已终止进程 {process_name} (PID: {proc.info['pid']})")
                    except psutil.AccessDenied:
                        logger.warning(f"无权限终止进程 {process_name} (PID: {proc.info['pid']})，尝试强制终止")
                        try:
                            process = psutil.Process(proc.info['pid'])
                            process.kill()
                            terminated = True
                            logger.info(f"已强制终止进程 {process_name} (PID: {proc.info['pid']})")
                        except Exception as e:
                            logger.error(f"强制终止进程 {process_name} 失败: {e}")
                    except Exception as e:
                        logger.error(f"终止进程 {process_name} 失败: {e}")
            return terminated
        except Exception as e:
            logger.error(f"终止进程 {process_name} 时发生错误: {e}")
            return False


class FileManager:
    """
    文件管理类，负责文件和文件夹操作
    """
    @staticmethod
    def get_current_username() -> str:
        """
        获取当前系统用户名
        
        Returns:
            当前用户名
        """
        try:
            return os.getlogin()
        except Exception as e:
            logger.error(f"获取当前用户名失败: {e}")
            try:
                import getpass
                return getpass.getuser()
            except Exception as e2:
                logger.error(f"使用备用方法获取用户名也失败: {e2}")
                return ""
    
    @staticmethod
    def delete_config_folder() -> bool:
        """
        删除当前用户的.config文件夹
        
        Returns:
            是否成功删除
        """
        try:
            username = FileManager.get_current_username()
            if not username:
                logger.error("无法获取用户名，无法删除.config文件夹")
                return False
            
            config_path = os.path.join("C:\\", "Users", username, ".config")
            anyi_config_path = os.path.join(config_path, "anyi")
            
            if os.path.exists(anyi_config_path):
                try:
                    # 使用异步方式删除文件夹
                    import asyncio
                    import concurrent.futures
                    
                    async def async_delete():
                        with concurrent.futures.ThreadPoolExecutor() as executor:
                            loop = asyncio.get_event_loop()
                            try:
                                # 尝试结束可能占用文件的进程
                                for proc in psutil.process_iter(['pid', 'name', 'open_files']):
                                    try:
                                        for file in proc.open_files():
                                            if anyi_config_path in file.path:
                                                proc.terminate()
                                                break
                                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                                        continue
                                
                                # 使用线程池执行删除操作
                                def delete_folder():
                                    import stat
                                    def on_rm_error(func, path, exc_info):
                                        try:
                                            os.chmod(path, stat.S_IWRITE)
                                            if os.path.isfile(path):
                                                os.unlink(path)
                                            elif os.path.isdir(path):
                                                os.rmdir(path)
                                        except Exception as e:
                                            logger.error(f"删除文件/文件夹失败: {path}, 错误: {e}")
                                    
                                    shutil.rmtree(anyi_config_path, onerror=on_rm_error)
                                
                                await loop.run_in_executor(executor, delete_folder)
                                return True
                            except Exception as e:
                                logger.error(f"异步删除文件夹失败: {e}")
                                return False
                    
                    # 运行异步删除
                    if asyncio.get_event_loop().is_closed():
                        loop = asyncio.new_event_loop()
                        asyncio.set_event_loop(loop)
                    result = asyncio.get_event_loop().run_until_complete(async_delete())
                    
                    if result:
                        logger.info(f"已删除Anyi配置文件夹: {anyi_config_path}")
                        return True
                    return False
                except Exception as e:
                    logger.error(f"删除Anyi配置文件夹失败: {e}")
                    return False
            else:
                logger.info(f"Anyi配置文件夹不存在: {anyi_config_path}")
                return True
        except Exception as e:
            logger.error(f"删除配置文件夹时发生错误: {e}")
            return False
    
    @staticmethod
    def validate_exe_path(path: str) -> bool:
        """
        验证可执行文件路径是否有效
        
        Args:
            path: 可执行文件路径
            
        Returns:
            路径是否有效
        """
        if not path:
            return False
        
        # 检查路径格式
        pattern = r'^[a-zA-Z]:\\(?:[^\\/:*?"<>|\r\n]+\\)*[^\\/:*?"<>|\r\n]*\.exe$'
        if not re.match(pattern, path):
            return False
        
        # 检查文件是否存在
        if not os.path.exists(path):
            return False
        
        # 检查是否为文件
        if not os.path.isfile(path):
            return False
        
        return True


class ProgramRunner:
    """
    程序运行类，负责运行Anyi.exe程序
    """
    def __init__(self, config_manager: ConfigManager):
        """
        初始化程序运行器
        
        Args:
            config_manager: 配置管理器实例
        """
        self.config_manager = config_manager
    
    def run_anyi(self) -> bool:
        """
        运行Anyi.exe程序
        
        Returns:
            是否成功运行
        """
        anyi_path = self.config_manager.get_anyi_path()
        
        # 如果没有保存的路径，提示用户输入
        if not anyi_path or not FileManager.validate_exe_path(anyi_path):
            return self._prompt_and_run_anyi()
        else:
            return self._run_anyi_exe(anyi_path)
    
    def _prompt_and_run_anyi(self) -> bool:
        """
        提示用户输入Anyi.exe路径并运行
        
        Returns:
            是否成功运行
        """
        retry_count = 0
        max_retries = 3
        
        while retry_count < max_retries:
            try:
                print("\n请输入Anyi.exe的完整路径:")
                anyi_path = input().strip()
                
                if FileManager.validate_exe_path(anyi_path):
                    if self._run_anyi_exe(anyi_path):
                        # 保存有效路径
                        self.config_manager.set_anyi_path(anyi_path)
                        return True
                else:
                    print("无效的路径，请确保输入正确的Anyi.exe完整路径")
            except Exception as e:
                logger.error(f"获取Anyi.exe路径时发生错误: {e}")
            
            retry_count += 1
            if retry_count < max_retries:
                print(f"请重新输入，还有{max_retries - retry_count}次尝试机会")
        
        logger.error("多次尝试运行Anyi.exe失败")
        return False
    
    def _run_anyi_exe(self, path: str) -> bool:
        """
        运行指定路径的Anyi.exe
        
        Args:
            path: Anyi.exe的路径
            
        Returns:
            是否成功运行
        """
        try:
            logger.info(f"尝试运行Anyi.exe: {path}")
            process = subprocess.Popen([path])
            
            # 等待一段时间检查进程是否仍在运行
            time.sleep(2)
            if process.poll() is None:  # 如果进程仍在运行
                logger.info(f"成功启动Anyi.exe: {path}")
                return True
            else:
                logger.error(f"Anyi.exe启动后立即退出: {path}")
                return False
        except Exception as e:
            logger.error(f"运行Anyi.exe失败: {e}")
            return False


class ImageAutomation:
    """
    图像自动化类，负责图像识别和自动点击
    """
    def __init__(self, images_dir: str = "images"):
        """
        初始化图像自动化
        
        Args:
            images_dir: 图像文件夹路径
        """
        self.images_dir = images_dir
        self.confidence = 0.7  # 降低图像匹配置信度以提高识别率
        self.max_wait_time = 30  # 最大等待时间（秒）
        self.check_interval = 0.5  # 增加检查间隔以确保界面加载完成
        self.retry_count = 3  # 单次操作的重试次数
        
        # 设置pyautogui安全设置
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.05  # 进一步降低默认暂停时间以加快操作速度
        
        # 初始化随机数生成器
        random.seed()
    
    def _get_image_path(self, image_name: str) -> str:
        """
        获取图像文件的完整路径
        
        Args:
            image_name: 图像文件名
            
        Returns:
            图像文件的完整路径
        """
        return os.path.join(self.images_dir, image_name)
    
    def wait_and_find_image(self, image_name: str, timeout: Optional[int] = None, region: Optional[Tuple[int, int, int, int]] = None) -> Optional[Tuple[int, int]]:
        """
        等待并查找图像
        
        Args:
            image_name: 图像文件名
            timeout: 超时时间（秒），None表示使用默认值
            region: 搜索区域 (left, top, width, height)，None表示全屏搜索
            
        Returns:
            图像中心坐标，未找到则返回None
        """
        if timeout is None:
            timeout = self.max_wait_time
            
        image_path = self._get_image_path(image_name)
        if not os.path.exists(image_path):
            logger.error(f"图像文件不存在: {image_path}")
            return None
        
        logger.info(f"等待图像出现: {image_name}")
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            for retry in range(self.retry_count):
                try:
                    # 每次重试前增加短暂延迟，确保界面完全加载
                    if retry > 0:
                        time.sleep(0.5)
                        logger.info(f"正在进行第{retry + 1}次图像识别尝试: {image_name}")
                    
                    location = pyautogui.locateCenterOnScreen(
                        image_path, 
                        confidence=self.confidence,
                        region=region
                    )
                    if location:
                        logger.info(f"找到图像 {image_name} 在位置 {location}")
                        return location
                except Exception as e:
                    if retry == self.retry_count - 1:
                        logger.error(f"查找图像 {image_name} 时发生错误: {e}")
                    else:
                        logger.warning(f"第{retry + 1}次查找图像 {image_name} 失败: {e}，准备重试")
            
            # 检查是否有7.png出现（全局监控）
            if image_name != "7.png":
                try:
                    popup_location = pyautogui.locateCenterOnScreen(
                        self._get_image_path("7.png"),
                        confidence=self.confidence
                    )
                    if popup_location:
                        logger.info(f"检测到弹窗 7.png 在位置 {popup_location}")
                        self.click_image_with_offset("7.png")
                except Exception:
                    pass  # 忽略7.png查找错误
            
            time.sleep(self.check_interval)
        
        logger.warning(f"等待图像 {image_name} 超时")
        return None
    
    def click_image_with_offset(self, image_name: str, timeout: Optional[int] = None, region: Optional[Tuple[int, int, int, int]] = None) -> bool:
        """
        查找图像并点击（带随机偏移）
        
        Args:
            image_name: 图像文件名
            timeout: 超时时间（秒），None表示使用默认值
            region: 搜索区域，None表示全屏搜索
            
        Returns:
            是否成功点击
        """
        for attempt in range(self.retry_count):
            if attempt > 0:
                logger.warning(f"第{attempt + 1}次尝试点击图像 {image_name}")
                time.sleep(1)  # 重试前等待界面稳定
                
            location = self.wait_and_find_image(image_name, timeout, region)
            if location:
                x, y = location
                try:
                    # 添加小的随机偏移，模拟人类点击
                    offset_x = random.randint(-5, 5)
                    offset_y = random.randint(-5, 5)
                    pyautogui.click(x + offset_x, y + offset_y)
                    logger.info(f"成功点击图像 {image_name} 在位置 ({x + offset_x}, {y + offset_y})")
                    return True
                except Exception as e:
                    logger.error(f"点击图像 {image_name} 时发生错误: {e}")
                    if attempt == self.retry_count - 1:
                        return False
                    continue
        
        logger.error(f"无法找到并点击图像 {image_name}")
        return False
    
    def ensure_english_input(self) -> None:
        """
        确保系统处于英文输入状态
        """
        try:
            # 按下Shift键切换到英文输入法
            keyboard.press_and_release('shift')
            time.sleep(0.2)
            # 再次按下以确保
            keyboard.press_and_release('shift')
            logger.info("已切换到英文输入法")
        except Exception as e:
            logger.error(f"切换输入法失败: {e}")
    
    def type_text_safely(self, text: str) -> None:
        """
        安全地输入文本（处理输入法问题）
        
        Args:
            text: 要输入的文本
        """
        try:
            # 确保英文输入状态
            self.ensure_english_input()
            
            # 使用剪贴板输入，避免输入法问题
            pyperclip.copy(text)
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'v')
            logger.info(f"已输入文本: {text}")
        except Exception as e:
            logger.error(f"输入文本失败: {e}")
            # 备用方法：直接输入
            try:
                pyautogui.write(text, interval=0.1)
                logger.info(f"使用备用方法输入文本: {text}")
            except Exception as e2:
                logger.error(f"备用输入方法也失败: {e2}")
    
    def generate_random_email(self) -> str:
        """
        生成随机邮箱地址
        
        Returns:
            随机生成的邮箱地址
        """
        # 生成6位随机字符（字母和数字）
        chars = string.ascii_letters + string.digits
        random_str = ''.join(random.choice(chars) for _ in range(6))
        email = f"{random_str}@163.com"
        logger.info(f"生成随机邮箱: {email}")
        return email
    
    def perform_automation_sequence(self) -> bool:
        """
        执行自动化操作序列
        
        Returns:
            是否成功完成所有操作
        """
        try:
            # 点击图像1-3
            for i in range(1, 4):
                image_name = f"{i}.png"
                if not self.click_image_with_offset(image_name):
                    logger.error(f"无法找到并点击图像 {image_name}")
                    return False
                time.sleep(random.uniform(0.3, 0.6))  # 缩短随机等待时间以加快操作速度
            
            # 生成随机邮箱并输入
            email = self.generate_random_email()
            self.type_text_safely(email)
            time.sleep(random.uniform(0.3, 0.6))  # 缩短随机等待时间
            
            # 点击图像4并输入邮箱
            if not self.click_image_with_offset("4.png"):
                logger.error("无法找到并点击图像 4.png")
                return False
            time.sleep(random.uniform(0.3, 0.6))  # 缩短随机等待时间
            self.type_text_safely(email)
            time.sleep(random.uniform(0.3, 0.6))  # 缩短随机等待时间
            
            # 点击图像5
            if not self.click_image_with_offset("5.png"):
                logger.error("无法找到并点击图像 5.png")
                return False
            time.sleep(random.uniform(0.3, 0.6))  # 缩短随机等待时间

            # 点击图像1
            if not self.click_image_with_offset("1.png"):
                logger.error("无法找到并点击图像 1.png")
                return False
            time.sleep(random.uniform(0.3, 0.6))  # 缩短随机等待时间

            # 点击图像6
            if not self.click_image_with_offset("6.png"):
                logger.error("无法找到并点击图像 6.png")
                return False
            time.sleep(random.uniform(0.3, 0.6))  # 缩短随机等待时间
            
            logger.info("成功完成自动化操作序列")
            return True
        except Exception as e:
            logger.error(f"执行自动化操作序列时发生错误: {e}")
            return False


class AnyiAutomation:
    """
    Anyi自动化主类，协调各个模块完成自动化任务
    """
    def __init__(self):
        """
        初始化Anyi自动化
        """
        self.config_manager = ConfigManager()
        self.process_manager = ProcessManager()
        self.file_manager = FileManager()
        self.program_runner = ProgramRunner(self.config_manager)
        self.image_automation = ImageAutomation()
    
    def run(self) -> None:
        """
        运行自动化流程
        """
        try:
            logger.info("开始Anyi自动化流程")
            
            # 1. 结束特定进程
            logger.info("步骤1: 结束特定进程")
            for process_name in ["Anyi.exe", "anyi-core.exe"]:
                if self.process_manager.is_process_running(process_name):
                    logger.info(f"发现进程 {process_name} 正在运行，尝试终止")
                    if self.process_manager.terminate_process(process_name):
                        logger.info(f"成功终止进程 {process_name}")
                    else:
                        logger.warning(f"无法终止进程 {process_name}")
                else:
                    logger.info(f"进程 {process_name} 未运行")
            
            # 2. 删除配置文件夹
            logger.info("步骤2: 删除配置文件夹")
            if self.file_manager.delete_config_folder():
                logger.info("成功删除配置文件夹")
            else:
                logger.warning("删除配置文件夹失败")
            
            # 3. 运行用户指定的程序
            logger.info("步骤3: 运行Anyi.exe")
            if not self.program_runner.run_anyi():
                logger.error("无法运行Anyi.exe，自动化流程终止")
                return
            
            # 等待程序启动并进行稳定性检查
            logger.info("等待Anyi.exe启动...")
            start_time = time.time()
            max_wait_time = 30  # 最大等待时间（秒）
            max_retries = 3  # 最大重试次数
            retry_count = 0
            
            while retry_count < max_retries:
                try:
                    # 动态检测程序窗口
                    window_found = False
                    retry_start_time = time.time()
                    
                    while time.time() - retry_start_time < max_wait_time:
                        try:
                            # 尝试查找第一个图标，作为程序启动成功的标志
                            if self.image_automation.wait_and_find_image("1.png", timeout=1):
                                logger.info("检测到程序窗口已启动")
                                window_found = True
                                break
                        except Exception:
                            pass
                        time.sleep(0.3)
                    
                    if not window_found:
                        raise Exception("程序窗口未出现")
                    
                    # 4. 执行图像自动化操作
                    logger.info("步骤4: 执行图像自动化操作")
                    if self.image_automation.perform_automation_sequence():
                        logger.info("成功完成所有自动化操作")
                        break
                    else:
                        raise Exception("自动化操作失败")
                        
                except Exception as e:
                    retry_count += 1
                    if retry_count < max_retries:
                        logger.warning(f"第{retry_count}次尝试失败: {e}，正在重试...")
                        # 终止当前进程并重新启动
                        self.process_manager.terminate_process("Anyi.exe")
                        self.process_manager.terminate_process("anyi-core.exe")
                        time.sleep(1)
                        if not self.program_runner.run_anyi():
                            logger.error("重新启动Anyi.exe失败")
                            break
                    else:
                        logger.error(f"达到最大重试次数，自动化操作失败: {e}")
                        break
        
        except Exception as e:
            logger.error(f"自动化流程执行过程中发生错误: {e}")
        
        finally:
            logger.info("Anyi自动化流程结束")


def main():
    """
    主函数
    """
    try:
        automation = AnyiAutomation()
        automation.run()
    except Exception as e:
        logger.critical(f"程序执行过程中发生严重错误: {e}")
        print(f"程序执行过程中发生错误: {e}")
    
    input("按Enter键退出...")


if __name__ == "__main__":
    main()