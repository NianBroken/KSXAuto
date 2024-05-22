# 导入必要的模块
import time  # 引入时间模块，用于处理时间相关操作
import configparser  # 引入配置解析器模块，用于读取配置文件
import sys  # 引入系统模块，处理系统相关参数和函数
import os  # 引入操作系统接口模块，处理文件和目录相关操作
import ctypes  # 引入ctypes模块，提供与C语言兼容的数据类型
import json  # 引入json模块，用于处理JSON数据
import requests  # 引入requests模块，用于HTTP请求
import glob  # 引入glob模块，用于文件路径模式匹配
from selenium import webdriver  # 从selenium库中引入webdriver，用于自动化操作浏览器
from selenium.webdriver.common.by import By  # 引入By类，用于定位页面元素
from selenium.common.exceptions import WebDriverException  # 引入WebDriverException类，用于处理浏览器异常
from datetime import datetime  # 从datetime模块中引入datetime类，用于处理日期和时间


def set_console_title(title):
    """设置控制台标题"""
    ctypes.windll.kernel32.SetConsoleTitleW(title)  # 调用Windows API设置控制台窗口标题


def load_config(file_path):
    """加载配置文件"""
    config = configparser.ConfigParser()  # 创建配置解析器对象
    config.read(file_path, encoding="utf-8")  # 读取指定路径的配置文件，使用UTF-8编码
    return config  # 返回配置解析器对象


def init_browser():
    """初始化浏览器配置并返回浏览器实例"""
    options = webdriver.EdgeOptions()  # 创建Edge浏览器选项对象
    options.add_experimental_option("detach", True)  # 设置浏览器启动后不关闭
    options.add_experimental_option("excludeSwitches", ["enable-logging"])  # 排除特定日志开关
    options.add_argument("--start-maximized")  # 设置浏览器启动时最大化窗口
    driver = webdriver.Edge(options=options)  # 创建Edge浏览器实例，应用配置选项
    driver.implicitly_wait(30)  # 设置隐式等待时间为30秒
    return driver  # 返回浏览器实例


def check_for_updates(current_path):
    global local_data, remote_data  # 声明全局变量，用于存储本地和远程版本信息数据
    """检查软件更新"""
    update_url = "https://gitee.com/nianbroken/KSXAuto/raw/main/version_info.json"  # 更新信息URL
    response = requests.get(update_url)  # 发起GET请求获取更新信息

    if response.status_code == 200:  # 如果请求成功
        remote_data = response.json()  # 解析响应内容为JSON格式
        remote_version = remote_data["software_info"]["version_id"]  # 获取远程版本号
        remote_changelog = remote_data["software_info"]["changelog"]  # 获取远程更新日志
        local_file_path = os.path.join(current_path, "version_info.json")  # 本地版本信息文件路径
        if os.path.exists(local_file_path):  # 如果本地版本信息文件存在
            with open(local_file_path, "r", encoding="utf-8") as local_file:  # 打开本地文件
                local_data = json.load(local_file)  # 解析本地文件内容为JSON格式
                local_version = local_data["software_info"]["version_id"]  # 获取本地版本号
                if local_version < remote_version:  # 如果本地版本号小于远程版本号
                    print(f"检测到新版本 {remote_data['software_info']['version_number']} 已发布")  # 提示新版本可用
                    print(f"新版本发布时间：{remote_data['software_info']['release_date']}")  # 提示新版本发布时间
                    print("更新日志如下：")  # 提示更新日志
                    for changelog in remote_changelog:  # 遍历远程更新日志列表
                        print(changelog)  # 输出每个更新日志条目
                    print(f"下载地址：{remote_data['software_info']['download_url']}")  # 提示下载地址
        else:
            print(f"检测到新版本 {remote_data['software_info']['version_number']} 已发布")  # 如果本地文件不存在，提示新版本可用
            print(f"下载地址：{remote_data['software_info']['download_url']}")  # 提示下载地址
    else:
        print(f"无法获取更新信息：{response.status_code}")  # 如果请求失败，提示错误码


def get_current_time():
    """返回当前时间"""
    return datetime.now().strftime("%m-%d %H:%M:%S")  # 获取当前时间并格式化为字符串


def login(driver, username, password):
    """执行用户登录操作"""
    if username:  # 如果用户名存在
        username_input = driver.find_element(By.XPATH, '//*[@id="username"]')  # 定位用户名输入框
        username_input.clear()  # 清空输入框内容
        username_input.send_keys(username)  # 输入用户名
        print(f"[{get_current_time()}] 被动填写账号")  # 输出日志信息

    if password:  # 如果密码存在
        password_input = driver.find_element(By.XPATH, '//*[@id="userTypePwd"]')  # 定位密码输入框
        password_input.clear()  # 清空输入框内容
        password_input.send_keys(password)  # 输入密码
        print(f"[{get_current_time()}] 被动填写密码")  # 输出日志信息

    if username and password:  # 如果用户名和密码都存在
        driver.find_element(By.XPATH, '//*[@id="loginBtn"]').click()  # 点击登录按钮
        print(f"[{get_current_time()}] 被动点击登录按钮")  # 输出日志信息


def get_user_full_name(driver):
    """获取用户的完整姓名"""
    return driver.find_element(By.XPATH, '//*[@id="viewFrameWorkBody"]/div/div[1]/div[1]/div/div/div[2]').text  # 定位并返回用户姓名


total_videos_already_printed = False  # 初始化已打印标志为False


def get_total_videos(driver):
    """获取课程视频总数"""
    global total_videos_already_printed  # 声明为全局变量
    total_videos = len(driver.find_elements(By.CLASS_NAME, "catalog-item"))  # 定位所有视频元素并返回总数
    if not total_videos_already_printed:  # 如果未打印过视频总数
        print(f"[{get_current_time()}] 视频总数：{total_videos}")  # 输出视频总数
        total_videos_already_printed = True  # 标记已打印过视频总数
    return total_videos  # 返回视频总数


unfinished_videos_already_printed = False  # 初始化已打印标志为False


def get_unfinished_videos(driver, total_videos):
    """获取未完成的视频列表"""
    global unfinished_videos_already_printed  # 声明为全局变量
    unfinished_videos = []  # 初始化未完成视频列表

    for i in range(1, total_videos + 1):  # 遍历所有视频
        completion_xpath = f'//*[@id="viewFrameWorkBody"]/div/div[3]/div[2]/div/div[2]/div[{i}]/div[2]/div[2]/div[2]'  # 生成完成状态的XPath
        completion_text = driver.find_element(By.XPATH, completion_xpath).text  # 获取完成状态文本
        if completion_text != "100%":  # 如果视频未完成
            unfinished_videos.append(i)  # 将视频ID添加到未完成列表中

    if unfinished_videos:  # 如果存在未完成视频
        if not unfinished_videos_already_printed:  # 如果未打印过未完成视频列表
            print(f"[{get_current_time()}] 未完成视频个数：{len(unfinished_videos)}")  # 输出未完成视频列表
            print(f"[{get_current_time()}] 未完成视频列表：{unfinished_videos}")  # 输出未完成视频列表
            unfinished_videos_already_printed = True  # 标记已打印过未完成视频列表

    return unfinished_videos  # 返回未完成视频列表


def get_first_unfinished_video(driver, video_id):
    """获取第一个未完成视频的元素及标题"""
    video_xpath = f'//*[@id="viewFrameWorkBody"]/div/div[3]/div[2]/div/div[2]/div[{video_id}]/div[2]/div[1]'  # 生成视频元素的XPath
    video_element = driver.find_element(By.XPATH, video_xpath)  # 定位视频元素
    return video_element, video_element.text  # 返回视频元素及其标题


def play_video(driver, video_element, mute, playback_rate):
    """播放指定视频"""
    video_element.click()  # 点击视频元素以播放视频
    video_elements = driver.find_elements(By.TAG_NAME, "video")  # 定位页面上的所有视频元素

    if mute and video_elements:  # 如果需要静音并且存在视频元素
        for video in video_elements:  # 遍历所有视频元素
            driver.execute_script("arguments[0].muted = true;", video)  # 静音视频
            driver.execute_script(f"arguments[0].playbackRate = {playback_rate};", video)  # 设置视频播放速度

    if driver.find_element(By.XPATH, '//*[@id="myVideo"]/button').is_displayed():  # 如果播放按钮可见
        driver.find_element(By.XPATH, '//*[@id="myVideo"]/button').click()  # 点击播放按钮
        return True  # 返回播放成功
    else:
        return False  # 返回播放失败


def get_video_remaining_time(driver):
    """获取当前视频剩余时间"""
    try:
        remaining_time_element = driver.find_element(By.XPATH, '//*[@id="myVideo"]/div[1]/div[2]')  # 尝试定位剩余时间元素
        if remaining_time_element.is_displayed():  # 如果剩余时间元素可见
            return remaining_time_element.text  # 返回剩余时间文本
        else:  # 如果剩余时间元素不可见
            return None  # 返回None
    except Exception:  # 如果发生异常
        return None  # 返回None


def get_alternate_video_remaining_time(driver):
    """获取当前视频备用剩余时间"""
    try:
        alternate_remaining_time_element = driver.find_element(By.XPATH, '//*[@id="myVideo"]/div[5]/div[5]/span[3]')  # 尝试定位备用剩余时间元素
        return alternate_remaining_time_element.text  # 返回备用剩余时间文本
    except Exception:  # 如果发生异常
        return None  # 返回None


def check_browser_closed(driver):
    """检查浏览器是否被关闭"""
    try:
        driver.current_url  # 尝试获取当前窗口句柄
        return False  # 如果成功获取，则浏览器未被关闭
    except WebDriverException:  # 如果发生WebDriverException异常
        return True  # 则浏览器被关闭


def process_course_page(driver, course, mute, playback_rate):
    """处理课程页面"""
    while True:  # 无限循环处理课程页面
        try:
            if "/exam/pc/course/#/study?courseId" not in driver.current_url:  # 如果当前URL不包含课程页面标识
                print(f"[{get_current_time()}] 等待进入课程页面")  # 输出日志信息
                if course:  # 如果课程URL存在
                    driver.get(course)  # 打开课程页面
                    print(f"[{get_current_time()}] 被动进入课程页面")  # 输出日志信息
                    time.sleep(2)  # 等待2秒
                else:
                    while True:  # 等待页面加载到课程页面
                        if "/exam/pc/course/#/study?courseId" in driver.current_url:  # 如果当前URL包含课程页面标识
                            print(f"[{get_current_time()}] 主动进入课程页面")  # 输出日志信息
                            time.sleep(1)  # 每秒检查一次
                            break  # 跳出循环

            total_videos = get_total_videos(driver)  # 获取课程视频总数

            while True:  # 无限循环处理未完成视频
                unfinished_videos = get_unfinished_videos(driver, total_videos)  # 获取未完成视频列表
                if not unfinished_videos:  # 如果没有未完成视频
                    print(f"[{get_current_time()}] 不存在未完成的视频")  # 输出日志信息
                    sys.exit()  # 退出程序

                first_unfinished_video_id = unfinished_videos[0]  # 获取第一个未完成视频ID
                first_unfinished_video, first_unfinished_video_title = get_first_unfinished_video(driver, first_unfinished_video_id)  # 获取第一个未完成视频元素及标题

                if play_video(driver, first_unfinished_video, mute, playback_rate):  # 播放第一个未完成视频
                    print(f"[{get_current_time()}] 正在播放第 {first_unfinished_video_id} 个视频：{first_unfinished_video_title}")  # 输出播放日志

                while True:  # 无限循环检查视频播放状态
                    remaining_time = get_video_remaining_time(driver)  # 获取当前视频剩余时间
                    alternate_remaining_time = get_alternate_video_remaining_time(driver)  # 获取当前视频备用剩余时间
                    try:
                        play_button_visible = driver.find_element(By.XPATH, '//*[@id="myVideo"]/button').is_displayed()  # 如果播放按钮可见
                    except Exception:
                        play_button_visible = False
                    if not remaining_time or not alternate_remaining_time or play_button_visible or remaining_time == "00:00:00" or alternate_remaining_time == "0:00":  # 如果视频已播放完毕
                        print(f"当前视频剩余：{remaining_time}", end="\r", flush=True)  # 输出剩余时间，覆盖同一行
                        break  # 跳出循环

        except Exception as e:  # 如果发生异常
            time.sleep(5)  # 等待2秒
            if check_browser_closed(driver):  # 如果浏览器被关闭
                print(f"[{get_current_time()}] 异常退出")  # 输出日志信息
                write_log_to_file(e)  # 将异常写入日志文件
                print(f"[{get_current_time()}] 错误内容已写入Log文件")  # 输出错误信息
                sys.exit()  # 退出程序
            else:
                continue  # 继续循环


def write_log_to_file(log):
    """将日志写入文件"""
    log_str = str(log)  # 将日志转换为字符串

    log_dir = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), "log")  # 获取日志文件夹路径
    if not os.path.exists(log_dir):  # 创建一个名为log的文件夹（如果不存在）
        os.makedirs(log_dir)  # 创建文件夹

    now = datetime.now().strftime("%Y-%m-%d %H-%M-%S-%f")  # 获取当前日期和时间（精确到毫秒）
    log_file_name = os.path.join(log_dir, f"log[{now}].txt")  # 构建日志文件名

    with open(log_file_name, "w") as f:  # 写入日志到文件
        f.write(f"{now}\n{log_str}")  # 写入日期和时间以及日志内容

    log_files = sorted(glob.glob(os.path.join(log_dir, "*.txt")), key=os.path.getctime)  # 检查日志文件数量
    if len(log_files) > 100:  # 如果日志文件数量超过100个
        while len(log_files) > 90:  # 删除最旧的日志，直到日志数量减少到90条
            os.remove(log_files[0])  # 删除最旧的日志文件
            log_files.pop(0)  # 更新日志文件列表


def main():
    try:
        print("------")  # 分隔线
        while True:  # 无限循环，直到用户按下回车键退出程序
            user_input = input("按下回车后将为你打开浏览器窗口")  # 提示用户按下回车键
            if user_input == "":  # 如果用户按下回车键
                break  # 跳出循环
            else:  # 如果用户输入了其他内容
                print("无需输入内容，直接按下回车键即可")  # 提示用户无需输入内容，直接按下回车键即可

        current_path = os.path.dirname(os.path.realpath(sys.argv[0]))  # 获取当前脚本的目录路径
        config_file_path = os.path.join(current_path, "config.ini")  # 配置文件路径
        config = load_config(config_file_path)  # 加载配置文件
        mute = config.getboolean("videoConfig", "mute", fallback=True)  # 获取静音配置，默认静音
        playback_rate = config.getfloat("videoConfig", "playbackRate", fallback=1.0)  # 获取视频播放速度，默认为1.0倍
        url = config.get("loginConfig", "url")  # 获取登录URL
        username = config.get("loginConfig", "username")  # 获取用户名
        password = config.get("loginConfig", "password")  # 获取密码
        course = config.get("loginConfig", "course")  # 获取课程URL

        set_console_title("KSXAuto")  # 设置控制台标题

        driver = init_browser()  # 初始化浏览器

        check_for_updates(current_path)  # 检查软件更新
        print("------")  # 分隔线
        print(f"静音：{mute}")  # 输出静音配置
        print(f"倍数：{playback_rate}")  # 输出播放速度配置
        print(f"运行路径：{current_path}")  # 输出当前运行路径
        print("------")  # 分隔线
        print("Copyright © 2024 NianBroken. All rights reserved.")  # 版权信息
        print("Github：https://github.com/NianBroken/KSXAuto/")  # Github地址
        print(f"[{get_current_time()}] KSXAuto {local_data['software_info']['version_number']} 程序启动成功")  # 输出启动成功信息
        print(f"[{get_current_time()}] 等待进入登录页面")  # 输出等待登录页面信息

        if url:  # 如果登录URL存在
            driver.get(url)  # 打开登录页面
            print(f"[{get_current_time()}] 被动进入登录页面")  # 输出被动进入登录页面信息
        else:
            while True:  # 无限循环等待登录页面
                current_page_url = driver.current_url  # 获取当前页面URL
                if "login/account/login" in current_page_url:  # 如果当前页面是登录页面
                    print(f"[{get_current_time()}] 主动进入登录页面")  # 输出主动进入登录页面信息
                    break  # 跳出循环
                time.sleep(1)  # 每秒检查一次

        login(driver, username, password)  # 执行登录操作
        time.sleep(1)  # 等待1秒

        current_page_url = driver.current_url  # 获取当前页面URL
        if "login/account/login" in current_page_url:  # 如果仍然在登录页面
            login_errormsg = driver.find_element(By.XPATH, '//*[@id="errormsg"]')  # 定位错误信息元素
            if login_errormsg.is_displayed() and "错误" in login_errormsg.text:  # 如果错误信息可见且包含"错误"
                print(f"[{get_current_time()}] {login_errormsg.text} 已切换主动模式 请手动输入账号密码")  # 输出错误信息
                print(f"[{get_current_time()}] 已切换主动模式 请主动输入账号密码")  # # 输出日志信息

        print(f"[{get_current_time()}] 等待进入主页")  # 输出等待进入主页信息

        while True:  # 无限循环等待主页
            current_page_url = driver.current_url  # 获取当前页面URL
            if "/exam/pc/home/" in current_page_url:  # 如果当前页面是主页
                print(f"[{get_current_time()}] 主动进入首页")  # 输出主动进入首页信息
                print(f"[{get_current_time()}] 姓名：{get_user_full_name(driver)}")  # 输出用户姓名
                time.sleep(2)  # 等待2秒
                break  # 跳出循环

        while True:  # 无限循环处理课程页面
            try:
                process_course_page(driver, course, mute, playback_rate)  # 处理课程页面
            except Exception:  # 如果发生异常
                continue  # 继续循环
    except Exception as e:  # 如果发生异常
        write_log_to_file(e)  # 将错误信息写入日志文件
        print(f"[{get_current_time()}] 发生错误 错误内容已写入Log文件")  # 输出错误信息
        sys.exit()  # 退出程序


if __name__ == "__main__":  # 如果脚本作为主程序运行
    main()  # 执行主函数
