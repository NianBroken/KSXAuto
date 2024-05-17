# 导入必要的模块
import time  # 引入时间模块，用于处理时间相关操作
import configparser  # 引入配置解析器模块，用于读取配置文件
import sys  # 引入系统模块，处理系统相关参数和函数
import os  # 引入操作系统接口模块，处理文件和目录相关操作
import ctypes  # 引入ctypes模块，提供与C语言兼容的数据类型
import json  # 引入json模块，用于处理JSON数据
import requests  # 引入requests模块，用于HTTP请求
from selenium import webdriver  # 从selenium库中引入webdriver，用于自动化操作浏览器
from selenium.webdriver.common.by import By  # 引入By类，用于定位页面元素
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
    driver.implicitly_wait(2)  # 设置隐式等待时间为2秒
    return driver  # 返回浏览器实例


def check_for_updates(current_path):
    """检查软件更新"""
    update_url = "https://gitee.com/nianbroken/KSXAuto/raw/main/version_info.json"  # 更新信息URL
    response = requests.get(update_url)  # 发起GET请求获取更新信息

    if response.status_code == 200:  # 如果请求成功
        remote_data = response.json()  # 解析响应内容为JSON格式
        remote_version = remote_data["software_info"]["version_id"]  # 获取远程版本号
        local_file_path = os.path.join(current_path, "version_info.json")  # 本地版本信息文件路径

        if os.path.exists(local_file_path):  # 如果本地版本信息文件存在
            with open(local_file_path, "r", encoding="utf-8") as local_file:  # 打开本地文件
                local_data = json.load(local_file)  # 解析本地文件内容为JSON格式
                local_version = local_data["software_info"]["version_id"]  # 获取本地版本号
                if local_version < remote_version:  # 如果本地版本号小于远程版本号
                    print(f"新版本可用：{remote_data['software_info']['download_url']}")  # 提示新版本可用
        else:
            print(f"新版本可用：{remote_data['software_info']['download_url']}")  # 如果本地文件不存在，提示新版本可用
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
        print(f"[{get_current_time()}] 输入账号")  # 输出日志信息

    if password:  # 如果密码存在
        password_input = driver.find_element(By.XPATH, '//*[@id="userTypePwd"]')  # 定位密码输入框
        password_input.clear()  # 清空输入框内容
        password_input.send_keys(password)  # 输入密码
        print(f"[{get_current_time()}] 输入密码")  # 输出日志信息

    if username and password:  # 如果用户名和密码都存在
        driver.find_element(By.XPATH, '//*[@id="loginBtn"]').click()  # 点击登录按钮
        print(f"[{get_current_time()}] 点击登录按钮")  # 输出日志信息


def get_user_full_name(driver):
    """获取用户的完整姓名"""
    return driver.find_element(By.XPATH, '//*[@id="viewFrameWorkBody"]/div/div[1]/div[1]/div/div/div[2]').text  # 定位并返回用户姓名


def get_total_videos(driver):
    """获取课程视频总数"""
    return len(driver.find_elements(By.CLASS_NAME, "catalog-item"))  # 定位所有视频元素并返回总数


def get_unfinished_videos(driver, total_videos):
    """获取未完成的视频列表"""
    unfinished_videos = []  # 初始化未完成视频列表

    for i in range(1, total_videos + 1):  # 遍历所有视频
        completion_xpath = f'//*[@id="viewFrameWorkBody"]/div/div[3]/div[2]/div/div[2]/div[{i}]/div[2]/div[2]/div[2]'  # 生成完成状态的XPath
        completion_text = driver.find_element(By.XPATH, completion_xpath).text  # 获取完成状态文本
        if completion_text != "100%":  # 如果视频未完成
            unfinished_videos.append(i)  # 将视频ID添加到未完成列表中

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
    return False  # 返回播放失败


def get_video_remaining_time(driver):
    """获取当前视频剩余时间"""
    try:
        remaining_time_element = driver.find_element(By.XPATH, '//*[@id="myVideo"]/div[1]/div[2]')  # 尝试定位剩余时间元素
        if remaining_time_element.is_displayed():  # 如果剩余时间元素可见
            return remaining_time_element.text  # 返回剩余时间文本
        return driver.find_element(By.XPATH, '//*[@id="myVideo"]/div[5]/div[5]/span[3]').text  # 返回备用路径的剩余时间文本
    except Exception:  # 如果发生异常
        return None  # 返回None


def process_course_page(driver, course, mute, playback_rate):
    """处理课程页面"""
    while True:  # 无限循环处理课程页面
        try:
            if course:  # 如果课程URL存在
                driver.get(course)  # 打开课程页面
                print(f"[{get_current_time()}] 进入课程页面")  # 输出日志信息
                time.sleep(2)  # 等待2秒
            else:
                while "/exam/pc/course/#/study?courseId" not in driver.current_url:  # 等待页面加载到课程页面
                    time.sleep(1)  # 每秒检查一次

            total_videos = get_total_videos(driver)  # 获取课程视频总数
            print(f"视频总数：{total_videos}")  # 输出视频总数

            while True:  # 无限循环处理未完成视频
                unfinished_videos = get_unfinished_videos(driver, total_videos)  # 获取未完成视频列表
                if not unfinished_videos:  # 如果没有未完成视频
                    print(f"[{get_current_time()}] 不存在未完成的视频")  # 输出日志信息
                    break  # 跳出循环

                first_unfinished_video_id = unfinished_videos[0]  # 获取第一个未完成视频ID
                first_unfinished_video, first_unfinished_video_title = get_first_unfinished_video(driver, first_unfinished_video_id)  # 获取第一个未完成视频元素及标题

                if play_video(driver, first_unfinished_video, mute, playback_rate):  # 播放第一个未完成视频
                    print(f"[{get_current_time()}] 正在播放第 {first_unfinished_video_id} 个视频：{first_unfinished_video_title}")  # 输出播放日志

                while True:  # 无限循环检查视频播放状态
                    remaining_time = get_video_remaining_time(driver)  # 获取当前视频剩余时间
                    if not remaining_time or remaining_time in ["00:00:00", "0:00"]:  # 如果视频已播放完毕
                        break  # 跳出循环

                    print(f"当前视频剩余：{remaining_time}", end="\r", flush=True)  # 输出剩余时间，覆盖同一行
        except Exception as e:  # 如果发生异常
            print(f"[{get_current_time()}] 发生错误：{str(e)}")  # 输出错误信息
            continue  # 继续循环


def main():
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
    print(f"[{get_current_time()}] 程序启动成功")  # 输出启动成功信息
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
            print(f"[{get_current_time()}] {login_errormsg.text} 已切换主动模式")  # 输出错误信息并切换模式

    print(f"[{get_current_time()}] 等待进入主页")  # 输出等待进入主页信息

    while True:  # 无限循环等待主页
        current_page_url = driver.current_url  # 获取当前页面URL
        if "/exam/pc/home/" in current_page_url:  # 如果当前页面是主页
            print(f"[{get_current_time()}] 主动进入首页")  # 输出主动进入首页信息
            print(f"姓名：{get_user_full_name(driver)}")  # 输出用户姓名
            time.sleep(2)  # 等待2秒
            break  # 跳出循环

    while True:  # 无限循环处理课程页面
        try:
            process_course_page(driver, course, mute, playback_rate)  # 处理课程页面
        except Exception as e:  # 如果发生异常
            print(f"[{get_current_time()}] 发生错误：{str(e)}")  # 输出错误信息
            continue  # 继续循环


if __name__ == "__main__":  # 如果脚本作为主程序运行
    main()  # 执行主函数
