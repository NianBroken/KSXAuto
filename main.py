import time
import configparser
import sys
import os
import ctypes
import json
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import datetime


# 设置控制台标题
def set_console_title(title):
    ctypes.windll.kernel32.SetConsoleTitleW(title)


# 设置控制台标题为
set_console_title("KSXAuto")

# 获取当前文件所在目录的绝对路径
current_path = os.path.dirname(os.path.realpath(sys.argv[0]))

# 创建ConfigParser对象
config = configparser.ConfigParser()

# 读取配置文件
config.read(os.path.join(os.path.dirname(current_path), "config.ini"))

# 获取配置项
videoConfig = "videoConfig"
muteConfig = config.get(videoConfig, "mute") or True
playbackRateConfig = config.get(videoConfig, "playbackRate") or 1

# 获取配置项
loginConfig = "loginConfig"
urlConfig = config.get(loginConfig, "url")
usernameConfig = config.get(loginConfig, "username")
passwordeConfig = config.get(loginConfig, "password")
courseConfig = config.get(loginConfig, "course")

# 登陆地址
url = urlConfig

# 课程地址
course = courseConfig

# 账号密码
username = usernameConfig
password = passwordeConfig

# 是否静音播放
mute = muteConfig

# 播放倍数
playbackRate = playbackRateConfig

# 初始化变量
click_play = None
delay = 1

# 浏览器配置
options = webdriver.EdgeOptions()
# 保持浏览器打开状态
options.add_experimental_option("detach", True)
# 禁止浏览器日志输出
options.add_experimental_option("excludeSwitches", ["enable-logging"])
# 启动浏览器最大化
options.add_argument("--start-maximized")
# 创建浏览器实例
driver = webdriver.Edge(options=options)
# 隐式等待30秒
driver.implicitly_wait(2)


def check_updates():
    # 读取远程Json信息
    local_url = "https://gitee.com/nianbroken/KSXAuto/raw/main/version_info.json"

    response = requests.get(local_url)

    if response.status_code == 200:
        local_url_content = response.text
        remote_software_data = json.loads(local_url_content)
        remote_software_info = remote_software_data["software_info"]
        remote_version_id = remote_software_info["version_id"]
        local_software_info_file_path = os.path.join(current_path, "version_info.json")
        if os.path.exists(local_software_info_file_path):
            with open(local_software_info_file_path, "r") as local_software_info:
                local_software_data = json.load(local_software_info)
                # 提取本地软件信息
                local_software_info = local_software_data["software_info"]
                local_version_id = local_software_info["version_id"]
                print(f"本地版本号：{local_version_id}")
                print(f"远程版本号：{remote_version_id}")
                if local_version_id < remote_version_id:
                    print("检测到存在新版本")
                    print(f"新版本下载地址：{remote_software_info['download_url']}")
        else:
            print(f"新版本下载地址：{remote_software_info['download_url']}")
    else:
        print("获取远程Json信息失败：", response.status_code)


# 获取当前时间
def current_time():
    current_time = datetime.now().strftime("%m-%d %H:%M:%S")
    return current_time


# 登陆函数
def user_login(username, password):
    # 如果配置文件中填写了用户名
    if username:
        # 找到账号输入框
        username_inputbox = driver.find_element(By.XPATH, '//*[@id="username"]')
        # 清空输入框内容
        username_inputbox.clear()
        # 输入账号
        username_inputbox.send_keys(username)
        print(f"[{current_time()}] 被动填写账号")

    # 如果配置文件中填写了密码
    if password:
        # 找到密码输入框
        password_inputbox = driver.find_element(By.XPATH, '//*[@id="userTypePwd"]')
        # 清空输入框内容
        password_inputbox.clear()
        # 输入密码
        password_inputbox.send_keys(password)
        print(f"[{current_time()}] 被动填写密码")

    # 如果配置文件中填写了用户名和密码，则点击登陆按钮
    if username and password:
        driver.find_element(By.XPATH, '//*[@id="loginBtn"]').click()
        print(f"[{current_time()}] 被动登录")


# 得到用户姓名
def get_full_name():
    # 声明使用全局变量
    global full_name

    # 得到用户姓名
    full_name = driver.find_element(
        By.XPATH,
        '//*[@id="viewFrameWorkBody"]/div/div[1]/div[1]/div/div/div[2]',
    ).text

    # 返回用户姓名
    return full_name


# 得到视频总数
def videos_number():
    # 声明使用全局变量
    global videos_number

    # 得到课程的所有视频
    course_video_list = driver.find_elements(By.CLASS_NAME, "catalog-item")

    # 得到视频总数
    videos_number = len(course_video_list)

    # 返回视频总数
    return videos_number


# 找到课程所有未完成的视频列表
def find_all_unfinished_videos_list():
    # 声明使用全局变量
    global unfinished_videos, unfinished_videos_number, formatted_unfinished_videos, first_unfinished_video_id

    # 初始化变量
    unfinished_videos = []

    # 如果存在视频
    if videos_number:
        # 遍历课程的所有目录
        for i in range(1, videos_number + 1):
            # 视频完成度xpath路径
            completion_degree = f'//*[@id="viewFrameWorkBody"]/div/div[3]/div[2]/div/div[2]/div[{i}]/div[2]/div[2]/div[2]'

            # 得到视频完成度
            completion_degree_text = driver.find_element(
                By.XPATH, completion_degree
            ).text

            # 如果视频完成度不为100%，则加入未完成视频列表中
            if completion_degree_text != "100%":
                unfinished_videos.append(i)

    # 未完成视频数量
    unfinished_videos_number = len(unfinished_videos)

    # 如果存在未完成的视频
    if unfinished_videos_number:
        # 未完成的视频序号
        formatted_unfinished_videos = "、".join(map(str, unfinished_videos))

        # 得到第一个未完成的视频的ID
        first_unfinished_video_id = unfinished_videos[0]
    else:
        formatted_unfinished_videos = None
        first_unfinished_video_id = None

    # 返回视频总数、未完成的视频列表、未完成视频数量、未完成的视频序号、第一个未完成的视频的ID、运行日志
    return (
        unfinished_videos,
        unfinished_videos_number,
        formatted_unfinished_videos,
        first_unfinished_video_id,
    )


# 得到第一个未完成的视频的element
def get_first_unfinished_video_element():
    # 声明使用全局变量
    global first_unfinished_video, first_unfinished_video_title

    # 得到第一个未完成的视频的Xpath路径
    first_unfinished_video_xpath = f'//*[@id="viewFrameWorkBody"]/div/div[3]/div[2]/div/div[2]/div[{first_unfinished_video_id}]/div[2]/div[1]'

    # 得到第一个未完成的视频
    first_unfinished_video = driver.find_element(By.XPATH, first_unfinished_video_xpath)

    # 得到第一个未完成的视频的标题
    first_unfinished_video_title = first_unfinished_video.text

    # 返回第一个未完成的视频的element、第一个未完成的视频的标题
    return first_unfinished_video, first_unfinished_video_title


# 播放第一个未完成的视频
def play_first_unfinished_video():
    # 声明使用全局变量
    global play_log, click_play

    # 打开第一个未完成的视频
    first_unfinished_video.click()

    # 获得所有视频elements
    video_elements = driver.find_elements(By.TAG_NAME, "Video")

    # 对视频静音
    if mute and len(video_elements):
        for video in video_elements:
            driver.execute_script("arguments[0].muted = true;", video)
            driver.execute_script(f"arguments[0].playbackRate = {playbackRate};", video)

    try:
        # 播放第一个未完成的视频
        if driver.find_element(By.XPATH, '//*[@id="myVideo"]/button').is_displayed():
            driver.find_element(By.XPATH, '//*[@id="myVideo"]/button').click()
            play_log = f"[{current_time()}] 正在为你播放第 {first_unfinished_video_id} 个视频：{first_unfinished_video_title}"
            # 本次循环是否点击了播放按钮
            click_play = True
        else:
            click_play = False
        # 输出相关信息

    except Exception:
        # 本次循环是否点击了播放按钮
        click_play = False


# 得到视频剩余时间
def get_video_remaining_time():
    global video_remaining_time_text1, video_remaining_time_text2, video_remaining_time1, time_failed
    try:

        video_remaining_time_text1_displayed = driver.find_element(
            By.XPATH, '//*[@id="myVideo"]/div[1]/div[2]'
        ).is_displayed()

        video_remaining_time_text2 = driver.find_element(
            By.XPATH, '//*[@id="myVideo"]/div[5]/div[5]/span[3]'
        ).text

        if video_remaining_time_text1_displayed:
            video_remaining_time_text1 = driver.find_element(
                By.XPATH, '//*[@id="myVideo"]/div[1]/div[2]'
            ).text
            video_remaining_time1 = f"当前视频剩余：{video_remaining_time_text1}"

        time_failed = False

        return video_remaining_time1
    except Exception:
        time_failed = True


check_updates()
print("------")
print(f"静音：{muteConfig}")
print(f"静音：{muteConfig}")
print(f"倍数：{playbackRateConfig}")
print(f"运行路径：{current_path}")
print("------")
print(f"[{current_time()}] 程序启动成功 等待操作")

if url:
    # 打开登陆页面
    driver.get(url)
    print(f"[{current_time()}] 被动进入登录页面")
else:
    while True:
        # 循环检测当前的url是否为登陆页面
        current_page_url = driver.current_url
        if "login/account/login" in current_page_url:
            print(f"[{current_time()}] 主动进入登录页面")
            break
        time.sleep(delay)

user_login(username, password)
time.sleep(delay)

# 得到登录的错误信息
current_page_url = driver.current_url
if "login/account/login" in current_page_url:
    login_errormsg = driver.find_element(By.XPATH, '//*[@id="errormsg"]')
    if login_errormsg.is_displayed() and "错误" in login_errormsg.text:
        print(f"[{current_time()}] {login_errormsg.text} 已切换主动模式")


while True:
    # 循环检测当前页面的url是否为主页
    current_page_url = driver.current_url
    if "/exam/pc/home/" in current_page_url:
        print(f"[{current_time()}] 主动进入首页")
        # 得到用户姓名
        get_full_name()
        print(f"姓名：{full_name}")
        time.sleep(2)
        break

if course:
    # 打开课程页面
    driver.get(course)
    print(f"[{current_time()}] 被动进入课程页面")
    time.sleep(2)
else:
    while True:
        current_page_url = driver.current_url
        if "/exam/pc/course/#/study?courseId" in current_page_url:
            print(f"[{current_time()}] 主动进入课程页面")

videos_number()
print(f"视频总数：{videos_number}")

try:
    while True:
        find_all_unfinished_videos_list()
        if first_unfinished_video_id:
            if click_play is None:
                print(f"未完成的视频总数：{unfinished_videos_number}")
                print(f"未完成的视频序号：{formatted_unfinished_videos}")
                print("------")
            get_first_unfinished_video_element()
            play_first_unfinished_video()
            if click_play:
                print(play_log)
            while True:
                try:
                    if driver.find_element(
                        By.XPATH, '//*[@id="myVideo"]/button'
                    ).is_displayed():
                        driver.find_element(
                            By.XPATH, '//*[@id="myVideo"]/button'
                        ).click()
                except Exception:
                    pass
                get_video_remaining_time()
                remaining_displayed = driver.find_element(
                    By.XPATH, '//*[@id="myVideo"]/div[1]/div[2]'
                ).is_displayed()
                print(video_remaining_time1, end="\r", flush=True)
                if (
                    video_remaining_time_text1 == "00:00:00"
                    or video_remaining_time_text2 == "0:00"
                    or time_failed
                    or not remaining_displayed
                ):
                    break
        else:
            print(f"[{current_time()}] 不存在未完成的视频")
            sys.exit()
except Exception:
    print(f"[{current_time()}] 主动退出")
    sys.exit()
