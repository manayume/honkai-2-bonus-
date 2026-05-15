import win32gui
import win32api
import win32con
import time
import tkinter as tk
import pyautogui
from typing import Tuple
import cv2
import numpy as np
from PIL import ImageGrab
import sys
import random

# 切换角色刷需要重新截取相应的图片，并命名
tiaozhan = "test/tiaozhan_diyala.png"
shuangbeiquan = "test/shuangbeiquan_diyala.png"

pyautogui.FAILSAFE = True  # 鼠标移动到屏幕角落可中断
pyautogui.PAUSE = 0.1  # 每次操作后暂停0.1秒


def select_and_highlight_window():
    print("请用鼠标点击目标窗口...")

    # 等待鼠标点击
    while True:
        if win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000:
            click_pos = win32api.GetCursorPos()
            while win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000:
                time.sleep(0.01)
            break
        time.sleep(0.05)

    # 获取窗口句柄
    hwnd = win32gui.WindowFromPoint(click_pos)
    window_title = win32gui.GetWindowText(hwnd)
    rect = win32gui.GetWindowRect(hwnd)

    print(f"选中窗口: {window_title}")
    print(f"窗口区域: {rect}")

    # 显示红框
    root = tk.Tk()
    root.overrideredirect(True)
    root.attributes('-topmost', True)
    root.attributes('-alpha', 0.3)
    root.geometry(f"{rect[2] - rect[0]}x{rect[3] - rect[1]}+{rect[0]}+{rect[1]}")

    canvas = tk.Canvas(root, highlightthickness=3, highlightcolor='red')
    canvas.pack(fill=tk.BOTH, expand=True)
    canvas.create_rectangle(2, 2, rect[2] - rect[0] - 4, rect[3] - rect[1] - 4,
                            outline='red', width=3)

    # 显示提示文字
    label = tk.Label(root, text="✓ 窗口已选中", font=("Arial", 16, "bold"),
                     fg="red", bg="yellow")
    label.place(relx=0.5, rely=0.1, anchor='n')

    # 3秒后自动关闭
    root.after(30, root.destroy)
    root.mainloop()

    return hwnd, rect


def mouse_click_in_window(window_rect: Tuple[int, int, int, int],
                          relative_x: int,
                          relative_y: int,
                          button: str = 'left',
                          clicks: int = 1,
                          interval: float = 0.1,
                          move_duration: float = 0.2,
                          check_window_active: bool = True) -> bool:
    """
    在指定窗口内移动鼠标并点击（基于窗口相对坐标）
    Args:
        window_rect: 窗口区域 (left, top, right, bottom)
        relative_x: 相对于窗口左上角的X坐标（像素）
        relative_y: 相对于窗口左上角的Y坐标（像素）
        button: 鼠标按键 'left', 'right', 'middle'
        clicks: 点击次数
        interval: 点击间隔（秒）
        move_duration: 鼠标移动时间（秒）
        check_window_active: 是否检查窗口是否激活

    Returns:
        bool: 点击成功返回True，失败返回False
    """
    try:
        # 解包窗口坐标
        left, top, right, bottom = window_rect

        # 检查窗口区域是否有效
        if left >= right or top >= bottom:
            print(f"错误：无效的窗口区域 {window_rect}")
            return False

        # 检查相对坐标是否在窗口范围内
        window_width = right - left
        window_height = bottom - top

        if relative_x < 0 or relative_x > window_width:
            print(f"警告：X坐标 {relative_x} 超出窗口宽度 {window_width}")
            # 限制坐标范围
            relative_x = max(0, min(relative_x, window_width))

        if relative_y < 0 or relative_y > window_height:
            print(f"警告：Y坐标 {relative_y} 超出窗口高度 {window_height}")
            relative_y = max(0, min(relative_y, window_height))

        # 计算屏幕绝对坐标
        screen_x = left + relative_x
        screen_y = top + relative_y

        # 获取屏幕尺寸
        screen_width, screen_height = pyautogui.size()

        # 检查屏幕坐标是否有效
        if screen_x < 0 or screen_x > screen_width or screen_y < 0 or screen_y > screen_height:
            print(f"错误：屏幕坐标 ({screen_x}, {screen_y}) 超出屏幕范围")
            return False
        '''
        print(f"目标信息:")
        print(f"  窗口区域: ({left}, {top}) -> ({right}, {bottom})")
        print(f"  窗口尺寸: {window_width} x {window_height}")
        print(f"  相对坐标: ({relative_x}, {relative_y})")
        print(f"  屏幕坐标: ({screen_x}, {screen_y})")
        '''
        # 可选：激活窗口
        if check_window_active:
            # 查找窗口句柄（通过窗口位置）
            hwnd = win32gui.WindowFromPoint((screen_x, screen_y))
            if hwnd:
                # 激活窗口
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.2)  # 等待窗口激活

        # 移动鼠标到目标位置
        # print(f"正在移动鼠标到 ({screen_x}, {screen_y})...")
        pyautogui.moveTo(screen_x, screen_y, duration=move_duration)

        # 短暂停顿，确保鼠标到位
        time.sleep(0.1)

        # 执行点击
        # print(f"执行 {button} 键点击...")
        pyautogui.click(button=button, clicks=clicks, interval=interval)

        print(f"✓ 点击成功！")
        return True

    except pyautogui.FailSafeException:
        print("✗ 操作被用户中断（鼠标移动到屏幕角落）")
        return False
    except Exception as e:
        print(f"✗ 点击失败: {str(e)}")
        return False


def match_template_in_window(window_hwnd, template_path, threshold=0.8):
    """
    在指定窗口中匹配模板，用红框标注并输出位置信息

    Args:
        window_hwnd: 窗口句柄（Windows）
        template_path: 模板图片路径（位于test文件夹中的图片）
        threshold: 匹配阈值 (0~1)

    Returns:
        list: 匹配位置列表，每个元素为 (x, y, w, h)
    """
    # 1. 获取窗口位置和尺寸
    rect = win32gui.GetWindowRect(window_hwnd)
    left, top, right, bottom = rect

    # 2. 截取窗口内容
    screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
    screen_img = np.array(screenshot)
    screen_img = cv2.cvtColor(screen_img, cv2.COLOR_RGB2BGR)

    # 3. 读取模板
    template = cv2.imread(template_path)
    if template is None:
        raise FileNotFoundError(f"无法加载模板图片: {template_path}")

    # 4. 灰度转换
    screen_gray = cv2.cvtColor(screen_img, cv2.COLOR_BGR2GRAY)
    template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
    h, w = template_gray.shape

    # 5. 模板匹配
    result = cv2.matchTemplate(screen_gray, template_gray, cv2.TM_CCOEFF_NORMED)
    locations = np.where(result >= threshold)

    # 6. 收集匹配位置（去重）
    matched_positions = []
    for pt in zip(*locations[::-1]):  # 注意：返回的是(y,x)顺序，需转成(x,y)
        # 检查是否与已有框重叠（简单去重）
        x, y = pt
        if not any(abs(x - px) < w and abs(y - py) < h for (px, py, _, _) in matched_positions):
            matched_positions.append((x, y, w, h))

    # 7. 在截图上绘制红框，并输出坐标信息
    '''
    result_img = screen_img.copy()
    for (x, y, w, h) in matched_positions:
        # 绘制红色矩形（边框粗细2）
        cv2.rectangle(result_img, (x, y), (x + w, y + h), (0, 0, 255), 2)
        # 输出信息
        print(f"匹配位置: ({x}, {y}), 宽度: {w}, 高度: {h}")
    '''
    return matched_positions
    # 8. 显示结果窗口（按任意键关闭）


'''
    cv2.imshow('Template Match Result', result_img)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
'''


def replenish_physical_strength(hwnd,  template_path, threshold):
    matches = match_template_in_window(hwnd, template_path, threshold)
    if matches:
        print("需要补充体力")
        return True
    else:
        return False


def buchong(hwnd, threshold, button):

    goumai = match_template_in_window(hwnd, "test/goumai.png", threshold)
    duihuan = match_template_in_window(hwnd, "test/duihuan.png", threshold)
    if goumai:
        matches = goumai
        clicks1 = 1
    if duihuan:
        matches = duihuan
        clicks1 = 2
    matches_x = matches[0][0]
    matches_y = matches[0][1]
    matches_w = matches[0][2]
    matches_h = matches[0][3]
    x_num = random.randrange(-matches_w // 2 + 10, matches_w // 2 - 10, 4)
    y_num = random.randrange(-matches_h // 2+10,matches_h // 2-10,4)
    print("x_num=", x_num)
    print("y_num=", y_num)
    x = matches_x + matches_w // 2 + x_num
    y = matches_y + matches_h // 2+y_num
    mouse_click_in_window(rect, x, y, button, clicks1)
    print("补充体力成功，重新获取窗口")
    matches = match_template_in_window(hwnd, tiaozhan, threshold)
    key = 1
    return matches, key


def find_and_click_once(hwnd, rect, template_path, step, threshold=0.8, button='left', clicks=1):
    """
    在指定窗口中匹配模板，用红框标注并输出位置信息

    Args:
        hwnd: 窗口句柄（Windows）
        rect: 窗口的坐标宽高
        template_path: 模板图片路径（位于test文件夹中的图片）
        threshold: 匹配阈值 (0~1)
        step: 当前的步骤

    Returns:
        key:是否触发补充体力
    """
    if step == 0:
        print("选备用装备")
    elif step == 1:
        print("开战！！！")
    elif step == 2:
        print("切备用")
    elif step == 3:
        print("再次挑战")
    elif step == 4:
        print("使用双倍券")
    matches = match_template_in_window(hwnd, template_path, threshold)
    key = 0
    t = 0
    i=0
    while not matches:
        if replenish_physical_strength(hwnd,  "test/buchong.png", 0.7):
            matches, key = buchong(hwnd, threshold, button)

        else:
            matches = match_template_in_window(hwnd, template_path, threshold)
            # time.sleep(2)
            if i==0:
                print("未检测到")
                i=i+1

    matches_x = matches[0][0]
    matches_y = matches[0][1]
    matches_w = matches[0][2]
    matches_h = matches[0][3]
    x_num = random.randrange(-matches_w // 2+10,matches_w // 2-10,4)
    y_num = random.randrange(-matches_h // 2+10,matches_h // 2-10,4)
    x = matches_x + matches_w // 2+x_num
    y = matches_y + matches_h // 2+y_num
    if step == 4:
        x = matches_x + 8 * matches_w // 9
    if step == 3:
        x = 527+random.randrange(-0,20,4)
        print("再次挑战 x=",x)
    mouse_click_in_window(rect, x, y, button, clicks)
    while (match_template_in_window(hwnd, template_path, threshold) == matches) and (step == 3 or step == 1) and t < 4:
        if replenish_physical_strength(hwnd,  "test/buchong.png", 0.7):
            matches, key = buchong(hwnd, threshold, button)
        print("未正确跳到下一步，重试", step)
        time.sleep(1)
        mouse_click_in_window(rect, x, y, button, clicks)
        t = t + 1
    if t == 4:
        print("已重试4次退出程序，请重启模拟器或游戏")
        sys.exit()
    return key


if __name__ == "__main__":

    j = int(input("是否切备用（输入1/0）"))  # 是否切备用
    s = int(input("是否用双倍券（输入1/0）"))  # 是否用双倍券
    hwnd, rect = select_and_highlight_window()
    template_path = ["test/beiyong.png", "test/kaizhan.png", "test/qie.png", tiaozhan, shuangbeiquan]
    for i in range(0, 100):
        print("第", i, "轮")
        # '''
        k = find_and_click_once(hwnd, rect, template_path[0], step=0)
        if k:
            continue
        k = find_and_click_once(hwnd, rect, template_path[1], step=1)
        if k:
            continue
        # '''
        # 不切备用
        if j != 0:
            k = find_and_click_once(hwnd, rect, template_path[2], step=2)
            if k:
                continue

        # 双倍券，不用注释掉即可
        if s != 0:
            k = find_and_click_once(hwnd, rect, template_path[4], step=4)
            if k:
                continue

        k = find_and_click_once(hwnd, rect, template_path[3], step=3)
        if k:
            continue
