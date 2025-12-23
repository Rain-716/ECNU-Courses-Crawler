import time
import json
import requests
import pandas as pd
import urllib3
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

# 禁用安全警告（针对校园网常见的自签名证书或SSL问题）
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= 核心配置区 =================

# 1. 账号信息
CREDENTIALS = {
    "username": "",  # 替换为你的学号
    "password": ""   # 替换为你的密码
}

# 2. 抓取页数设置
MAX_PAGES = 4  # 想要抓取多少页

# 3. 列配置（自定义删除列、重命名和排序）
# 格式： "API原始字段名": "Excel显示的中文列名"
# 【关键】：Excel中的列顺序将严格按照下面这个字典的顺序排列
TARGET_COLUMNS = {
    "course.nameZh": "课程名称",
    "nameZh": "教学班",
    "stdCount": "实际人数",
    "limitCount": "上限人数",
    "course.credits": "学分",
    "compulsorys": "课程属性",
    "courseType.nameZh": "课程类型",
    "courseProperty.nameZh": "课程性质",
    "examMode.nameZh": "考核方式",
    "teachLang.nameZh": "授课语言",
    "course.code": "课程代码",
    "course.id": "课程编号",
    "code": "代码",
    "id": "编号",
    "openDepartment.nameZh": "开课部门",
    "teacherAssignmentList": "授课教师",
    "scheduleText.dateTimeText.text": "日期时间",
    "scheduleText.roomSeatText.text": "房间座位数",
    "minorCourse": "辅修课程",
    "requiredPeriodInfo.total": "总学时",
    "requiredPeriodInfo.weeks": "周数",
    "requiredPeriodInfo.periodsPerWeek": "周学时",    
    "hasTeachingSyllabus": "有教学大纲",
    "hasTeachingStructuredDocument": "有教学结构文档",
    "teachingSyllabusFileInfo.name": "教学大纲文件名称",
    "teachingSyllabusFileInfo.id": "教学大纲文件编号",
    "teachingSyllabusFileInfo.mimeType": "教学大纲文件类型",
    "teachingSyllabusFileInfo.bytes": "教学大纲文件字节数",
    "teachingSyllabusFileInfo.sizeOfKb": "教学大纲文件大小(KB)",
    "teachingSyllabusFileInfo.sizeOfMb": "教学大纲文件大小(MB)",
    "teachingSyllabusFileInfo.openKey": "教学大纲文件公钥",
    "teachingSyllabusFileInfo.key": "教学大纲文件钥",
    "teachingSyllabusFileInfo.updateDateTime": "教学大纲文件更新时间",
    "teachingSyllabusFileInfo.createDateTime": "教学大纲文件创建时间",
    "teachingSyllabusFileInfo.business": "教学大纲文件",
    "remark": "备注",
}

# 4. URL 模板 (勿动 queryPage__={page})
BASE_URL_TEMPLATE = (
    "https://byyt.ecnu.edu.cn/student/for-std/lesson-search/search/826224?bizTypeAssoc=2&searchTeachingSyllabus=true&"
    "semesterAssocs=1629&"
    "queryPage__={page}%2C1000&"  # {page} 是占位符
    "assembleFields=course.code%2CminorCourse.nameZh%2CcourseType%2CopenDepartment%2CteacherAssignmentList%2CexamMode%2Ccampus%2CteachLang%2CroomType%2CtimeTableLayout%2CcrossBizTypes%2CcourseProperty"
)

LOGIN_URL = "https://sso.ecnu.edu.cn/login?service=https:%2F%2Fbyyt.ecnu.edu.cn%2Fsso%2Flogin"
OUTPUT_FILE = "ECNU_Courses_2_2.xlsx"

# ============================================

def init_driver():
    """初始化 Selenium Driver"""
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    # 屏蔽日志输出
    options.add_argument('--log-level=3') 
    options.add_argument('--disable-logging')
    options.add_argument('--disable-gpu')
    # 忽略 Google 服务相关的错误
    options.add_argument('--disable-software-rasterizer')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--ignore-certificate-errors') # 忽略证书错误
    
    # 自动下载并管理 ChromeDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def login_and_get_session(driver):
    """Selenium 登录并获取 requests Session"""
    print(f"[*] 正在访问登录页面...")
    driver.get(LOGIN_URL)
    wait = WebDriverWait(driver, 10)
    
    try:
        # 1. 自动填充账号密码
        username_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#nameInput")))
        password_input = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        submit_btn = driver.find_element(By.CSS_SELECTOR, "#submitBtn")

        print("[*] 输入账号密码...")
        username_input.clear()
        username_input.send_keys(CREDENTIALS['username'])
        password_input.clear()
        password_input.send_keys(CREDENTIALS['password'])
        submit_btn.click()

        # 2. 等待跳转至教务系统首页
        # 这是为了确保 Server 端 Session 已经完全建立
        print("[*] 等待跳转至教务系统首页 (/home)...")
        wait.until(EC.url_contains("/home"))
        
        # 额外等待确保 Cookie 稳定
        time.sleep(3) 
        
        # 3. 转移 Cookie 到 Requests Session
        session = requests.Session()
        # 伪装 User-Agent
        session.headers.update({
            "User-Agent": driver.execute_script("return navigator.userAgent")
        })
        for cookie in driver.get_cookies():
            session.cookies.set(cookie['name'], cookie['value'])
        
        print("[+] 登录成功，Session 已建立。")
        return session

    except Exception as e:
        print(f"[!] 登录失败或超时: {e}")
        return None

def fetch_data(session):
    """遍历页码抓取数据"""
    all_courses = []
    print(f"[*] 开始抓取前 {MAX_PAGES} 页数据...")
    
    for page in range(1, MAX_PAGES + 1):
        url = BASE_URL_TEMPLATE.format(page=page)
        print(f"    -> 正在请求第 {page} 页...")
        
        try:
            # verify=False 忽略 SSL 报错
            resp = session.get(url, verify=False, timeout=5)
            
            # 安全检查：是否被重定向回了 HTML 页面
            if "<html" in resp.text[:100].lower():
                print(f"      [!] 警告：第 {page} 页返回了 HTML，可能 Session 已失效。")
                continue

            data = resp.json()
            
            # 兼容处理：获取 list 数据
            rows = data.get('data', data) if isinstance(data, dict) else data
            
            if rows:
                all_courses.extend(rows)
                print(f"      成功获取 {len(rows)} 条记录")
            else:
                print("      本页无数据")
                
        except Exception as e:
            print(f"      [!] 第 {page} 页抓取失败: {e}")
            
    return all_courses

def save_and_format_excel(courses_data):
    """清洗数据、筛选列、保存并美化"""
    if not courses_data:
        print("[!] 总共未获取到任何数据，程序结束。")
        return

    print("[*] 正在处理数据结构...")
    
    # 1. 展平 JSON (将 course.code 这种嵌套结构展平)
    df = pd.json_normalize(courses_data)
    
    # 2. 【核心步骤】筛选和排序列
    # 找出 TARGET_COLUMNS 中存在于数据里的列
    valid_cols = [col for col in TARGET_COLUMNS.keys() if col in df.columns]
    
    if valid_cols:
        # 筛选并重排序
        df = df[valid_cols]
        # 重命名为中文
        df.rename(columns=TARGET_COLUMNS, inplace=True)
    else:
        print("[!] 警告：配置的列名在数据中均未找到，将导出所有原始列。")
        # 调试用：打印实际存在的列名
        # print(df.columns.tolist())

    # 3. 数据清洗：防止 Excel 遇到 list/dict 报错
    for col in df.columns:
        if df[col].apply(lambda x: isinstance(x, (list, dict))).any():
            df[col] = df[col].astype(str)

    # 4. 保存原始 Excel
    df.to_excel(OUTPUT_FILE, index=False)

    # 5. 使用 OpenPyXL 进行样式美化
    print("[*] 正在进行 Excel 样式美化（调整列宽、自动换行）...")
    wb = load_workbook(OUTPUT_FILE)
    ws = wb.active

    # 定义样式：自动换行，垂直居中，左对齐
    alignment_style = Alignment(wrap_text=True, vertical='center', horizontal='left')
    header_font = Font(bold=True) # 表头加粗

    for column_cells in ws.columns:
        col_letter = column_cells[0].column_letter
        max_length = 0
        
        for cell in column_cells:
            # 应用对齐样式
            cell.alignment = alignment_style
            
            # 表头加粗
            if cell.row == 1:
                cell.font = header_font
            
            # 计算最大内容长度（用于列宽）
            try:
                if cell.value:
                    # 简单估算：中文字符长度加权
                    val_str = str(cell.value)
                    val_len = len(val_str.encode('utf-8'))
                    if val_len > max_length:
                        max_length = val_len
            except:
                pass
        
        # 智能调整列宽：最小 10，最大 50 (防止某一列过宽占满屏幕)
        adjusted_width = min(max(max_length * 0.7, 10), 50)
        ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(OUTPUT_FILE)
    print(f"[+] 大功告成！文件已保存至: {OUTPUT_FILE}")

def main():
    driver = init_driver()
    try:
        # 第一步：登录
        session = login_and_get_session(driver)
        
        if session:
            # 第二步：抓取
            courses = fetch_data(session)
            
            # 第三步：处理并保存
            save_and_format_excel(courses)
    finally:
        # 无论成功失败，关闭浏览器
        driver.quit()

if __name__ == "__main__":
    main()