# ECNU Courses Crawler

**这是为 `ECNU_Courses_Crawler.py` 项目准备的 README。**
该脚本用于自动登录华东师范大学教务系统并抓取课程信息，导出并美化为 Excel 文件。

---

# 目录

* 项目简介
* 功能亮点
* 运行环境与依赖
* 快速开始
* 配置说明
* 输出说明
* 常见问题与排查建议
* 扩展与改进方向

---

# 项目简介

该脚本通过 Selenium 自动化登录到学校的单点登录（SSO），将浏览器会话 Cookie 转移到 `requests.Session()`，然后使用课程查询 API 分页抓取数据，最后将数据展平成 DataFrame 并导出为 Excel，同时使用 `openpyxl` 调整列宽与样式，最终文件名默认为 `ECNU_Courses_Final.xlsx`。更多实现细节请查看源文件。 

---

# 功能亮点

* 自动登录并保持会话（Selenium -> requests cookie 转移）。 
* 分页抓取（可配置抓取页数）。 
* JSON 展平、列筛选、中文列名重命名与排序（通过 `TARGET_COLUMNS` 配置）。 
* 导出 Excel 并使用 OpenPyXL 美化（自动换行、列宽、自适应表头加粗）。 

---

# 运行环境与依赖

建议在 Python 3.8+ 虚拟环境中运行。

主要依赖可通过 `pip` 安装：

```
pip install selenium webdriver-manager requests pandas openpyxl urllib3
```

> 备注：脚本使用 `webdriver-manager` 自动下载 ChromeDriver，所以只要本机安装了兼容版本的 Chrome 即可。

---

# 快速开始（安装 + 运行）

1. 克隆或将 `ECNU_Courses_Crawler.py` 放到工作目录。 
2. 安装依赖（见上文）。
3. 打开 `ECNU_Courses_Crawler.py`，在配置区填写/修改账号与参数（见下一节）。 
4. 运行脚本：

```bash
python ECNU_Courses_Crawler.py
```

运行后成功会生成 `ECNU_Courses_Final.xlsx`（默认名），脚本会在终端打印抓取进度与保存路径。 

---

# 配置说明（脚本顶部的“核心配置区”）

脚本顶部有一段核心配置，你可以直接在脚本里修改：

* `CREDENTIALS`：存放学号与密码。 

  ```py
  CREDENTIALS = {
      "username": "你的学号",
      "password": "你的密码"
  }
  ```

* `MAX_PAGES`：要抓取的页数（整数）。默认 4。 

* `TARGET_COLUMNS`：控制导出 Excel 的列、显示中文名与顺序。字典的顺序即 Excel 列顺序。可以删除不需要的字段（字段名需与 API 返回的 JSON key 对应，例如 `course.nameZh` 等）。 

* `BASE_URL_TEMPLATE`：抓取用的 API URL 模板，其中 `{page}` 会被替换成页码（无需手动改动除非学期/参数变化）。 

* `LOGIN_URL`：SSO 登录页地址（通常不需要修改）。 

* `OUTPUT_FILE`：导出的 Excel 文件名，默认 `ECNU_Courses_Final.xlsx`。 

---

# 输出说明

* 成品文件：`ECNU_Courses_Final.xlsx`（默认）。包含筛选后、重命名后的列，已启用自动换行与列宽调整。 

---

# 常见问题与排查建议

1. **登录失败 / 页面被重定向回 HTML**

   * 脚本会检测到返回 HTML 并提示 Session 失效。可能原因：账号被二次验证、校园网需要额外认证、SSO 页面元素变化或验证码。
   * 解决方法：先手动在浏览器登录确认是否需要额外步骤；在 `init_driver()` 中去掉 `--headless`，观察 Selenium 自动化流程（可以看到页面与可能需要点击的验证码等）。 

2. **抓不到数据或响应超时**

   * 检查 `BASE_URL_TEMPLATE` 中的学期参数和其他 query 是否仍然有效。
   * 增加 `session.get` 的 `timeout` 或在失败处增加重试逻辑。

3. **Excel 中某些列显示为 Python 列表或字典文本**

   * 脚本已有处理：将含 list/dict 的列转为字符串。 

4. **ChromeDriver 版本或 Chrome 不兼容**

   * `webdriver-manager` 会尝试自动安装兼容的驱动，但如果本地 Chrome 版本过旧或过新，请更新 Chrome 或显式安装合适版本的 driver。

---

# 扩展与改进方向

* 将配置抽离为 `config.json` 或 `config.yaml`，并实现命令行参数支持（`argparse`），提高灵活性。
* 增加多线程/异步抓取与重试机制以提高稳定性。
* 将关键字段（如 `TARGET_COLUMNS`）生成一个可视化配置界面或用 Excel 模板驱动列选择。
* 增加导出 CSV/SQLite/数据库存储选项。

---