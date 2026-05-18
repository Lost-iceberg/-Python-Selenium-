# 小红书综合模式 + 用户模式采集脚本

## 1. 项目说明


脚本当前支持两类任务：

1. **综合模式**
   - 按关键词搜索小红书内容
   - 采集综合搜索结果中的视频笔记与图文笔记
   - 采集笔记标题、标签、发布时间、点赞数、收藏数、评论数、链接等字段
   - 支持抓取指定数量的评论信息
   - 输出视频与图文笔记的数量统计和占比分析

2. **用户模式**
   - 按关键词搜索用户
   - 访问指定数量的用户主页
   - 采集用户基础资料
   - 在用户主页继续采集指定数量的视频笔记与图文笔记
   - 抓取这些笔记下的评论信息
   - 统计用户主页中已加载笔记的视频/图文数量
   - 尝试以最早可见笔记的发布时间近似表示内容创建时间

---

## 2. 文件结构

```text
redbook_spider_full_public.py
README_redbook_spider_full_public.md
output/
├── Raw_integrated_data.xlsx
├── Raw_user_data.xlsx
├── Raw_user_video_data.xlsx
└── Raw_user_text_data.xlsx
```

说明：

- `redbook_spider_full_public.py`：主程序
- `README_redbook_spider_full_public.md`：使用说明
- `output/`：程序运行后自动创建的数据输出目录

---

## 3. 环境要求

建议使用：

- Python 3.10 及以上
- Google Chrome 浏览器
- Selenium
- openpyxl

安装依赖：

```bash
pip install selenium openpyxl
```

---

## 4. 驱动配置

脚本默认使用 Selenium Manager 自动处理浏览器驱动：

```python
CHROMEDRIVER_PATH = ""
```

若本机环境无法自动启动浏览器，可以手动填写 ChromeDriver 路径：

```python
CHROMEDRIVER_PATH = r"你的 chromedriver.exe 路径"
```

示例：

```python
CHROMEDRIVER_PATH = r"D:\tools\chromedriver.exe"
```

该路径仅为示例，脚本中没有保留任何原设备的个人路径信息。

---

## 5. 输出路径配置

脚本使用相对路径，所有输出文件默认保存到脚本同级目录下的 `output` 文件夹：

```python
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
```

输出文件如下：

```python
INTEGRATED_FILE_PATH = str(OUTPUT_DIR / "Raw_integrated_data.xlsx")
USER_FILE_PATH = str(OUTPUT_DIR / "Raw_user_data.xlsx")
USER_VIDEO_FILE_PATH = str(OUTPUT_DIR / "Raw_user_video_data.xlsx")
USER_TEXT_FILE_PATH = str(OUTPUT_DIR / "Raw_user_text_data.xlsx")
```

---

## 6. 运行方式

在终端进入脚本所在目录后执行：

```bash
python redbook_spider_full_public.py
```

程序运行后会：

1. 自动打开小红书探索页
2. 等待手动登录
3. 登录完成后按回车
4. 选择运行模式
5. 根据提示输入关键词和采集数量

---

## 7. 综合模式说明

### 7.1 功能

综合模式用于按关键词直接采集笔记数据。

运行后需选择：

```text
1. 指定总爬取数量
2. 分别指定视频笔记和图文笔记数量
```

### 7.2 示例

假设输入：

```text
关键词：披萨 披萨店
每条笔记评论数：3
模式：2
视频笔记数量：5
图文笔记数量：10
```

程序会依次搜索：



并在综合结果页分别采集：

- 5 条视频笔记
- 10 条图文笔记
- 每条笔记前 3 条可识别评论

### 7.3 输出文件

综合模式数据写入：

```text
output/Raw_integrated_data.xlsx
```

工作表：

- `integrated_note`：笔记明细
- `integrated_analysis`：综合统计结果

---

## 8. 用户模式说明

### 8.1 功能

用户模式用于搜索用户并进一步访问主页采集资料和笔记。

运行时需要输入：

- 搜索关键词
- 访问的用户数量
- 每个用户主页要采集的视频笔记数量
- 每个用户主页要采集的图文笔记数量
- 每条笔记要采集的评论数量

### 8.2 示例

假设输入：

```text
关键词：披萨 探店
访问用户数量：3
每个用户主页视频笔记数量：2
每个用户主页图文笔记数量：2
每条笔记评论数量：3
```

程序会：

1. 搜索“披萨”
2. 进入用户搜索结果页
3. 访问前 3 个可获取用户主页链接的账号
4. 采集用户资料
5. 每个用户主页再采集：
   - 2 条视频笔记
   - 2 条图文笔记
   - 每条笔记的 3 条评论
6. 再对“探店”重复以上流程

### 8.3 输出文件

用户模式数据分别写入：

```text
output/Raw_user_data.xlsx
output/Raw_user_video_data.xlsx
output/Raw_user_text_data.xlsx
```

其中：

- `Raw_user_data.xlsx`：用户基础资料
- `Raw_user_video_data.xlsx`：用户主页采集到的视频笔记
- `Raw_user_text_data.xlsx`：用户主页采集到的图文笔记

每个搜索关键词会对应一个单独工作表，工作表名称会自动处理为合法的 Excel Sheet 名称。

---

## 9. 主要采集字段

### 9.1 综合模式与用户主页笔记字段

常见字段包括：

- `note_id`
- `name`
- `tags`
- `publish_time`
- `like_count`
- `collect_count`
- `comment_count`
- `note_url`
- `note_type`
- `video_length`
- `content`
- `user_id`

若设置评论抓取数量大于 0，还会生成：

- `comment1_content`
- `comment1_time`
- `comment1_loc`
- `comment1_like`
- `comment1_reply`

以及第 2 条、第 3 条……评论对应字段。

### 9.2 用户资料字段

用户模式基础资料包括：

- `User ID`
- `Account Name`
- `Certification Info`
- `Creation Time`
- `Gender`
- `Age`
- `Tags`
- `IP Location`
- `Following`
- `Followers`
- `Likes and Collections`
- `Video Count`
- `Text Count`
- `User URL`
- `Bio`

---

## 10. 重要说明

### 10.1 “Creation Time”的含义

`Creation Time` 并不是平台官方账号注册时间。

脚本采用的做法是：

1. 尽量滚动到用户主页底部
2. 点击最后一条可见笔记
3. 提取该笔记发布时间

因此它更接近：

> 当前页面尽可能加载后，最早可见笔记的发布时间

不能完全等同于真实账号创建时间。

### 10.2 “Video Count / Text Count”的含义

这两个字段是脚本在当前页面滚动加载后，基于已加载笔记卡片统计出的数量。

若平台未完全加载所有笔记，或者页面结构变化，该数量可能与账号真实总笔记数存在差异。

### 10.3 页面结构变动风险

该脚本依赖网页元素定位：

- XPath
- CSS Selector
- 页面中的类名和 DOM 结构

若小红书页面前端结构更新，部分定位可能失效，需要针对新页面重新调整选择器。

---

## 11. 使用建议

- 初次测试时，把采集数量设小一些，例如：
  - 用户数：1
  - 视频笔记：1
  - 图文笔记：1
  - 评论数：1
- 先确认浏览器能正常打开、登录、写入 Excel 后，再提高数量。
- 如需提高稳定性，可保留当前随机延迟逻辑，不建议将延迟全部删除。
- 默认不建议开启无头模式，因为手动登录和部分页面渲染在无头环境下可能不稳定。

---

