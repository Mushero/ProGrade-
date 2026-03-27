# 平时成绩管理系统 (Student Performance Management System)

> **一款为高校教师量身定制的轻量级、响应式课堂评分与点名系统。**


---

## 🌟 项目简介

本系统旨在解决高校课堂教学中平时成绩记录繁琐、互动性不足的问题。通过 **FastAPI** 后端驱动与 **Vue3** 现代化前端界面的有机结合，实现了学生名单管理、实时积分扣减、随机点名互动以及一键成绩归档功能。

---

## 🚀 核心功能

* **📊 智能化名单管理**：支持从标准 Excel 模板一键导入学生信息，自动识别班级、学号与姓名。
* **🎯 实时积分反馈**：内置自定义评分规则（玩手机、积极提问等），点击即刻完成评分，自动生成时间戳流水日志。
* **🎲 “天选之子”点名**：动态丝滑的随机点名动画，增加课堂趣味性，点名后自动跳转至打分界面。
* **📥 完整数据归档**：支持一键导出 Excel 报表，包含学生总分表与详细的扣分流水，方便期末直接核算。
* **💻 本地化部署**：基于 SQLite 本地数据库，数据存储在本地 `data.db`，无需联网，保护学生隐私安全。

---

## 🛠️ 技术栈

| 模块 | 技术实现 | 优势 |
| :--- | :--- | :--- |
| **前端** | Vue 3 + Tailwind CSS | 极简设计、极致响应速度 |
| **后端** | FastAPI + Uvicorn | 异步高性能、接口响应秒开 |
| **数据库** | SQLite | 轻量、免配置、易于迁移 |
| **GUI 容器** | PyWebView (Edge Chromium) | 现代化的桌面原生体验 |
| **打包工具** | pyinstaller | 方便快捷 |

---

## 📦 安装与使用

### 对于普通教师 (使用编译版)
1.  联系作者获取编译好的程序文件，可以**直接运行**
2.  下载发布的 `score_pro.exe` 文件。
3.  将其放置在您希望存储数据的文件夹中。
4.  **双击运行**：首次启动将自动创建 `data.db` 数据库。
5.  **导入名单**：准备好包含“学号、姓名、平时总分”的 Excel 文件进行导入。

### 对于开发者 (二次开发)
1.  **环境配置**：
    ```bash
    python -m venv venv
    source venv/bin/activate  # Windows: venv\Scripts\activate
    pip install fastapi uvicorn webview pandas openpyxl pyinstaller
    ```
2.  **运行程序**：
    ```bash
    python score_pro.py
    ```
3.  **编译打包** (推荐)：
    ```bash
    pyinstaller --noconfirm --onedir --windowed --add-data "data.db;." --hidden-import "uvicorn.logging" --hidden-import "uvicorn.loops" --hidden-import "uvicorn.loops.auto" --hidden-import "uvicorn.protocols.auto" --hidden-import "uvicorn.protocols.websockets.auto" --hidden-import "uvicorn.protocols.http.auto" --hidden-import "fastapi.middleware.cors" --hidden-import "multipart" "grade.py"
    ```

---

## 🤝 鸣谢与贡献

本项目由 **suwnd** 独立设计与开发。如果您在使用过程中有任何建议或发现了 Bug，欢迎通过以下方式反馈：

---

## ⚖️ 版权声明

本项目仅供学术交流与日常教学使用。未经作者许可，禁止用于任何形式的商业用途。

---
