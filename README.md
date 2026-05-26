# PPT Master 🎯 — Google Gemini 版

AI 演示文稿生成网站，基于 Google Gemini（免费）+ python-pptx

## 功能
- ✅ 完全免费（基于 Gemini 免费额度）
- ✅ 用户可选 Flash-Lite（1000次/天）或 Flash（250次/天）
- ✅ 生成真实 .pptx 文件可下载编辑
- ✅ 8 种配色主题，7 种语言
- ✅ 无需用户填写任何 API Key

## 部署到 Railway（10分钟，全球可访问）

### 第一步：获取 Gemini API Key（免费）
1. 打开 https://aistudio.google.com/apikey
2. 登录 Google 账号
3. 点击 "Create API Key"
4. 复制保存 Key

### 第二步：上传到 GitHub
1. 去 https://github.com/new 新建仓库（名字如 ppt-master）
2. 把这个文件夹里所有文件上传上去

### 第三步：部署到 Railway
1. 打开 https://railway.app
2. 用 GitHub 账号登录
3. 点击 "New Project" → "Deploy from GitHub repo"
4. 选择你的 ppt-master 仓库
5. 等待部署完成（约2分钟）

### 第四步：设置环境变量
1. 在 Railway 项目页面，点击你的服务
2. 点击 "Variables" 标签
3. 添加：GEMINI_API_KEY = 你的Key
4. Railway 自动重启服务

### 完成！
Railway 会给你一个公网地址如：
https://ppt-master-production.up.railway.app

把这个地址分享给其他人，他们就可以免费使用了！

## 本地运行（测试用）

```bash
# 安装依赖
pip install -r requirements.txt

# 设置环境变量
set GEMINI_API_KEY=你的Key  # Windows
export GEMINI_API_KEY=你的Key  # Mac/Linux

# 启动
uvicorn main:app --reload --port 8000

# 打开浏览器
# http://localhost:8000
```

## 项目结构

```
pptmaster-gemini/
├── main.py            # FastAPI 后端 + PPT 生成逻辑
├── requirements.txt   # Python 依赖
├── Dockerfile         # Docker 容器配置
├── railway.toml       # Railway 部署配置
├── README.md
└── static/
    └── index.html     # 完整前端页面（无需 API Key 输入）
```

## 免费额度说明

| 模型 | 每天免费次数 | 速度 |
|------|------------|------|
| Gemini 2.5 Flash-Lite | 1000 次 | 最快 |
| Gemini 2.5 Flash | 250 次 | 较快 |

超出免费额度后，费用约 $0.001/次（不到1分钱）
