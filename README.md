# 产品开发竞对分析网站

这是一个网页端 VOC-PRODUCT 报表生成工具。用户通过浏览器上传 Amazon ASIN 原始 Excel，系统在服务器端生成 VOC、订单尺码颜色、使用场景、用户画像、开发方向等 Excel 报表。

## 核心能力

- 上传多个 ASIN Excel 文件
- 生成 VOC-PRODUCT 完整报告
- 生成竞对横向分析报告
- 支持背心/带罩杯背心、性感透视款、大码薄杯款
- 支持模块预览和单模块下载
- 支持客户自带大模型接口增强分析

## 大模型费用说明

本系统不内置、不提供、不代付任何大模型 API Key。

客户如需启用大模型增强分析，需要在网页中填写自己的模型接口：

- API Base URL
- Model Name
- API Key
- 接口协议：OpenAI Compatible

模型调用费用由客户自己的模型服务账号承担。API Key 默认只在当前生成任务中临时使用，不写入历史记录，不写入导出的 Excel。

如果不启用大模型，系统会使用本地规则和统计逻辑生成基础报表。

## 本地/服务器启动

安装依赖：

```bash
pip install -r requirements.txt
```

启动服务：

```bash
python webapp/server.py
```

访问：

```text
http://127.0.0.1:8787/
```

服务器部署时可设置：

```bash
HOST=0.0.0.0
PORT=8787
MAX_UPLOAD_MB=300
python webapp/server.py
```

## Docker 部署

构建镜像：

```bash
docker build -t voc-product-web .
```

运行：

```bash
docker run -p 8787:8787 voc-product-web
```

访问：

```text
http://服务器IP:8787/
```

## OpenAI-compatible 接口示例

DeepSeek 示例：

```text
API Base URL: https://api.deepseek.com
Model Name: deepseek-chat
API Key: 客户自己的 Key
```

其他兼容 OpenAI Chat Completions 的服务也可以使用同样方式接入。

## 生成流程

1. 用户打开网页。
2. 填写导入人、产品分类、款号。
3. 选择报告类型和产品项目。
4. 可选：开启大模型增强分析并填写客户模型接口。
5. 上传 Excel 文件。
6. 点击生成报告。
7. 下载完整 Excel 或单独下载各模块。

## 安全与仓库注意事项

不要提交以下内容到 GitHub：

- `.env`
- 客户 API Key
- 上传的原始 Excel
- 生成的输出报告
- `outputs/`
- 本地缓存和日志

仓库已通过 `.gitignore` 默认排除这些文件。
