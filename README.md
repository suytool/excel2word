# Excel转Word (suytool)

飞牛 fnOS fpk 应用：批量将 Excel 数据填充到 Word 模板。

## 技术栈
- Go 1.23 + excelize/v2
- 前端：原生 HTML/CSS/JS
- 打包：fnpack CLI

## 本地开发
```bash
go mod tidy
go run .
# 浏览器打开 http://localhost:8080
```

## 交叉编译 Linux amd64
```bash
CGO_ENABLED=0 GOOS=linux GOARCH=amd64 go build -o excel2word -ldflags="-s -w" .
```

## fpk 打包
```bash
./fnpack create myapp -t native   # 创建模板
# 将二进制放入 app/bin/，静态文件放入 app/static/
./fnpack build -d myapp            # 生成 fpk
```

## 使用说明
1. 上传 Excel 数据源（第一行为表头）
2. 上传 Word 模板，用 {{列名}} 标记占位符
3. 配置表头行、数据范围、每份行数、文件名模板
4. 生成后可单独下载或打包下载
