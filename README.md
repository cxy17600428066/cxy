# cxy

当前仓库用于保存商品内容平台与商品手册相关的原型文件、预览脚本和导出素材。

## 目录说明

- `全渠道商品内容平台-PC工作台.html`
  PC 工作台主页面原型。
- `商品手册移动端预览-弹窗详情.html`
  移动端预览页面原型。
- `product-manual-preview.html`
  商品手册预览页面。
- `verify-step4.html`
  第 4 步校验页面。
- `preview-server.js`
  Node 版本地静态预览服务。
- `preview-server.ps1`
  PowerShell 版本地静态预览服务。
- `*.png`
  页面检查截图。
- `脱骨侠产品手册.pdf`
  产品手册 PDF 原文件。
- `外包商品列表.xls`
  商品列表 Excel 文件。

## 本地预览

### 方式一：Node

```powershell
node .\preview-server.js
```

打开浏览器访问：

```text
http://127.0.0.1:8123/
```

### 方式二：PowerShell

```powershell
powershell -ExecutionPolicy Bypass -File .\preview-server.ps1
```

## Git 使用

当前仓库远程地址：

```text
git@github.com:cxy17600428066/cxy.git
```

首次同步完成后，常用流程有两种：

### 查看状态

```powershell
git status
```

### 手动提交并推送

```powershell
git add .
git commit -m "更新说明"
git push
```

### 一键提交并推送

```powershell
.\git-quick.ps1 "更新说明"
```

如果不传提交说明，脚本会自动使用当前时间生成默认提交信息。

### 本地 Git 快捷别名

仓库会配置一个本地别名，之后也可以直接这样用：

```powershell
git quick "更新说明"
```

## 忽略规则

仓库已忽略以下本地运行目录，避免提交浏览器缓存：

- `.edge-headless/`
- `.edge-profile/`
