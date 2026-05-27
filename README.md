# 参数适用性矩阵生成器

从利率装配源表自动生成参数适用性矩阵。

## 使用方式

### 方式一：直接运行 EXE（Windows）

1. 从 [Releases](../../releases) 下载最新 `参数适用性矩阵生成器.exe`
2. 双击运行
3. 在弹出的文件夹选择对话框中选择利率装配源表文件夹
4. 转换结果输出到 EXE 同级的 `output/` 文件夹

### 方式二：Python 运行

```bash
pip install -r requirements.txt
python generate_param_applicability.py
```

## 源表结构要求

| 行号 | 内容 |
|------|------|
| Row 2 | 产品名称（合并单元格） |
| Row 7 | 事件类型（正常 / 逾期 / 挪用） |
| Row 11 | 表头行，Col D = "字段名称"，含"启用标识"列 |
| Row 12+ | 参数明细（Col A = 一级分类，Col B = 二级分类，Col D = 字段名称） |

## 输出文件

| 文件 | 说明 |
|------|------|
| `参数适用性_最终矩阵.xlsx` | 产品 × 参数 适用性矩阵（✅/❌/—） |
| `参数适用性_最终矩阵_核对记录.xlsx` | 核对记录（汇总、分类补齐审计、多数映射、格式校验、来源明细） |

## 单元格含义

| 符号 | 含义 |
|------|------|
| ✅ | 启用 |
| ❌ | 不启用 / 不适用 |
| — | 源产品未覆盖该参数 |

## 构建 EXE

项目使用 GitHub Actions 自动构建。推送 tag 即可发布 Release：

```bash
git tag v1.0
git push origin v1.0
```

本地构建：
```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --name 参数适用性矩阵生成器 generate_param_applicability.py
```
