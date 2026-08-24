# 幸运观众抽奖

一个可直接部署到 GitHub Pages 的单页抽奖工具。名单只在浏览器本地读取和处理，不会上传。

## 使用方式

1. 准备 `.xlsx`、`.xls` 或 `.csv` 名单，第一行建议使用 `姓名` 和 `学号` 作为列名。
2. 打开网页，选择或拖入名单文件。
3. 设置本轮抽取人数，点击“开始抽奖”。默认会在下一轮排除已中奖观众。

如果没有表头，程序将第一列当作姓名、第二列当作学号。重复学号和空白姓名会被自动忽略。

可参考 [名单模板.csv](template/名单模板.csv)。用 Excel 打开并另存为 `.xlsx` 后即可导入。

## 本地预览

直接双击 `index.html` 即可使用。Excel 解析组件已随项目保存，因此导入名单无需网络连接。

## 部署到 GitHub Pages

1. 在 GitHub 新建仓库（例如 `lucky-audience-lottery`）。
2. 将本项目全部文件推送到仓库的默认分支。
3. 进入仓库的 **Settings → Pages**，选择 **Deploy from a branch**，并选择默认分支的根目录 (`/`)。
4. 保存后，GitHub 会显示访问链接。

## 项目结构

```text
.
├── assets/
│   ├── app.js
│   └── styles.css
├── template/名单模板.csv
└── index.html
```
