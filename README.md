# 期中考试个人看板 · GitHub Pages 部署版

## 仓库结构
```
index.html          网页本体
core.js / app.js    页面逻辑（与 index.html 同目录即可）
data/exam.xlsx      考试数据（管理员每学期覆盖上传这一份文件即可）
```

## 首次部署（纯网页操作，无需 git）
1. GitHub 新建一个仓库（Public）
2. 把本目录全部内容上传（Add file → Upload files，可整批拖入）
3. Settings → Pages → Branch 选 main（或 master）→ Save
4. 约 1 分钟后访问 https://<用户名>.github.io/<仓库名>/

## 每学期更新（运营者唯一要做的动作）
- 打开仓库 → data/ → 把新的监考表改名为 exam.xlsx 覆盖上传 → Commit
- 访客下次打开页面就是新数据（Ctrl+F5 可强制刷新缓存）

## 数据加载机制
- 页面打开后自动 fetch data/exam.xlsx 并在浏览器本地解析，无任何服务器参与
- 若加载失败（文件未上传/格式损坏），页面会提示并自动展示「手动选择本地文件」备用入口
- 注意：仓库是 Public 时，exam.xlsx 任何人都可下载；如需隐藏名单请改用私有部署方案
