# Link Bind Office Plugin

这是一个基于 Microsoft Office Add-in 框架开发的 PowerPoint 插件，提供链接绑定和管理功能。使用 React + TypeScript 构建，支持任务窗格、命令执行和对话框交互。 

## 项目结构

```
.
├── src/                          # 源代码目录
│   ├── commands/                 # Ribbon 命令处理
│   │   ├── commands.html         # 命令页面 HTML
│   │   └── commands.ts           # 命令逻辑实现
│   ├── dialog/                   # 对话框组件
│   │   ├── DialogApp.tsx         # 对话框 React 组件
│   │   ├── dialog.html           # 对话框 HTML
│   │   └── index.tsx             # 对话框入口文件
│   ├── resource-picker/          # 资源选择器
│   │   ├── ResourcePickerApp.tsx # 资源选择器 React 组件
│   │   ├── index.tsx             # 资源选择器入口
│   │   └── resource-picker.html  # 资源选择器 HTML
│   ├── shared/                   # 共享模块
│   │   ├── config.ts             # 配置文件
│   │   ├── linkBinder.ts         # 链接绑定核心逻辑
│   │   └── types.ts              # TypeScript 类型定义
│   └── taskpane/                 # 任务窗格
│       ├── components/           # React 组件
│       │   ├── App.tsx           # 主应用组件
│       │   ├── Header.tsx        # 头部组件
│       │   └── LinkBinderPanel.tsx # 链接绑定面板
│       ├── index.tsx             # 任务窗格入口
│       └── taskpane.html         # 任务窗格 HTML
├── manifest.json                 # Office Add-in 清单 (JSON 格式)
├── manifest.xml                  # Office Add-in 清单 (XML 格式)
├── webpack.config.js             # Webpack 打包配置
├── tsconfig.json                 # TypeScript 配置
├── babel.config.json             # Babel 转译配置
├── package.json                  # 项目依赖和脚本
└── README.md                     # 项目说明文档
```

## 技术栈

- **前端框架**: React 18
- **UI 库**: Fluent UI React Components v9
- **语言**: TypeScript 5.4+
- **构建工具**: Webpack 5, Babel 7
- **样式**: Less 4.2
- **Office API**: Office.js
- **Polyfills**: core-js, es6-promise (支持 IE 11)

## 快速开始

### 前置要求

- Node.js (LTS 版本)
- npm 或 yarn
- Microsoft Office (PowerPoint，支持加载项的版本)

### 安装依赖

```bash
npm install
```

### 开发模式

启动本地 HTTPS 开发服务器并侧载插件：

```bash
npm start
```

此命令会：
1. 生成自签名证书（首次运行需要信任）
2. 启动 Webpack Dev Server (HTTPS)
3. 自动打开 PowerPoint 并加载插件

### 仅启动开发服务器

如果只需要启动开发服务器而不自动打开 Office：

```bash
npm run dev-server
```

### 构建生产版本

```bash
npm run build
```

构建后的文件输出到 `dist/` 目录。

### 其他常用命令

- `npm run watch` - 监听文件变化并重新构建
- `npm run validate` - 验证 manifest.xml 配置
- `npm run lint` - 代码检查
- `npm run lint:fix` - 自动修复代码问题
- `npm run prettier` - 格式化代码
- `npm run stop` - 停止调试会话

## 核心功能模块

### 1. Task Pane (任务窗格)
- **位置**: `src/taskpane/`
- **功能**: 主界面侧边栏，提供链接绑定的可视化操作界面
- **主要组件**:
  - `LinkBinderPanel.tsx`: 链接绑定操作面板
  - `Header.tsx`: 顶部导航和标题

### 2. Commands (命令处理)
- **位置**: `src/commands/`
- **功能**: 处理 Ribbon 菜单按钮点击事件
- **用途**: 执行后台操作，如插入文本、触发特定功能

### 3. Dialog (对话框)
- **位置**: `src/dialog/`
- **功能**: 弹出式对话框，用于用户交互和确认操作
- **特点**: 独立的 React 应用，通过 Office Dialog API 调用

### 4. Resource Picker (资源选择器)
- **位置**: `src/resource-picker/`
- **功能**: 资源选择和绑定界面
- **用途**: 帮助用户从列表中选择要绑定的资源

### 5. Shared (共享模块)
- **位置**: `src/shared/`
- **功能**: 跨模块共享的工具和配置
- **核心文件**:
  - `linkBinder.ts`: 链接绑定核心业务逻辑
  - `types.ts`: TypeScript 类型定义
  - `config.ts`: 全局配置常量

## 部署说明

### 为什么适合放到 GitHub Pages

Office Add-in 要求入口页、命令页和图标资源走 HTTPS。这个项目本身是纯前端静态站点，构建产物都在 `dist/`，所以非常适合直接部署到 GitHub Pages。

### GitHub Pages 部署步骤

1. 把仓库推到 GitHub。
2. 进入仓库 `Settings -> Pages`。
3. 在 `Build and deployment` 中把 `Source` 设为 `GitHub Actions`。
4. 如果你使用默认 Pages 地址，站点会发布到：

```text
https://<github-user>.github.io/<repo-name>/
```

5. 如果你使用自定义域名，请在仓库 `Settings -> Secrets and variables -> Actions -> Variables` 里新增：

```text
ADDIN_BASE_URL=https://你的域名或Pages地址
```

例如：

```text
ADDIN_BASE_URL=https://plugins.example.com/link-bind
```

6. 推送到 `main` 或 `master` 后，仓库里的 `.github/workflows/deploy-pages.yml` 会自动构建并发布 `dist/`。

### 本地构建 GitHub 发布版

如果你想先在本地生成一份发布版：

```bash
ADDIN_BASE_URL=https://<github-user>.github.io/<repo-name> npm run build
```

构建完成后：

- 页面文件在 `dist/`
- 生产环境 manifest 在 `dist/manifest.xml`

注意：给 PowerPoint 侧载时，使用的是 `dist/manifest.xml`，不是项目根目录下开发用的 `manifest.xml`。

### manifest 与域名规则

- `SourceLocation`、`Commands.Url`、图标地址会自动替换成 `ADDIN_BASE_URL`
- `AppDomain` 会自动取这个地址的 origin

例如当：

```text
ADDIN_BASE_URL=https://evdo.github.io/link-bind
```

生成结果会是：

- 页面地址：`https://evdo.github.io/link-bind/taskpane.html`
- 命令页：`https://evdo.github.io/link-bind/commands.html`
- 域名白名单：`https://evdo.github.io`

### 侧载到 Office 时需要做什么

1. 等 GitHub Pages 发布成功。
2. 下载或直接使用仓库构建产物里的 `dist/manifest.xml`。
3. 在 PowerPoint 中通过加载项侧载这个 manifest。
4. 之后 Office 会从你的 HTTPS 站点加载插件页面。

## 开发规范

- 遵循 ESLint 规则 (`eslint-plugin-office-addins`)
- 代码格式化使用 Prettier
- TypeScript 开启严格模式
- 组件采用函数式组件 + Hooks
- 使用 Fluent UI 保持 Office 风格一致性

## 常见问题

### 1. 证书信任问题
首次运行时会提示安装证书，请选择"是"以信任 localhost 证书。

### 2. 端口冲突
默认使用端口 3000，如被占用可通过环境变量调整。

### 3. IE 11 兼容性
项目配置了完整的 Polyfill 支持 IE 11，但建议优先在现代浏览器中调试。

## 许可证

MIT

## 相关链接

- [Office Add-ins 官方文档](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Fluent UI React](https://react.fluentui.dev/)
- [GitHub 仓库](https://github.com/himotou/office-plugins.git)
