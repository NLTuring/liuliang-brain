# 留量大脑 · 协作经纬台 (Rural E-Commerce & Supply Chain Governance Workbench)

![JavaScript](https://img.shields.io/badge/Frontend-Vanilla%20JS%20%2B%20Vite%206-F7DF1E?logo=javascript&logoColor=black)
![UI](https://img.shields.io/badge/Design-Editorial%20Grid%20%26%20Data%20Viz-1E293B)

面向县域农产品上行与乡村电商协同治理的**流量沉淀、供应链排产与多主体分润核算交互工作台**。

---

## 核心特性

- **流量转化与复购沉淀看板**
  - 追踪公域直播/短视频助农流量向县域私域池的留存转化率，提供批次订单追踪与长尾复购分析。
- **县域供应链协同与产能调度**
  - 连接农户采收、合作社分级分拣、冷链物流与售后质检节点，避免助农爆单导致的履约违约与品控滑坡。
- **可视化台账与交互推演**
  - 采用高对比度纸感网格视觉设计（`#F5F2EA` 主题），支持订单台账检索、异常节点诊断弹窗与治理策略模拟。

---

## 快速启动

环境要求：Node.js 18+。

```bash
npm install
npm run dev
```

启动后浏览器访问终端提示的本地地址即可进入**留量大脑 · 协作经纬台**。

### 构建生产版本

```bash
npm run build
npm run preview
```

---

## 目录结构

```text
liuliang-brain/
├── src/            # 工作台前端核心交互逻辑 (main.js) 与网格样式体系 (style.css)
├── figures/        # 系统架构与机制演化图集
├── index.html      # 应用入口
└── vite.config.js  # Vite 构建配置
```
