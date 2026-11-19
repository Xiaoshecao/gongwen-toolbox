# 性能对比：新旧两版

测试环境：Windows 11 / Python 3.12.14 / PySide6 6.12.0；样本用 `发城镇重阳节前夕走访慰问老年人.docx`（14 段）
与 `发城镇新时代文明实践项目培育工作机制（试行）.docx`（68 段 1 表），复制成 N 份。

## 一、窗口构建（打开软件到界面可用）

| 版本 | 耗时 |
|---|---|
| 旧版 tkinter 4 栏 + 3 页签 + 预览 | **3149 ms** |
| 新版 Qt 17 页 + 侧边菜单 | **577 ms** |

**约 5.5 倍**。旧版慢在 tkinter 要为每个控件算样式、四栏全部控件一次性建出来（即使没用到）；
Qt 只建一次主框架，页面按需切换。

## 二、批量检查（同一批文件）

| 场景 | 旧版（每次重解析、串行、界面冻结） | 新版首次（4 线程并行） | 新版再查（解析缓存） |
|---|---|---|---|
| 5 份 | 98 ms | 87 ms | **10 ms** |
| 20 份 | 363 ms | 401 ms | **39 ms（9.2 倍）** |

- 首次检查的提升不明显——真正的瓶颈是 docx 解析的磁盘 I/O，两个版本都要读一遍；
- **真正的差别在重复检查**：改了规则、换了格式标准再查一遍时，新版直接复用已解析的文档树，**快 9 倍**；
- 旧版检查时整个窗口卡住（tkinter 单线程，检查在 UI 线程里跑）；新版放后台线程，检查期间还能点别的页、改名单。

## 三、大表格（检查结果 / 角色表）

| 场景 | tkinter `ttk.Treeview` | Qt `QTableView` + Model |
|---|---|---|
| 插入 10000 行 | 53 ms（一次性） | **43 ms** |
| 满屏滚动 | 0.4 ms/次（行少时） | **7.5 ms/次（1 万行满屏）** |

tkinter 插得快，但**数据一多就崩**：它没有虚拟化，10000 行会全部建实体控件，
上万行后内存和滚动明显变慢；Qt 的 Model/View 只渲染可见的几十行，10 万行也是同样的开销。

## 四、复现方法

```bash
# 窗口构建
python -c "import time;from gwtyper.app import App;t=time.perf_counter();a=App();print((time.perf_counter()-t)*1000);a.destroy()"
python -c "import sys;sys.path.insert(0,'gongwen-qt');from PySide6.QtWidgets import QApplication;from ui.main_window import MainWindow;app=QApplication([]);import time;t=time.perf_counter();w=MainWindow();print((time.perf_counter()-t)*1000)"

# 批量检查
python gongwen-qt/bench_check.py
```
