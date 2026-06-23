# -*- coding: utf-8 -*-
"""
C 盘空间分析工具
=================
- 扫描 C 盘各目录大小，找出空间占用大户
- 检测 AppData 缓存（剪映、美图、WPS 等）
- 生成带饼图的 Excel 报告 + 清理建议
- 可选一键清理 npm / pip 缓存

运行方式:
  python c_disk_analyzer.py
  或双击 dist/ 下的 exe
"""

import os
import sys
import subprocess
import threading
from datetime import datetime


def is_admin():
    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def fmt(b):
    if b is None or b < 0:
        return "N/A"
    if b >= 1e9:
        return f"{b/1e9:.2f} GB"
    if b >= 1e6:
        return f"{b/1e6:.2f} MB"
    if b >= 1e3:
        return f"{b/1e3:.2f} KB"
    return f"{int(b)} B"


def dir_size(path, timeout=30, depth=0):
    if not os.path.exists(path):
        return None
    res = [None]

    def work():
        try:
            total = 0
            base = path.rstrip("\\").count("\\")
            for root, dirs, files in os.walk(path):
                if depth > 0 and root.count("\\") - base > depth:
                    dirs.clear()
                    continue
                for f in files:
                    try:
                        fp = os.path.join(root, f)
                        if os.path.exists(fp):
                            total += os.path.getsize(fp)
                    except Exception:
                        pass
            res[0] = total
        except Exception:
            pass

    t = threading.Thread(target=work, daemon=True)
    t.start()
    t.join(timeout)
    return res[0]


def scan():
    data = {"total": 0, "used": 0, "free": 0, "folders": [], "appdata": []}

    # 磁盘容量
    try:
        import ctypes
        fb = ctypes.c_ulonglong(0)
        tb = ctypes.c_ulonglong(0)
        ctypes.windll.kernel32.GetDiskFreeSpaceExW(
            "C:\\", None, ctypes.byref(tb), ctypes.byref(fb)
        )
        data["total"] = tb.value
        data["free"] = fb.value
        data["used"] = data["total"] - data["free"]
    except Exception:
        pass

    print(f"\n总容量: {fmt(data['total'])}")
    print(f"已用:   {fmt(data['used'])} ({data['used']/data['total']*100:.1f}%)")
    print(f"可用:   {fmt(data['free'])}")

    # 一级目录
    print("\n扫描一级目录...")
    skip = {
        "System Volume Information", "Recovery", "$Recycle.Bin",
        "Windows.old", "Config.Msi", "Documents and Settings"
    }
    try:
        for item in os.scandir("C:\\"):
            if not item.is_dir(follow_symlinks=False):
                continue
            if item.name in skip:
                continue
            sz = dir_size(item.path, timeout=35)
            if sz and sz > 0:
                data["folders"].append({"name": item.name, "size": sz})
    except Exception:
        pass
    data["folders"].sort(key=lambda x: x["size"], reverse=True)

    # AppData
    print("\n扫描 AppData 缓存...")
    local = os.environ.get("LOCALAPPDATA", "")
    if local and os.path.exists(local):
        for item in os.scandir(local):
            if item.is_dir():
                sz = dir_size(item.path, timeout=12, depth=3)
                if sz and sz > 100e6:
                    data["appdata"].append({"name": item.name, "size": sz})
    data["appdata"].sort(key=lambda x: x["size"], reverse=True)
    return data


def gen_excel(data, path):
    # 延迟导入，避免没装 openpyxl 时直接报错
    import openpyxl
    from openpyxl.chart import PieChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import DataPoint
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = openpyxl.Workbook()
    hf = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    hfill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    df = Font(name="微软雅黑", size=10)
    tf = Font(name="微软雅黑", bold=True, size=14, color="2F5496")
    stf = Font(name="微软雅黑", bold=True, size=12, color="C00000")
    rf = Font(name="微软雅黑", size=10, bold=True, color="FF0000")
    gf = Font(name="微软雅黑", size=10, bold=True, color="00B050")
    bd = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    gfill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    rfill = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")
    ofill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
    yfill = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")

    def shdr(ws, hd, wd):
        for i, (h, w) in enumerate(zip(hd, wd), 1):
            c = ws.cell(row=1, column=i, value=h)
            c.font = hf; c.fill = hfill
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = bd
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    def wr(ws, r, data, dc=None):
        for ci, v in enumerate(data, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font = df
            c.alignment = Alignment(vertical='center', wrap_text=True)
            c.border = bd
            if ci == dc:
                c.alignment = Alignment(horizontal='center', vertical='center')
                if v == "能":
                    c.fill = gfill
                elif v == "不能":
                    c.fill = rfill
                elif v == "部分能":
                    c.fill = ofill
                else:
                    c.fill = yfill

    # === Sheet 1: C盘目录总览 ===
    ws1 = wb.active
    ws1.title = "C盘目录总览"
    shdr(ws1, ["目录路径", "说明", "大小", "能否删除", "备注"], [50, 22, 14, 12, 55])
    rows = []
    for f in data["folders"]:
        n = f["name"].lower()
        if "windows" in n:
            cd, nt = "不能", "系统核心不要动"
        elif "program files" in n:
            cd, nt = "不能", "通过设置卸载勿手动删"
        elif "programdata" in n:
            cd, nt = "不能", "删了导致软件异常"
        elif n == "users":
            cd, nt = "视情况", "桌面/文档是个人文件"
        elif n in ("tmp", "temp"):
            cd, nt = "能", "临时文件安全可删"
        else:
            cd, nt = "视情况", "自行判断"
        rows.append([f["name"], f["name"], fmt(f["size"]), cd, nt])
    for i, r in enumerate(rows, 2):
        wr(ws1, i, r, dc=4)
    ws1.auto_filter.ref = f"A1:E{len(rows)+1}"
    ws1.freeze_panes = "A2"

    # === Sheet 2: AppData缓存 ===
    ws2 = wb.create_sheet("AppData缓存")
    shdr(ws2, ["软件", "类型", "大小", "能否删除", "建议"], [30, 22, 15, 12, 55])
    known = {
        "jianyingpro": ("剪映专业版", "视频缓存", "部分能", "剪映->设置->清缓存"),
        "meituapp": ("美图", "图片缓存", "部分能", "美图->设置->清缓存"),
        "meitu": ("美图", "图片缓存", "部分能", "美图->设置->清缓存"),
        "wechat": ("微信开发者工具", "开发工具", "部分能", "工具->设置->清缓存"),
        "kingsoft": ("WPS Office", "办公缓存", "部分能", "WPS->设置->清缓存"),
        "doubao": ("豆包", "AI缓存", "部分能", "豆包->设置->清缓存"),
        "microsoft": ("Microsoft", "系统组件", "不能", "系统相关不要动"),
        "jetbrains": ("JetBrains IDE", "IDE缓存", "部分能", "File->Invalidate Caches"),
        "feishu": ("飞书", "通讯缓存", "部分能", "飞书->设置->清缓存"),
        "google": ("Google/Chrome", "浏览器", "部分能", "Chrome->设置->清数据"),
        "qianniu": ("千牛", "卖家工具", "部分能", "千牛->设置->清缓存"),
        "pip": ("pip缓存", "Python包", "能", "pip cache purge"),
        "dingtalk": ("钉钉", "通讯缓存", "部分能", "钉钉->设置->清缓存"),
        "openai": ("OpenAI", "AI工具", "部分能", "工具->设置->清缓存"),
        "npm": ("npm缓存", "Node.js包", "能", "npm cache clean --force"),
        "mozilla": ("Firefox", "浏览器", "部分能", "Firefox->设置->清数据"),
    }
    rows2 = []
    for a in data["appdata"]:
        kl = a["name"].lower()
        mk = None
        for k, (n, t, cd, act) in known.items():
            if k in kl:
                mk = (n, t, cd, act)
                break
        if mk:
            rows2.append([mk[0], mk[1], fmt(a["size"]), mk[2], mk[3]])
        else:
            rows2.append([a["name"], "应用缓存", fmt(a["size"]), "部分能", "软件设置里清"])
    for i, r in enumerate(rows2, 2):
        wr(ws2, i, r, dc=4)
    ws2.auto_filter.ref = f"A1:E{len(rows2)+1}"
    ws2.freeze_panes = "A2"

    # === Sheet 3: 空间总结 + 饼图 + 建议 ===
    ws3 = wb.create_sheet("空间总结")
    fg = data["free"] / 1e9
    ug = data["used"] / 1e9
    tg = data["total"] / 1e9
    ws = sum(f["size"] for f in data["folders"] if "windows" in f["name"].lower()) / 1e9
    p64 = sum(f["size"] for f in data["folders"] if f["name"] == "Program Files") / 1e9
    p86 = sum(f["size"] for f in data["folders"] if f["name"] == "Program Files (x86)") / 1e9
    pd = sum(f["size"] for f in data["folders"] if f["name"] == "ProgramData") / 1e9
    at = sum(a["size"] for a in data["appdata"]) / 1e9

    cats = []
    if ws > 0:
        cats.append(("Windows系统", round(ws, 1)))
    if at > 0:
        cats.append(("AppData缓存", round(at, 1)))
    if p64 > 0:
        cats.append(("Program Files(64)", round(p64, 1)))
    if p86 > 0:
        cats.append(("Program Files(32)", round(p86, 1)))
    if pd > 0:
        cats.append(("ProgramData", round(pd, 1)))
    oth = max(0, ug - ws - p64 - p86 - pd - at)
    if oth > 1:
        cats.append(("其他系统文件", round(oth, 1)))
    cats.append(("可用空间", round(fg, 1)))

    for ri, row in enumerate([["分类", "大小(GB)"]] + [[n, s] for n, s in cats], 1):
        for ci, v in enumerate(row, 1):
            c = ws3.cell(row=ri, column=ci, value=v)
            if ri == 1:
                c.font = hf; c.fill = hfill
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = bd
            else:
                c.font = df; c.border = bd
    ws3.column_dimensions['A'].width = 28
    ws3.column_dimensions['B'].width = 15

    # 饼图
    pie = PieChart()
    pie.title = f"C盘{tg:.0f}GB空间分布"
    pie.style = 10
    pie.width = 22
    pie.height = 15
    labels = Reference(ws3, min_col=1, min_row=2, max_row=len(cats) + 1)
    dr = Reference(ws3, min_col=2, min_row=1, max_row=len(cats) + 1)
    pie.add_data(dr, titles_from_data=True)
    pie.set_categories(labels)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showCatName = True
    pie.dataLabels.showLeaderLines = True
    for i, c in enumerate([
        "4472C4", "5B9BD5", "ED7D31", "A5A5A5", "FFC000",
        "FF6999", "264478", "70AD47", "BFBFBF"
    ]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = c
        pie.series[0].data_points.append(pt)
    ws3.add_chart(pie, "D1")

    # 清理建议
    rs = len(cats) + 4
    ws3.merge_cells(f"A{rs}:B{rs}")
    ws3.cell(row=rs, column=1, value="C盘空间清理建议").font = tf

    r = rs + 1
    ws3.merge_cells(f"A{r}:B{r}")
    ws3.cell(row=r, column=1, value="一、能马上清理的(安全)").font = stf
    for i, (n, a) in enumerate([
        ("npm缓存", "npm cache clean --force"),
        ("pip缓存", "pip cache purge"),
        ("临时文件", "Win+R->cleanmgr->C盘->清理系统文件")
    ], r + 1):
        ws3.merge_cells(f"A{i}:B{i}")
        ws3.cell(row=i, column=1, value=f"  [{i-r}] {n}").font = gf
        ws3.merge_cells(f"C{i}:E{i}")
        ws3.cell(row=i, column=3, value=a).font = Font(name="微软雅黑", size=10)

    r2 = r + 4
    ws3.merge_cells(f"A{r2}:B{r2}")
    ws3.cell(row=r2, column=1, value="二、需在软件内清理的缓存").font = stf
    big = [a for a in data["appdata"] if a["size"] > 0.5e9][:8]
    for i, a in enumerate(big, r2 + 1):
        ws3.merge_cells(f"A{i}:B{i}")
        ws3.cell(
            row=i, column=1,
            value=f"  [{i-r2}] {a['name']}({fmt(a['size'])})"
        ).font = Font(name="微软雅黑", size=10, bold=True, color="ED7D31")
        ws3.merge_cells(f"C{i}:E{i}")
        ws3.cell(row=i, column=3, value="软件设置里清理").font = Font(
            name="微软雅黑", size=10
        )

    r3 = r2 + len(big) + 2
    ws3.merge_cells(f"A{r3}:B{r3}")
    ws3.cell(row=r3, column=1, value="三、绝对不能碰的目录").font = stf
    for i, item in enumerate([
        "C:\\Windows系统核心",
        "C:\\ProgramData程序数据",
        "Program Files走设置卸载勿手动删",
        "pagefile.sys虚拟内存自动管理"
    ], r3 + 1):
        ws3.merge_cells(f"A{i}:E{i}")
        ws3.cell(row=i, column=1, value=f"  {item}").font = rf

    r4 = r3 + 6
    ws3.merge_cells(f"A{r4}:B{r4}")
    ws3.cell(row=r4, column=1, value="四、日后注意事项").font = stf
    for i, tip in enumerate([
        "1. 剪映等视频软件做完后立即清理缓存",
        "2. 大文件放到D盘或其他盘",
        "3. 微信/钉钉关掉自动下载",
        "4. 每月跑一次cleanmgr",
        "5. 终极方案:换大容量SSD"
    ], r4 + 1):
        ws3.merge_cells(f"A{i}:E{i}")
        ws3.cell(row=i, column=1, value=f"  {tip}").font = Font(
            name="微软雅黑", size=10
        )

    wb.save(path)
    return path


def main():
    print("=" * 60)
    print("  C 盘空间分析工具 v1.0")
    print("  正在扫描，请稍候...")
    print("=" * 60)
    if not is_admin():
        print("\n建议以管理员身份运行以获得完整结果\n")

    data = scan()

    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    xp = os.path.join(desktop, f"C盘分析报告_{ts}.xlsx")

    try:
        gen_excel(data, xp)
        print(f"\nExcel 报告已生成: {xp}")
    except Exception as e:
        print(f"\n生成 Excel 失败: {e}")

    # 处理双击运行无 stdin 的情况
    def safe_input(prompt):
        try:
            return input(prompt).strip().lower()
        except (RuntimeError, EOFError):
            return "n"

    print("\n是否清理以下缓存?")
    r = safe_input("  清理 npm 缓存? (y/n): ")
    if r == "y":
        subprocess.run(["npm", "cache", "clean", "--force"],
                       capture_output=True, timeout=30)
        print("  npm 缓存已清理")
    r = safe_input("  清理 pip 缓存? (y/n): ")
    if r == "y":
        subprocess.run([sys.executable, "-m", "pip", "cache", "purge"],
                       capture_output=True, timeout=30)
        print("  pip 缓存已清理")

    print("\n完成!")
    try:
        input("按 Enter 退出...")
    except (RuntimeError, EOFError):
        pass


if __name__ == "__main__":
    main()
