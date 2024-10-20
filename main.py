"""爬取 hfutxc 寝室卫生检查系统——查询宿舍床铺评分的数据"""
import tkinter as tk
from tkinter import ttk
import time
import os

import dorm_request
import dorm_decode
from dorm_dict import DORM_DICT
from faculty_dict import FACULTY_DICT
import dorm_dict_gen
import faculty_dict_gen

# 寝室楼栋
BUILDINGS = ("1N", "1S", "2N", "2S", "3N", "3S", "4N", "4S", "5N",
             "5S", "6N", "6S", "7#", "8#", "9N", "9S", "10N", "10S")

# 院系
FACULTIES = ("生态环境系", "文法系", "英语系", "城市建设工程系",
             "电气与自动化系", "计算机与信息系", "经济与贸易系", "能源化工系",
             "食品科学系", "物流管理系", "材料工程系", "机械工程系")


# 年级
GRADES = ("21", "22", "23", "24")

# 遍历层号范围，默认为range(1, 7)
FLOORS = range(1, 7)

# 遍历房间号范围，默认为range(1, 41)
ROOMS = range(1, 41)


class HfutxcDormCrawler:
    """创建主窗口"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title('hfutxc-dorm-crawler')

        # 各选项卡（上半部分页面）变量
        self.req = tk.BooleanVar(value=True)  # 是否获取
        self.faculty = tk.StringVar(value="机械工程系")  # 院系
        self.grade = tk.StringVar(value="22")  # 年级
        self.building = tk.StringVar(value="9N")  # 楼栋
        self.full_mode = tk.StringVar(value="以楼栋各自生成")  # 全校生成模式
        self.dict_xlsx_dorm = tk.BooleanVar(value=True)  # 是否生成寝室字典xlsx
        self.dict_xlsx_faculty = tk.BooleanVar(value=True)  # 是否生成院系年级字典xlsx

        # 进度显示（下半部分）变量
        self.dorm_sum = tk.IntVar()  # 寝室总数
        self.dorm_current = tk.IntVar()  # 当前寝室序号
        self.dorm_current_name = tk.StringVar(value='当前寝室')  # 当前寝室名称
        self.dorm_counter = tk.StringVar(value='当前序号/总数')  # 当前序号/总数
        self.dorm_percent = tk.DoubleVar(value=0)  # 进度百分比
        self.log = tk.StringVar(value='log')  # 输出
        self.interrupt = False  # 停止标志

        # 设置页变量
        self.csv = tk.BooleanVar(value=True)  # 是否生成csv
        self.xlsx = tk.BooleanVar(value=True)  # 是否生成xlsx
        self.xlsx_open = tk.BooleanVar(value=True)  # 生成后是否自动打开xlsx
        self.delay = tk.DoubleVar(value=0.01)  # 请求间隔，单位ms
        self.timeout = tk.DoubleVar(value=10)  # 请求超时
        self.year = tk.IntVar(value='2024')  # 目标学年
        self.year_term1 = tk.BooleanVar(value=True)  # 目标学年第一学期
        self.year_term2 = tk.BooleanVar(value=False)  # 目标学年第二学期

        # 创建 Notebook 小部件
        self.notebook = ttk.Notebook(root, padding="10 10 12 12")
        self.notebook.grid(pady=10, column=1, row=1)

        # 创建五个 Frame 作为选项卡内容，添加到 Notebook 中
        self.notebook.add(self.create_frame1(), text='院系年级')
        self.notebook.add(self.create_frame2(), text='楼栋')
        self.notebook.add(self.create_frame3(), text='全校')
        self.notebook.add(self.create_frame4(), text='更新字典')
        self.notebook.add(self.create_frame5(), text='设置')

        # 监听选项卡切换事件
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

        self.create_frame0()

    def create_frame0(self) -> None:
        """创建frame0"""
        self.frame = ttk.Frame(self.root, padding="10 10 12 12")
        self.frame.grid(column=1, row=2)

        self.frame.columnconfigure(4, weight=1)
        self.frame.rowconfigure(4, weight=1)

        # row 2
        ttk.Label(self.frame, text="寝室数：").grid(column=1, row=2)
        ttk.Label(self.frame, textvariable=self.dorm_sum).grid(column=2, row=2)
        self.req_cb = ttk.Checkbutton(
            self.frame, text='在线获取', variable=self.req, command=self.req_hint)
        self.req_cb.grid(column=3, row=2)

        self.button = ttk.Button(self.frame, text="开始", command=self.start)
        self.button.grid(column=4, row=2)

        # row3
        self.progress = ttk.Progressbar(
            self.frame, orient=tk.HORIZONTAL, mode='determinate', length='250', variable=self.dorm_percent)
        self.progress.grid(column=1, row=3, columnspan=2)
        ttk.Label(self.frame, textvariable=self.dorm_counter).grid(
            column=3, row=3)
        ttk.Label(self.frame, textvariable=self.dorm_current_name).grid(
            column=4, row=3)

        # row4
        ttk.Label(self.frame, textvariable=self.log).grid(
            column=1, row=4, columnspan=4)

        for child in self.frame.winfo_children():
            child.grid_configure(padx=5, pady=5)

    def create_frame1(self) -> ttk.Frame:
        """创建frame1"""
        frame = ttk.Frame(self.notebook, padding="10 10 12 12")

        frame.columnconfigure(4, weight=1)
        frame.rowconfigure(4, weight=1)

        # row 1
        ttk.Label(frame, text="院系：").grid(column=1, row=1)
        faculty_slection = ttk.Combobox(
            frame, textvariable=self.faculty)
        faculty_slection['values'] = FACULTIES
        faculty_slection.state(["readonly"])
        faculty_slection.grid(column=2, row=1)
        faculty_slection.bind('<<ComboboxSelected>>', self.get_dorm_sum)

        ttk.Label(frame, text="年级：").grid(column=3, row=1)
        grade_selection = ttk.Combobox(frame, textvariable=self.grade)
        grade_selection['values'] = GRADES
        grade_selection.grid(column=4, row=1)
        grade_selection.bind('<<ComboboxSelected>>', self.get_dorm_sum)

        for child in frame.winfo_children():
            child.grid_configure(padx=5, pady=5)
        return frame

    def create_frame2(self) -> ttk.Frame:
        """创建frame2"""
        frame = ttk.Frame(self.notebook, padding="10 10 12 12")

        frame.columnconfigure(4, weight=1)
        frame.rowconfigure(4, weight=1)

        # row 1
        ttk.Label(frame, text="楼栋：").grid(column=1, row=1)
        building_slection = ttk.Combobox(
            frame, textvariable=self.building)
        building_slection['values'] = BUILDINGS
        building_slection.state(["readonly"])
        building_slection.grid(column=2, row=1)
        building_slection.bind('<<ComboboxSelected>>', self.get_dorm_sum)

        for child in frame.winfo_children():
            child.grid_configure(padx=5, pady=5)
        return frame

    def create_frame3(self) -> ttk.Frame:
        """创建frame3"""
        frame = ttk.Frame(self.notebook, padding="10 10 12 12")

        frame.columnconfigure(4, weight=1)
        frame.rowconfigure(4, weight=1)

        # row 1
        ttk.Label(frame, text="模式：").grid(column=1, row=1)
        faculty_slection = ttk.Combobox(
            frame, textvariable=self.full_mode)
        faculty_slection['values'] = ("以楼栋各自生成", "以院系年级各自生成", "合并生成")
        faculty_slection.state(["readonly"])
        faculty_slection.grid(column=2, row=1)
        faculty_slection.bind('<<ComboboxSelected>>', self.get_dorm_sum)
        ttk.Label(frame, text="注：耗时较长，请耐心等待").grid(
            column=3, row=1, columnspan=2)

        for child in frame.winfo_children():
            child.grid_configure(padx=5, pady=5)
        return frame

    def create_frame4(self) -> ttk.Frame:
        """创建frame4"""
        frame = ttk.Frame(self.notebook, padding="10 10 12 12")

        frame.columnconfigure(4, weight=1)
        frame.rowconfigure(4, weight=1)

        # row 1
        ttk.Checkbutton(frame, text='生成寝室字典xlsx',
                        variable=self.dict_xlsx_dorm).grid(column=1, row=1)
        ttk.Checkbutton(frame, text='生成院系年级字典xlsx',
                        variable=self.dict_xlsx_faculty).grid(column=2, row=1)
        ttk.Label(frame, text="注：会遍历全校寝室，请仅在必要时使用").grid(
            column=3, row=1, columnspan=2)

        for child in frame.winfo_children():
            child.grid_configure(padx=5, pady=5)
        return frame

    def create_frame5(self) -> ttk.Frame:
        """创建frame5"""
        frame = ttk.Frame(self.notebook, padding="10 10 12 12")

        frame.columnconfigure(4, weight=1)
        frame.rowconfigure(4, weight=1)

        # row 1
        ttk.Checkbutton(frame, text='生成csv',
                        variable=self.csv).grid(column=1, row=1)
        ttk.Checkbutton(frame, text='生成xlsx',  variable=self.xlsx,
                        command=self.xlsx_changed).grid(column=2, row=1)
        ttk.Checkbutton(frame, text='自动打开xlsx', variable=self.xlsx_open,
                        command=self.xlsx_open_changed).grid(column=3, row=1)

        # row 2
        ttk.Label(frame, text="请求间隔(s)：").grid(column=1, row=2)
        ttk.Entry(frame, textvariable=self.delay).grid(column=2, row=2)
        ttk.Label(frame, text="请求超时(s)：").grid(column=3, row=2)
        ttk.Entry(frame, textvariable=self.timeout).grid(column=4, row=2)

        # row 3
        ttk.Label(frame, text="学年：").grid(column=1, row=3)
        ttk.Spinbox(frame, from_=2024, to=2099,
                    textvariable=self.year).grid(column=2, row=3)
        ttk.Checkbutton(frame, text='第一学期', variable=self.year_term1,
                        command=self.year_term1_changed).grid(column=3, row=3)
        ttk.Checkbutton(frame, text='第二学期', variable=self.year_term2,
                        command=self.year_term2_changed).grid(column=4, row=3)

        for child in frame.winfo_children():
            child.grid_configure(padx=5, pady=5)
        return frame

    def on_tab_change(self, *arg) -> None:
        """切换选项卡时操作"""
        # 更新显示寝室数量
        self.get_dorm_sum()
        # 设置中禁用开始按钮
        current_tab = self.notebook.select()
        tab_text = self.notebook.tab(current_tab, "text")
        if tab_text == "设置":
            self.button.state(['disabled'])
        else:
            self.button.state(['!disabled'])

    def get_dorm_sum(self, *arg) -> None:
        """计算寝室总数并设置为dorm_sum"""
        # 获取当前notebook的tab的text
        current_tab = self.notebook.select()
        tab_text = self.notebook.tab(current_tab, "text")
        if tab_text == "院系年级":
            group = self.faculty.get()[:2]+self.grade.get()
            dorms = FACULTY_DICT[group]
        elif tab_text == "楼栋":
            group = self.building.get()
            dorms = DORM_DICT[group]
        elif tab_text == "全校":
            if self.full_mode.get() == "以楼栋各自生成":
                for building in BUILDINGS:
                    group = building
                    dorms = DORM_DICT[building]
            elif self.full_mode.get() == "以院系年级各自生成":
                for faculty in FACULTIES:
                    for grade in GRADES:
                        group = faculty[:2]+grade
                        dorms = FACULTY_DICT[group]
            elif self.full_mode.get() == "合并生成":
                dorms = []
                for building in BUILDINGS:
                    dorms.extend(DORM_DICT[building])
        elif tab_text == "更新字典":
            dorms = []
            for buingding in BUILDINGS:
                for floor in FLOORS:
                    for room in ROOMS:
                        dorms.append(f"{buingding}{floor}{room:02d}")
        elif tab_text == "设置":
            dorms = []
        self.dorm_sum.set(len(dorms))

    def start(self) -> None:
        """开始指定的操作"""
        # 将按钮更换为停止
        self.button.configure(text='停止获取', command=self.stop)
        self.button.update()  # 刷新显示，避免无响应
        # 获取当前notebook的tab的text
        current_tab = self.notebook.select()
        tab_text = self.notebook.tab(current_tab, "text")
        if tab_text == "院系年级":
            group = self.faculty.get()[:2]+self.grade.get()
            dorms = FACULTY_DICT[group]
            self.dorms_process(group, dorms, self.xlsx_open.get())
        elif tab_text == "楼栋":
            group = self.building.get()
            dorms = DORM_DICT[group]
            self.dorms_process(group, dorms, self.xlsx_open.get())
        elif tab_text == "全校":
            if self.full_mode.get() == "以楼栋各自生成":
                for building in BUILDINGS:
                    group = building
                    dorms = DORM_DICT[building]
                    self.dorms_process(group, dorms, False)
            elif self.full_mode.get() == "以院系年级各自生成":
                for faculty in FACULTIES:
                    for grade in GRADES:
                        group = faculty[:2]+grade
                        dorms = FACULTY_DICT[group]
                        self.dorms_process(group, dorms, False)
            elif self.full_mode.get() == "合并生成":
                group = "全校"
                dorms = []
                for building in BUILDINGS:
                    dorms.extend(DORM_DICT[building])
                self.dorms_process(group, dorms, self.xlsx_open.get())
        elif tab_text == "更新字典":
            group = "全校"
            dorms = []
            for buingding in BUILDINGS:
                for floor in FLOORS:
                    for room in ROOMS:
                        dorms.append(f"{buingding}{floor}{room:02d}")
            if self.req.get():
                self.dorms_req_n(group, dorms)
            dorm_dict_gen.dorm_dict_genf(
                BUILDINGS, FLOORS, ROOMS, self.dict_xlsx_dorm.get())
            faculty_dict_gen.faculty_dict_genf(BUILDINGS, FLOORS, ROOMS,
                                               self.dict_xlsx_faculty.get())
            self.button.configure(text='开始', command=self.start)

    def dorms_req_n(self, group: str, dorms: list) -> bool:
        """请求指定寝室们的成绩数据并保存为htm文件，返回是否全部完成"""
        self.dorm_sum.set(len(dorms))
        for dorm_index, dorm_name in enumerate(dorms, start=1):
            self.dorm_current.set(dorm_index)
            self.dorm_counter.set(
                f"{self.dorm_current.get()}/{self.dorm_sum.get()}")
            self.dorm_current_name.set(dorm_name)
            self.dorm_percent.set(int(100*(dorm_index/len(dorms))))
            time.sleep(self.delay.get())
            self.root.update()  # 刷新显示，避免无响应
            # 中断获取
            if self.interrupt:
                self.interrupt = False
                return False
            # 循环重试
            success = dorm_request.dorm_req(dorm_name, self.timeout.get())
            while not success:
                self.log.set("重试中")
                self.root.update()
                success = dorm_request.dorm_req(dorm_name, self.timeout.get())
            self.log.set(f"获取{group}数据中……")
            self.root.update()
        return True

    def dorms_process(self, group: str, dorms: list, xlsx_open: bool) -> None:
        """处理指定寝室们的成绩数据并保存为csv或xlsx文件"""
        if self.req.get():
            self.dorms_req_n(group, dorms)
        self.log.set(f"处理{group}数据中……")
        self.root.update()
        year_term_index = []
        if self.year_term1.get():
            year_term_index.append((self.year.get(), 1))
        if self.year_term2.get():
            year_term_index.append((self.year.get(), 2))
        dorm_count, week_count, output_filename_xlsx = dorm_decode.dorms_dec(
            group, dorms, year_term_index, self.csv.get(), self.xlsx.get())
        self.log.set(f"处理{group}完成，有效寝室数：{dorm_count}，有效周数：{week_count}")
        # 打开xlsx
        if xlsx_open:
            os.startfile(output_filename_xlsx)
        self.button.configure(text='开始', command=self.start)

    def xlsx_changed(self) -> None:
        """不生成xlsx，禁用自动打开xlsx"""
        if not self.xlsx.get():
            self.xlsx_open.set(False)

    def xlsx_open_changed(self) -> None:
        """不生成xlsx，禁用自动打开xlsx"""
        if not self.xlsx.get():
            self.xlsx_open.set(False)

    def stop(self) -> None:
        """中止操作"""
        self.interrupt = True

    def req_hint(self) -> None:
        """切换在线获取时提示"""
        self.log.set("非在线获取时请先确认已在线获取过")

    def year_term1_changed(self) -> None:
        """避免未选择学期"""
        if not (self.year_term1.get() or self.year_term2.get()):
            self.year_term2.set(True)

    def year_term2_changed(self) -> None:
        """避免未选择学期"""
        if not (self.year_term1.get() or self.year_term2.get()):
            self.year_term1.set(True)


if __name__ == "__main__":
    # 运行主循环
    roottk = tk.Tk()
    app = HfutxcDormCrawler(roottk)
    roottk.mainloop()
