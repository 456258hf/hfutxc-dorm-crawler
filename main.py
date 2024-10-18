"""爬取 hfutxc 寝室卫生检查系统——查询宿舍床铺评分的数据"""
import tkinter as tk
from tkinter import ttk
import sys
from ast import literal_eval
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
FACULTIES = ("生态", "文法", "英语", "城市", "电气", "计算",
             "经济", "能源", "食品", "物流", "材料", "机械")

# 年级
GRADES = ("21", "22", "23", "24")

# 层号范围，默认为range(1, 7)
FLOORS = range(1, 7)

# 房间号范围，默认为range(1, 41)
ROOMS = range(1, 41)


class FeetToMeters:
    """创建主窗口"""

    def __init__(self, root):
        self.root = root
        root.geometry('800x600')
        root.title('hfutxc-dorm-crawler')

        # 定义变量
        self.faculty = tk.StringVar(value="机械")  # 院系
        self.grade = tk.StringVar(value="22")  # 年级
        self.building = tk.StringVar(value="9N")  # 楼栋
        self.req = tk.BooleanVar(value=False)  # 是否获取
        self.full_mode = tk.StringVar(value="以楼栋各自生成")  # 全校生成模式
        self.delay = tk.DoubleVar(value=0.01)  # 请求间隔
        self.timeout = tk.DoubleVar(value=10)  # 请求超时
        self.csv = tk.BooleanVar(value=True)  # 是否生成csv
        self.xlsx = tk.BooleanVar(value=True)  # 是否生成xlsx
        self.dict_xlsx_dorm = tk.BooleanVar(value=True)  # 是否生成寝室字典xlsx
        self.dict_xlsx_faculty = tk.BooleanVar(value=True)  # 是否生成院系年级字典xlsx

        # 目标学期，格式为(年,学期序号)，单个学期需在tuple后打,
        self.year_term_index = tk.StringVar(value="((24, 1),)")

        # 创建 Notebook 小部件
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(pady=10, expand=True, fill='both')

        # 创建五个 Frame 作为选项卡内容
        # frame1
        frame1 = self.create_frame1()
        frame2 = self.create_frame2()
        frame3 = self.create_frame3()
        frame4 = self.create_frame4()
        frame5 = self.create_frame5()

        # frame1.pack(fill='both', expand=True)
        # frame2.pack(fill='both', expand=True)
        # frame3.pack(fill='both', expand=True)
        # frame4.pack(fill='both', expand=True)
        # frame5.pack(fill='both', expand=True)

        # 将 Frame 添加到 Notebook 中
        self.notebook.add(frame1, text='院系年级')
        self.notebook.add(frame2, text='楼栋')
        self.notebook.add(frame3, text='全校')
        self.notebook.add(frame4, text='更新字典')
        self.notebook.add(frame5, text='配置')

        text_widget = tk.Text(root, wrap='word', width=30)
        text_widget.pack(expand=True, fill='both')
        # sys.stdout = PrintRedirector(text_widget)

    def create_frame1(self):
        """创建frame1"""
        frame = ttk.Frame(self.notebook, padding="10 10 12 12")
        frame.grid(column=4, row=3)

        frame.columnconfigure(4, weight=1)
        frame.rowconfigure(3, weight=1)

        # row 1
        ttk.Label(frame, text="院系：").grid(column=1, row=1)
        faculty_slection = ttk.Combobox(
            frame, textvariable=self.faculty, width=5)
        faculty_slection['values'] = FACULTIES
        faculty_slection.state(["readonly"])
        faculty_slection.grid(column=2, row=1)

        ttk.Label(frame, text="年级：").grid(column=3, row=1)
        grade_selection = ttk.Combobox(frame, textvariable=self.grade, width=5)
        grade_selection['values'] = GRADES
        grade_selection.grid(column=4, row=1)

        # row 2
        ttk.Checkbutton(frame, text='在线获取',  variable=self.req,
                        onvalue=True, offvalue=False).grid(column=1, row=2)
        ttk.Button(frame, text="执行", command=self.start1).grid(
            column=4, row=2)

        for child in frame.winfo_children():
            child.grid_configure(padx=5, pady=5)
        return frame

    def create_frame2(self):
        """创建frame2"""
        frame = ttk.Frame(self.notebook, padding="10 10 12 12")
        frame.grid(column=4, row=3)

        frame.columnconfigure(4, weight=1)
        frame.rowconfigure(3, weight=1)

        # row 1
        ttk.Label(frame, text="楼栋：").grid(column=1, row=1)
        building_slection = ttk.Combobox(
            frame, textvariable=self.building, width=5)
        building_slection['values'] = BUILDINGS
        building_slection.state(["readonly"])
        building_slection.grid(column=2, row=1)

        # row 2
        ttk.Checkbutton(frame, text='在线获取',  variable=self.req,
                        onvalue=True, offvalue=False).grid(column=1, row=2)
        ttk.Button(frame, text="执行", command=self.start2).grid(
            column=4, row=2)

        for child in frame.winfo_children():
            child.grid_configure(padx=5, pady=5)
        return frame

    def create_frame3(self):
        """创建frame3"""
        frame = ttk.Frame(self.notebook, padding="10 10 12 12")
        frame.grid(column=4, row=3)

        frame.columnconfigure(4, weight=1)
        frame.rowconfigure(3, weight=1)

        # row 1
        ttk.Label(frame, text="模式：").grid(column=1, row=1)
        faculty_slection = ttk.Combobox(
            frame, textvariable=self.full_mode, width=20)
        faculty_slection['values'] = ("以楼栋各自生成", "以院系年级各自生成", "合并生成")
        faculty_slection.state(["readonly"])
        faculty_slection.grid(column=2, row=1)

        # row 2
        ttk.Checkbutton(frame, text='在线获取',  variable=self.req,
                        onvalue=True, offvalue=False).grid(column=1, row=2)
        ttk.Button(frame, text="执行", command=self.start3).grid(
            column=4, row=2)

        for child in frame.winfo_children():
            child.grid_configure(padx=5, pady=5)
        return frame

    def create_frame4(self):
        """创建frame4"""
        frame = ttk.Frame(self.notebook, padding="10 10 12 12")
        frame.grid(column=4, row=3)

        frame.columnconfigure(4, weight=1)
        frame.rowconfigure(3, weight=1)

        # row 1
        ttk.Label(frame, text="注：会遍历全校寝室，请仅在必要时使用").grid(
            column=1, row=1, columnspan=2)
        ttk.Checkbutton(frame, text='生成寝室字典xlsx',  variable=self.dict_xlsx_dorm,
                        onvalue=True, offvalue=False).grid(column=3, row=1)
        ttk.Checkbutton(frame, text='生成院系年级字典xlsx',  variable=self.dict_xlsx_faculty,
                        onvalue=True, offvalue=False).grid(column=4, row=1)

        # row 2
        ttk.Checkbutton(frame, text='在线获取',  variable=self.req,
                        onvalue=True, offvalue=False).grid(column=1, row=2)
        ttk.Button(frame, text="执行", command=self.start4).grid(
            column=4, row=2)

        for child in frame.winfo_children():
            child.grid_configure(padx=5, pady=5)
        return frame

    def create_frame5(self):
        """创建frame5"""
        frame = ttk.Frame(self.notebook, padding="10 10 12 12")
        frame.grid(column=4, row=3)

        frame.columnconfigure(4, weight=1)
        frame.rowconfigure(3, weight=1)

        # row 1
        ttk.Label(frame, text="学期：").grid(column=1, row=1)
        ttk.Entry(frame, textvariable=self.year_term_index).grid(
            column=2, row=1)

        # row 2
        ttk.Checkbutton(frame, text='执行时生成csv',  variable=self.csv,
                        onvalue=True, offvalue=False).grid(column=1, row=2)
        ttk.Checkbutton(frame, text='执行时生成xlsx',  variable=self.xlsx,
                        onvalue=True, offvalue=False).grid(column=2, row=2)

        for child in frame.winfo_children():
            child.grid_configure(padx=5, pady=5)
        return frame

    def start1(self):
        """执行指定的操作 院系年级"""
        dorms = FACULTY_DICT[self.faculty.get()+self.grade.get()]
        if self.req.get():
            unsuccessful_dorms = dorm_request.dorms_req(dorms)
            while unsuccessful_dorms:  # 循环获取直至无失败寝室
                unsuccessful_dorms = dorm_request.dorms_req(
                    unsuccessful_dorms)
        dorm_decode.dorms_dec(self.faculty.get()+self.grade.get(), dorms,
                              literal_eval(self.year_term_index.get()), self.csv.get(), self.xlsx.get())

    def start2(self):
        """执行指定的操作 楼栋"""
        dorms = DORM_DICT[self.building.get()]
        if self.req.get():
            unsuccessful_dorms = dorm_request.dorms_req(dorms)
            while unsuccessful_dorms:  # 循环获取直至无失败寝室
                unsuccessful_dorms = dorm_request.dorms_req(
                    unsuccessful_dorms)
        dorm_decode.dorms_dec(self.building.get(), dorms,
                              literal_eval(self.year_term_index.get()), self.csv.get(), self.xlsx.get())

    def start3(self):
        """执行指定的操作 全校"""
        if self.full_mode.get() == "以楼栋各自生成":
            for building in BUILDINGS:
                dorms = DORM_DICT[building]
                if self.req.get():
                    unsuccessful_dorms = dorm_request.dorms_req(dorms)
                    while unsuccessful_dorms:  # 循环获取直至无失败寝室
                        unsuccessful_dorms = dorm_request.dorms_req(
                            unsuccessful_dorms)
                dorm_decode.dorms_dec(
                    building, dorms, literal_eval(self.year_term_index.get()), self.csv.get(), self.xlsx.get())
        elif self.full_mode.get() == "以院系年级各自生成":
            for faculty in FACULTIES:
                for grade in GRADES:
                    faculty_grade = faculty+grade
                    dorms = FACULTY_DICT[faculty_grade]
                    if self.req.get():
                        unsuccessful_dorms = dorm_request.dorms_req(dorms)
                        while unsuccessful_dorms:  # 循环获取直至无失败寝室
                            unsuccessful_dorms = dorm_request.dorms_req(
                                unsuccessful_dorms)
                    dorm_decode.dorms_dec(faculty_grade, dorms, literal_eval(self.year_term_index.get(
                    )), self.csv.get(), self.xlsx.get())
        elif self.full_mode.get() == "合并生成":
            dorms = []
            for building in BUILDINGS:
                dorms.extend(DORM_DICT[building])
            if self.req.get():
                unsuccessful_dorms = dorm_request.dorms_req(dorms)
                while unsuccessful_dorms:  # 循环获取直至无失败寝室
                    unsuccessful_dorms = dorm_request.dorms_req(
                        unsuccessful_dorms)
            dorm_decode.dorms_dec(
                "all", dorms, literal_eval(self.year_term_index.get()), self.csv.get(), self.xlsx.get())

    def start4(self):
        """执行指定的操作 更新字典"""
        dorms = []
        for buingding in BUILDINGS:
            for floor in FLOORS:
                for room in ROOMS:
                    dorms.append(f"{buingding}{floor}{room:02d}")
        if self.req.get():
            unsuccessful_dorms = dorm_request.dorms_req(dorms)
            while unsuccessful_dorms:  # 循环获取直至无失败寝室
                unsuccessful_dorms = dorm_request.dorms_req(unsuccessful_dorms)
        dorm_dict_gen.dorm_dict_genf(
            BUILDINGS, FLOORS, ROOMS, self.dict_xlsx_dorm.get())
        faculty_dict_gen.faculty_dict_genf(BUILDINGS, FLOORS, ROOMS,
                                           self.dict_xlsx_faculty.get())


class PrintRedirector:
    """重定向print()"""

    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, text):
        """print"""
        self.text_widget.insert(tk.END, text)
        self.text_widget.see(tk.END)  # 自动滚动到最后一行

    def flush(self):
        """需要实现flush方法以兼容Python的标准输出"""


if __name__ == "__main__":
    # 运行主循环
    roottk = tk.Tk()
    app = FeetToMeters(roottk)
    roottk.mainloop()
