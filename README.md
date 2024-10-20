# hfutxc-dorm-crawler

爬取 hfutxc [寝室卫生检查系统](http://39.106.82.121/query)——查询宿舍床铺评分的数据
GUI 绝赞制作中

---

## Requirements:

```bash
pip install -r requirements.txt
```

## Usage:

当前寝室字典生成于 2024-10-18

```bash
py main.py
```

## Todo:

### 优化：

1. ☑️（变相解决） frame 类化
2. ☑️ decode 函数包装
3. ☑️ 院系名称写全
4. ☑️ 优化学期输入方式
5. 全校部分，架构优化，总数、进度条与输出显示优化
6. 合并两个 dict_gen
7. 优化 dict_gen 输出位置
8. decode 流控，添加进度条，使按钮可以停止 decode
9. 优化 faculty_gen 速度

### 添加功能：

1. ☑️ 进度显示与结果输出采用 progressbar
2. ☑️ 完成添加 delay 与 timeout 配置
3. ☑️ 完成寝室数量选定后自动显示
4. ☑️ 按键添加暂停、停止
5. ☑️ 添加生成 xlsx 后自动打开选项
6. 添加配置保存按键
7. 添加切换后自动保存默认值功能
8. 添加 request decode const 等配置的 GUI 编辑功能
9. 添加最大重试次数

### 远期想法：

1. 使用数据库管理
2. 使用 PyQt 再写一份 GUI
3. 使用 treeview 控件或者其他库实现脱离 excel 生成结果

## Bug:

1. ☑️（变相解决） 切换界面时不会自动停止、清除已显示内容
2. ☑️ dict_gen 进度条异常
3. 处理 excel 已打开时的 PermissionError（添加相关log提示）
4. 全校分院系年级、分楼栋时停止异常
