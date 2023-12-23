# -*- coding: UTF-8 -*-

# Python3.x 导入方法
import tkinter
from tkinter import ttk
from tkinter.filedialog import askopenfile

from pass_traversal import IntPassFeature, StrPassFeature, PassTraverser, ExcelUnlock


class UserInterface:
    def __init__(self):
        self.root = tkinter.Tk()  # 创建窗口对象的背景色
        self.excel_path = tkinter.StringVar()
        self.pass_count = tkinter.IntVar()
        self.row_index = 4
        self.pass_features_list = []

        root = self.root
        tkinter.Label(root, text="目标路径:").grid(row=0, column=0)
        tkinter.Entry(root, textvariable=self.excel_path).grid(row=0, column=1)
        tkinter.Button(root, text="路径选择", command=self.select_path).grid(row=0, column=2)

        self.progressbar = ttk.Progressbar(root)  # tkinter.Progressbar(root)
        self.progressbar.grid(row=1, column=1)

        tkinter.Button(root, text="破解", command=self.traversal_pass).grid(row=1, column=2)

        tkinter.Label(root, text="可能的密码数量（估算）:").grid(row=2, column=0)
        tkinter.Entry(root, textvariable=self.pass_count).grid(row=2, column=1)
        tkinter.Button(root, text="添加密码节", command=self.add_pass_node).grid(row=3, column=0)

    def traversal_pass(self):
        self.progressbar['value'] = 0
        feature_list = []
        pass_count = 1
        for item in self.pass_features_list:
            if item['pass_type'].get() == '数字型':
                step = 1
                try:
                    step = int(item['added_info'].get())
                except:
                    pass
                feature = IntPassFeature(item['min_value'].get(), item['max_value'].get(), step)
            else:
                chars = list(set(list(item['added_info'].get())))
                feature = StrPassFeature(chars, item['min_value'].get(), item['max_value'].get())
            pass_count *= feature.count()
            feature_list.append(feature)
        self.pass_count.set(pass_count)
        self.progressbar['maximum'] = pass_count
        pass_word = self.traversal_excel_pass(self.excel_path.get(), feature_list)
        if pass_word is None:
            tkinter.messagebox.showinfo(title="破解失败", message='破解失败')
        else:
            tkinter.messagebox.showinfo(title="破解成功", message=f'密码为：{pass_word}')


    def traversal_excel_pass(self, path: str, feature_list: list):
        pass_traversal = PassTraverser(feature_list)
        g = pass_traversal.get_pass_generator()
        i = 0
        log_list = []
        for item in g:
            i += 1
            log_list.append(item)
            if i % 100 == 0:
                self.root.update()
                with open('./pass.log', 'a', encoding='utf-8') as fp:
                    fp.write('\n'.join(log_list))
                log_list = []
            # print(item)
            self.progressbar['value'] += 1
            if ExcelUnlock.deciphering_execl(item, path):
                print('密码是：', item)
                return item

    def select_path(self):
        path_ = askopenfile()
        self.excel_path.set(path_.name)

    def add_pass_node(self):
        pass_type = tkinter.StringVar()
        pass_type.set('字符型')
        min_value = tkinter.IntVar()
        max_value = tkinter.IntVar()
        added_info = tkinter.StringVar()
        pass_feature = {
            'pass_type': pass_type,
            'min_value': min_value,
            'max_value': max_value,
            'added_info': added_info,
        }
        self.pass_features_list.append(pass_feature)
        root = self.root
        self.row_index += 1
        tkinter.Label(root, text="密码类型:").grid(row=self.row_index, column=0)
        tkinter.OptionMenu(root, pass_type, "数字型", "字符型").grid(row=self.row_index, column=1)
        tkinter.Label(root, text="最小值/最小长度:").grid(row=self.row_index, column=2)
        tkinter.Entry(root, textvariable=min_value).grid(row=self.row_index, column=3)
        tkinter.Label(root, text="最大值/最大长度:").grid(row=self.row_index, column=4)
        tkinter.Entry(root, textvariable=max_value).grid(row=self.row_index, column=5)
        tkinter.Label(root, text="步长/字符集:").grid(row=self.row_index, column=6)
        tkinter.Entry(root, textvariable=added_info).grid(row=self.row_index, column=7)


if __name__ == '__main__':
    uif = UserInterface()
    uif.root.mainloop()  # 进入消息循环
