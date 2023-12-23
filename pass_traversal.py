# 字符 数字
# 字符 关键点： 最小长度、最大长度；深度优先遍历、广度优先遍历
# 数字 关键点： 最大值、最小值
from abc import abstractmethod
from collections import deque

import pywintypes
import win32com.client


class PassFeature:
    def __init__(self):
        pass

    @abstractmethod
    def get_pass_generator(self):
        pass

    @abstractmethod
    def count(self):
        pass


class StrPassFeature(PassFeature):
    def __init__(self, pass_range: list, min_len: int, max_len: int, traversal_type=1):
        """
        字符型密码遍历
        :param pass_range: 字符范围
        :param min_len: 最小长度
        :param max_len: 最大长度
        :param traversal_type: 遍历方式
        """
        self.pass_range = pass_range
        self.min_len = max(min_len, 0)
        self.max_len = max_len
        self.traversal_type = traversal_type
        self.pass_que = None

    def get_pass_generator(self):
        if self.min_len == 0:
            self.pass_que = deque([''])
        else:
            self.pass_que = deque(self.pass_range)
        while True:
            _pass = self.pass_que.popleft()
            if len(_pass) < self.max_len:
                self.pass_que.extend([_pass + _i for _i in self.pass_range])
            while len(_pass) < self.min_len:
                _pass = self.pass_que.popleft()
                if len(_pass) < self.max_len:
                    self.pass_que.extend([_pass + _i for _i in self.pass_range])
            yield _pass
            if len(self.pass_que) == 0:
                return

    def count(self):
        count = 0
        _l = self.min_len
        while _l <= self.max_len:
            count += len(self.pass_range) ** _l
            _l += 1
        return count


class IntPassFeature(PassFeature):
    def __init__(self, min_value: int, max_value: int, step: int = 1):
        self.min_value = min_value
        self.max_value = max_value
        self.step = step

    def get_pass_generator(self):
        _cur_pass = self.min_value
        while _cur_pass <= self.max_value:
            yield _cur_pass
            _cur_pass += self.step

    def count(self):
        return self.max_value - self.min_value + 1


class PassTraverser:
    def __init__(self, feature_list: list[PassFeature]):
        self.feature_list = feature_list

    def get_pass_generator(self):
        return self._get_pass_generator(0, '')

    def _get_pass_generator(self, feature_index: int, pre_pass: str):
        feature_list = self.feature_list

        pass_feature = feature_list[feature_index]
        generator = pass_feature.get_pass_generator()
        for item in generator:
            _pass = f'{pre_pass}{item}'
            if len(feature_list) == feature_index + 1:
                yield _pass
            else:
                yield from self._get_pass_generator(feature_index + 1, _pass)

    def count(self):
        count = 1
        for item in self.feature_list:
            count *= item.count()
        return count


class ExcelUnlock:
    xlsx = win32com.client.Dispatch('Excel.Application')  # 获得Excel对象

    @staticmethod
    def deciphering_execl(password: str, path: str) -> bool:
        try:
            # print(_pass)
            # 'D:\CodeSpace\python\pythonProject\data\审计中心监察专项事项登记台账.xlsx'
            # D:\CodeSpace\python\pythonProject\data\工作簿1.xlsx
            wb = ExcelUnlock.xlsx.Workbooks.Open(path, False, False, None, Password=password)
            print(f"成功了 密码是:{password}")  # 成功以后则直接跳出
            wb.Close()
            return True
        except pywintypes.com_error as e:
            if '密码' in e.excepinfo[2]:
                return False
            raise e


def traversal_excel_pass(path: str, feature_list: list):
    pass_traversal = PassTraverser(feature_list)
    g = pass_traversal.get_pass_generator()
    for item in g:
        # print(item)
        if ExcelUnlock.deciphering_execl(item, path):
            print('密码是：', item)
            return item


if __name__ == '__main__':
    path = r'D:\CodeSpace\python\pythonProject\data\temp.xlsx'
    f1 = IntPassFeature(18, 20)
    f2 = StrPassFeature(list(set(list('f'))), 0, 1)
    feature_list = [f1, f2]
    traversal_excel_pass(path, feature_list)
