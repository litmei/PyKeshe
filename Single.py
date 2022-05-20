#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Author : 柠檬味_LMei - QQ
# @Email : mccllm@qq.com
# @Time : 2022/05/17
# @Detailed : 将原项目多文件（见“原项目图.png”）整合进一个文件中，同时实现识别文件、录入文件、计算文件、输出文件等功能

import os
import sys
import json
import xlwings as xw
import numpy as np
import re
import matplotlib.pyplot as plt
import matplotlib
import scipy.interpolate as spi


# 原ks3分文件内容 - 用于绘制拟合曲线并输出图片文件
class Painter:

	__data_label = (("算数平均值", "加权平均值", "标准差", "光滑度", "三阶中心矩", "信息熵"),  # title
					("T(K)", "T(K)", "温度T(K)", "", "T^3(K^3)", ""))  # y label
	__abc = "abcdef"  # 序号
	__st_fp = ""  # 文件夹路径
	__y_value = []
	__x_value = []
	__x_label = ""
	__o_file_black_list = []

	# 录入曲线图数据
	@staticmethod
	# filepath-文件夹路径 count-同文件夹下第几个 x-数据文件名（仅数字） ys-y值打包 xl-x标签
	def st_enter_line_data(filepath: str, count: int, x: float, ys: list, xl: str):
		if count == 1:  # 清空原数据
			if Painter.__st_fp != "":
				print("程序代码逻辑有误，请检查！")
				return False
			else:
				Painter.__st_fp = filepath
				Painter.__x_label = xl
		Painter.__y_value.append(ys)
		Painter.__x_value.append(x)
		return True

	# 画拟合曲线图
	@staticmethod
	def st_draw_line_img(config):
		# 0. 读入配置
		Painter.__o_file_black_list = config["输出图片黑名单"]
		suffix: str = config["输出图片后缀"]

		# 1. 判断数据是否支持绘图
		if len(Painter.__data_label[0]) != len(Painter.__y_value):
			if len(Painter.__x_value) <= 1:
				print("录入数据过少，无法绘制！")
				return

		# 2. 绘图
		plt.rcParams['font.family'] = ['SimSun']  # 解决 中文字体 无法显示问题
		plt.rcParams['figure.figsize'] = [9, 5]
		plt.rcParams['font.size'] = 18
		plt.rcParams['axes.unicode_minus'] = False  # 解决 减号 无法显示问题
		my_font = matplotlib.font_manager.FontProperties(family="SimSun", size=18)

		x_sorted = []
		y_sorted = []
		ii = 0
		y_value = []
		count = 0
		while True:
			if Painter.__data_label[0][ii] in Painter.__o_file_black_list:  # 不输出黑名单图片，重复运行删除对应图片
				try:
					os.remove(Painter.__st_fp + str(Painter.__data_label[0][ii]) + suffix)
					print(Painter.__data_label[0][ii] + "（旧文件）被成功移除")
				except Exception as result:
					# print(result)
					pass
				ii += 1
				if ii >= len(Painter.__y_value[0]):
					break
				continue

			x_sorted.clear()
			y_sorted.clear()
			y_value.clear()

			print(Painter.__data_label[0][ii])  # 显示当前正在处理哪个图片

			for num in Painter.__y_value:
				y_value.append(num[ii])
			points = dict(zip(Painter.__x_value, y_value))

			for x in sorted(points.keys()):
				x_sorted.append(x)
				y_sorted.append(points[x])

			x_list = np.linspace(x_sorted[0], x_sorted[len(x_sorted) - 1], 100)  # 待求插值的位置

			func1 = spi.interp1d(x_sorted, y_sorted, kind="cubic")  # 拟合函数

			plt.plot(x_sorted, y_sorted, "ro", label="样本点")  # 画点
			plt.plot(x_list, func1(x_list), "b", label="三次样条插值")  # 画线
			plt.plot(x_sorted, y_sorted, "r--", label="线性插值")
			plt.legend(prop=my_font)  # 显示标签
			plt.grid()  # 网格
			plt.title("({0}){1}".format(Painter.__abc[count], Painter.__data_label[0][ii]), size=20, y=-0.3)  # 标题
			plt.xlabel(Painter.__x_label, x=0.93)  # x轴标签
			plt.ylabel(Painter.__data_label[1][ii], y=0.9)  # y轴标签
			plt.subplots_adjust(bottom=0.4, left=0.3, right=0.95, top=0.95)
			plt.xticks(rotation=45)
			count += 1

			#plt.show()  # 展示
			plt.savefig(Painter.__st_fp + str(count) + Painter.__data_label[0][ii] + suffix)  # 输出文件
			plt.clf()  # 清空画布

			ii += 1
			if ii >= len(Painter.__y_value[0]):
				break

		# 3. 清空数据
		Painter.__st_fp = ""
		Painter.__y_value.clear()
		Painter.__x_value.clear()
		Painter.__x_label = ""


# 原ks2分文件内容 - 频数统计分段及计算各参数（平均值，标准差，光滑度，三阶矩，信息熵）
class CalculateTool:

	# 带参初始化
	def __init__(self, filepath: str, cfg_path: str, file_count: int, app: xw.App):
		self.filepath = filepath  # 路径备份
		self.folder_path = filepath[0:filepath.rfind("\\")]  # 文件夹路径
		self.filename = filepath[filepath.rfind("\\") + 1:filepath.rfind(".")]  # 文件名

		config = json.load(open(cfg_path, encoding="utf-8"))
		self.dis_num = config["分段数量"]  # 分段数量
		self.o_file_name = config["导出文件名"]

		self.file_count = file_count
		self.app = app

		raw_nums = np.loadtxt(filepath, comments="%")  # 导入文件中的数据，忽略“%”开头的注释行
		temp_nums = raw_nums[:, 2]  # 拿出数据温度列的数据
		self.nums = np.sort(temp_nums)  # 从小到大排序后的数据列
		del raw_nums, temp_nums
		self.elem_counts = np.shape(self.nums)[0]  # 数据个数

		# 数据声明
		self.mean = self.average = self.std = self.smoothness = self.c_moment3 = self.entropy = 0.0
		self.mid_pair = {}

	# 算数平均值
	def get_mean(self):
		return self.mean

	def __cal_mean(self):
		self.mean = np.mean(self.nums)

	# 加权平均值
	def get_average(self):
		return self.average

	def __cal_average(self):
		self.average = np.average(self.nums)

	# 标准差
	def get_std(self):
		return self.std

	def __cal_std(self):
		self.std = np.std(self.nums)

	# 光滑度
	def get_smoothness(self):
		return self.smoothness

	def __cal_smoothness(self):
		self.smoothness = 1 - 1 / (1 + self.std ** 2)

	# 三阶中心距
	def get_c_moment3(self):
		return self.c_moment3

	def __cal_c_moment3(self):
		self.c_moment3 = np.sum((self.nums - self.mean) ** 3) / self.elem_counts

	# 三阶矩
	def get_moment3(self):
		return self.c_moment3 ** (1 / 3)

	# 频数分布
	def get_fd_data(self):
		return self.mid_pair

	def __do_fd(self):
		min_val = self.nums[0]  # 最小值
		max_val = self.nums[self.elem_counts - 1]  # 最大值
		dis_val = (max_val - min_val) / self.dis_num  # 每段长度
		mid_val = min_val + dis_val / 2  # 中值
		mid_count = 0  # 中值计数
		counts = 0  # 总计数
		next_num = min_val + dis_val  # 右边界值
		while True:
			num = self.nums[counts]
			if num < next_num:  # 当前比较数字 小于 当前区间右边界
				mid_count += 1
				counts += 1
			else:  # 大于等于右边界
				self.mid_pair[mid_val] = mid_count
				mid_val += dis_val
				mid_count = 0
				next_num += dis_val
				if next_num >= max_val:  # 如果下一阶段为最后一阶段
					mid_count = self.elem_counts - counts
					self.mid_pair[mid_val] = mid_count
					break

	# 信息熵
	def get_entropy(self):
		return self.entropy

	def __cal_entropy(self):
		for value in self.mid_pair.values():
			value /= self.elem_counts
			if value == 0:
				continue
			self.entropy += value * np.log2(value)
		self.entropy = -self.entropy

	# 一轮计算
	def calculate(self):
		print("C -> " + self.filename)
		self.__cal_mean()
		self.__cal_average()
		self.__cal_std()
		self.__cal_smoothness()
		self.__cal_c_moment3()
		self.__do_fd()
		self.__cal_entropy()

	# 输出到文件
	def output_file(self):
		wb = None
		try:
			if self.file_count == 1:
				wb = self.app.books.add()
				wb.sheets[0].name = "频数统计"
				wb.sheets.add("计算结果")
			else:
				wb = self.app.books.active
		except Exception as result:
			print("打开时出现问题：" + result.__str__())

		try:
			st = wb.sheets["计算结果"]
			if self.file_count == 1:
				st.range((1, 1)).value = ["名称", "算数平均值", "加权平均值", "标准差", "光滑度", "三阶中心矩", 
										  "信息熵", "", "频数统计分段数"]
				st.range((2, 9)).value = self.dis_num
			st.range((self.file_count + 1, 1)).value = [self.filename, self.mean, self.average, self.std,
														self.smoothness, self.c_moment3, self.entropy]

			st = wb.sheets["频数统计"]
			st.range((self.file_count * 3 - 2, 1)).value = [self.filename]
			st.range((self.file_count * 3 - 1, 1)).options(transpose=True).value = self.mid_pair
		except Exception as result:
			print("输入数据时出现问题：" + result.__str__())

		# 导出图片 -- 录入本次计算结果
		filename_num = float(self.filename)
		ys = [self.mean, self.average, self.std, self.smoothness, self.c_moment3, self.entropy]
		xl = self.folder_path[self.folder_path.rfind("\\") + 1:len(self.folder_path)]
		xl = re.sub(r"^\d*", "", xl)  # 正则，去除文件夹名称中的数字
		Painter.st_enter_line_data(self.folder_path + "\\", self.file_count, filename_num, ys, xl)


# 原ks1分文件内容 - 文件管理系统：找到指定路径下应该去计算的文件
class FileControlSystem:
	def __init__(self, filepath: str, cfg_path: str):
		self.filepath = filepath
		self.cfg_path = cfg_path
		self.config = json.load(open(cfg_path, encoding="utf-8"))
		self.suffix = self.config["允许后缀"]
		self.black_list = self.config["过滤文件名"]
		self.o_f_name = self.config["导出文件名"]
		self.count = 0
		self.app = xw.App(visible=False, add_book=False)
		self.wb = None

	# 开始工作
	def begin(self):
		self.__work()
		self.app.quit()  # 处理完成，保存并退出

	# 主要工作，计算文件，输出文件
	def __work(self, filepath: str = "", deep: int = -1):
		d = deep  # -1：初始值 0：零层文件夹 1：一层文件夹 2：二层文件夹
		if d > 2:  # 深度大于2，该文件夹下内容不读取
			return False
		if filepath == "":
			filepath = self.filepath
		if os.path.isfile(filepath):  # 是单个文件
			pos0 = filepath.rfind("\\")
			pos1 = filepath.rfind(".")
			pos2 = len(filepath)
			if filepath[pos1:pos2] == self.suffix:  # 后缀名相符
				if filepath[pos0 + 1:pos1] in self.black_list:  # 文件名在黑名单中
					return False
				self.count += 1
				o = CalculateTool(filepath, self.cfg_path, self.count, self.app)
				o.calculate()
				o.output_file()
				return True
		if os.path.isdir(filepath):  # 是文件夹
			print(filepath)
			path_list = os.listdir(filepath)  # 获取当前文件夹下文件列表
			if len(path_list) == 0:
				return False
			if filepath[len(filepath) - 1] != "\\":
				filepath += "\\"
			self.__clear_obsolete_files(filepath)  # 每进入一个新文件夹，清除一次旧文件
			self.count = 0  # 进入新文件夹要清零
			b = False
			for name in path_list:  # 递归
				if self.__work(filepath + name, d + 1):
					b = True
			if b:  # 保存xlsx时需要执行的操作
				try:
					self.wb = self.app.books.active
					self.wb.save(filepath + self.o_f_name + ".xlsx")
					self.wb.close()
				except Exception as result:
					print(result)
					return False
			if b:
				Painter.st_draw_line_img(self.config)
			return True

	# 清空之前的计算结果文件
	def __clear_obsolete_files(self, folder_path: str):
		try:
			path = folder_path + self.o_f_name + ".xlsx"
			os.remove(path)
		except FileNotFoundError:
			pass


# 获取要处理的根目录
def __get_path():
	filepath = input("输入文件或文件夹路径来开始：\n")
	while True:
		if not os.path.exists(filepath):  # 检测是否为路径
			if len(filepath) < 2:
				break
			# 去除两侧符号后再检测一次（比如去除双引号），Python可以使用连续相等来判断
			if filepath[0] == filepath[len(filepath) - 1] == "\"":
				filepath = filepath[1:len(filepath) - 1]
				if os.path.exists(filepath):
					break
			filepath = input("路径不存在，请重试：\n")
		else:
			break
	return filepath


# 读配置文件
def __read_cfg():
	try:
		config = json.load(open(sys.path[0] + "\\config.json", encoding="utf-8"))
		if len(config.__str__()) <= __sizeof_min_cfg():
			raise Exception
	except FileNotFoundError:
		__write_def_cfg()
		print("未找到配置文件，已生成并使用默认配置文件“config.json“到执行文件同级目录：\n" + sys.path[0])
	except (json.decoder.JSONDecodeError, Exception):
		__write_def_cfg()
		print("配置文件有误，已生成并使用默认配置文件“config.json”到执行文件同级目录：\n" + sys.path[0])
	return sys.path[0] + "\\config.json"


# 写默认配置文件
def __write_def_cfg():
	config = {"分段数量": 100, "允许后缀": ".txt", "导出文件名": "处理结果", "过滤文件名": ["Calculation", "Distribution"],
			  "输出图片黑名单": ["加权平均值"], "输出图片后缀": ".png"}
	with open(sys.path[0] + "\\config.json", "w") as f:
		json.dump(config, f, indent=4, ensure_ascii=False)


# 获取最小配置文件大小
def __sizeof_min_cfg():
	config = {"分段数量": None, "允许后缀": None, "导出文件名": None, "过滤文件名": [None], "输出图片黑名单": [None],
			  "输出图片后缀": None}
	return len(config.__str__())


# 展示菜单
def show_menu():
	print("*" * 30)
	print("课设 - Python 版本")
	print("author：LMei")
	print("*" * 30)
	filepath = __get_path()
	config = __read_cfg()
	return filepath, config


# main
if __name__ == '__main__':
	filepath, cfg_path = show_menu()
	fcs = FileControlSystem(filepath, cfg_path)
	fcs.begin()
	print("数据已处理完成")
