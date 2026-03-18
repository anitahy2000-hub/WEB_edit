# 打印 hello,world!
# print("hello,world!")

# 定义整数列表（Python 中列表用 [] 定义，不需要指定类型）
# words = [0, 1, 2, 3, 4]
# words = ["dada", "grd"]
# # 遍历列表并打印每个元素
# for word in words:
#     print(word)
# 1. 初始化计数器
count = 1

# 2. while 循环：只要 count ≤ 5，就执行循环体
while count <= 5:
    print("当前数字是：", count)  # 打印计数器
    count = count + 1  # 关键：更新计数器（否则会无限循环！）
# true false
# 循环结束后执行的代码
print("循环结束啦！")