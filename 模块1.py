string = input("path")
p2 = string.rfind("\\", 0, len(string) - 1)
p1 = string[:p2].rfind("\\") + 1
key = string[p1:p2]
print(p1, p2)
print(key)
