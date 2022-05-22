import os


def del_file(path: str, deep: int = -1):
    filename_list = ("3光滑度",)
    if deep > 2:
        return False

    if deep == -1:
        deep = 0
    else:
        deep += 1

    print(path)
    if not os.path.exists(path):
        return False
    if os.path.isfile(path):
        name = path[path.rfind("\\") + 1:path.rfind(".")]
        if name in filename_list:
            os.remove(path)
            return True
        return False
    if os.path.isdir(path):
        if not path[len(path) - 1] == "\\":
            path += "\\"
        folder = path[:path.rfind("\\") + 1]
        for path in os.listdir(path):
            del_file(folder + path, deep)
    else:
        return False

if __name__ == '__main__':
    path = input("path: ")
    del_file(path)



