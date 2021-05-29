import re

cell_tar = '呼吸机： 1次（见面 0次；电话 1次）   '
cell_tar = ''
result = re.findall(r"\d+", cell_tar)[0]
print(result)