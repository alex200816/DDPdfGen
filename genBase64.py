import base64

# 图标路径
icon_path = "DDlogo.ico"
#导出的base64路径
gen_path = "dd.py"

open_icon = open(icon_path, "rb")
b64str = base64.b64encode(open_icon.read())
open_icon.close()
write_data = "img=%s" % b64str
f = open(gen_path, "w+")
f.write(write_data)
f.close()