# xls2lua

首先申明，第一份文件是fanlix在2008年1月25日放在csdn上分享给大家的[1]，后面由于个人需求修改了一小部分内容

这是一个用python语言写的，excel转lua工具。方便设计人员使用excel填写配置表，导出lua文件作为最终的游戏配置文件

# 准备

* 安装python

* 安装xlrd
[下载地址](https://pypi.python.org/pypi/xlrd)


# 使用

* 命令：python xls2lua.py src_dir dst_dir
默认使用当前路径下的src和dst目录

* 把所有的excel文件拷贝到src_dir目录

* 运行python xls2lua.py src_dir dst_dir

# 说明

* 默认从excel第三行开始解析

# 参考

[1][csdn](http://blog.csdn.net/bhwst/article/details/6778978)