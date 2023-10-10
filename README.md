# 分布式Redis-中文书Pdg文件合并

### 项目介绍

本项目基于老马的Pdf2Pic 进行开发的自动解压zip并且合并pdg文件的程序。

- **密码来源于网络，本人不承担任何法律责任。**
- **该项目只用于学习交流，不得用于商业用途。**

### 项目结构

```
├── README.md
├── main.py # 主程序
├── requirements.txt    # 依赖
├── Api # 接口
│   ├── pdg_to_pdf.py   # pdg转pdf
├── passwords   # zip文件解压密码
├── Pdg2Pic # pdg合成程序
```

### 使用方法

1. 安装依赖
  - `pip install -r requirements.txt`

2. 运行`main.py`文件

### 功能列表
- [x] 暴力破解密码，zip文件解压
- [x] pdg 文件合并，转换为pdf文件
- [x] 合并进度监控，避免程序卡死
- [x] 清除意外程序，避免无法进行获取组件的问题
- [x] 增加pdf 清晰化调整，降低pdf文件大小（慎用）


### 注意事项

- 因为需要使用pywinauto，在点击组件或者是选择文件夹的时候，鼠标和键盘不能进行操作，远程桌面退出时不能关闭远程桌面程序。
- 绝大多数文件都带有加密，第一页有可能没有密码，但是其他文件存在密码，所以我们应该检查前三页文件尝试进行密码破解。
- 压缩包中的文件存有pdg文件、png文件、pdf文件、pdf有可能存在密码或者是损坏，需要进行检查。
- 合并pdg文件的时候需要设置添加目录，否则会导致后期无法进行计算第一页的开始页码。

### 参考项目
- [zip2pdf](https://github.com/Davy-Zhou/zip2pdf)
- [老马原创空间](https://www.cnblogs.com/stronghorse/p/14594337.html)
- [全国图书馆参考咨询联盟](http://www.ucdrs.superlib.net/)

**如果对您有帮助，欢迎star，谢谢！**