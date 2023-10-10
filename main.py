#!/usr/bin/env python
# -*- coding:utf-8 -*-
'''
# @FileName  :main.py
# @Time      :2023/9/3 21:49
# @Author    :William
# desc       :
'''
import os
import shutil
import subprocess
from loguru import logger
from Api.pdg_to_pdf import brute_force_zip_password,zipunpack,pdg2pdf,check_pdf

# 创建临时文件夹，用于存放临时文件夹，和pdf 文件
base_path = os.path.dirname(os.path.abspath(__file__))
temp_path = os.path.join(base_path, "temp")
if not os.path.exists(temp_path):
    os.mkdir(temp_path)
password_path = os.path.join(base_path, "passwords", "passwords.txt")
zip_7z_path = os.path.join(base_path, "7-Zip", "7z.exe")
pdg_exe_path = os.path.join(base_path, "Pdg2Pic", "Pdg2Pic.exe")


# 关闭意外的闲置任务进程
def close_exe(process_name:str):
    try:
        # 使用taskkill命令终止进程
        subprocess.run(['taskkill', '/F', '/IM', process_name])
        print(f"进程 {process_name} 已终止")
    except Exception as e:
        print(f"无法终止进程 {process_name}: {e}")


def main(ssn,zip_path,del_zip=True):
    '''
    :param ssn: SSN号
    :param zip_path: zip文件路径
    '''
    try:
        process_list = ["Pdg2Pic_PDG转PDF工具.exe", "Pdg2Pic.exe"]
        [close_exe(x) for x in process_list]

        # 将文件拷贝解压到临时文件夹
        old_zip_path = zip_path
        new_zip_path = os.path.join(temp_path, f"{ssn}.{old_zip_path.rsplit('.', 1)[-1]}")
        logger.debug(f"旧zip路径：{old_zip_path}，新zip路径：{new_zip_path}")
        shutil.copyfile(old_zip_path, new_zip_path)
        assert os.path.exists(new_zip_path), f"{ssn} 复制zip文件失败"

        output_dir = os.path.join(temp_path, ssn)  # zip 输出文件夹
        pdf_path = str(output_dir) + ".pdf"  # pdf 文件路径

        zip_passwd = brute_force_zip_password(new_zip_path, password_path)  # 获取zip 文件密码
        assert zip_passwd, "未找到密码"

        unpack_res = zipunpack(new_zip_path, output_dir, zip_7z_path, zip_passwd)  # 解压zip文件
        assert unpack_res, "解压失败"

        pdg2pdf(pdg_exe_path, output_dir)
        assert os.path.exists(pdf_path), "pdf文件合并失败！"

        pdf_status = check_pdf(pdf_path)
        assert pdf_status == 2, "pdf文件损坏或者未合并完成"

        # 只有所有步骤都成功，才会删除原始zip文件
        if del_zip:
            os.remove(old_zip_path) # 删除原始zip文件

    except AssertionError as e:
        logger.error(f"错误：{e}")
    except Exception as e:
        logger.error(f"未知错误：{e}")
    finally:
        # 因为有时候pdf 合并失败，所以需要删除临时文件夹
        try:
            os.remove(new_zip_path)
            shutil.rmtree(output_dir)
        except Exception as e:
            logger.error(f"删除临时文件失败：{e}")

if __name__ == '__main__':
    zip_path = r"D:\Desktop\独秀\做人生路上的“不倒翁”_13594976.zip"
    SSN = "13594976"
    main(SSN,zip_path,del_zip=False)
