#!/usr/bin/env python
# -*- coding:utf-8 -*-
'''
# @FileName  :main.py
# @Time      :2023/7/30 21:46
# @Author    :William
# desc       :
'''
import io
import os
import time
import fitz
import pypdf
import subprocess
import zipfile
from loguru import logger
from pywinauto import Application
from PIL import Image

from concurrent.futures import ThreadPoolExecutor
t = ThreadPoolExecutor(max_workers=2)


def brute_force_zip_password(zip_file_path, password_file_path, num_files=3):
    '''
    Zip 文件暴力破解密码
    :param zip_file_path: zip文件路径
    :param password_file_path: 密码文件路径
    :param num_files: 验证密码的文件数量
    '''
    # zip文件获取密码
    def try_password_on_first_files(zip_ref, password):
        try:
            # 为防止有些文件未添加密码，所以需要至少尝试前3个文件，进行验证密码是否正确
            file_list = [name for name in zip_ref.namelist() if not name.endswith('/')][:num_files]
            passwd = False
            for e,file in enumerate(file_list):
                try:
                    zip_ref.read(file, pwd=password.encode())
                    passwd = True
                except Exception as e:
                    passwd = False
                    continue
            return passwd
        except RuntimeError:
            return False

    with open(password_file_path, 'r', encoding="utf-8") as f:
        passwords = f.readlines()
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            for password in passwords:
                password = password.strip()  # Remove newline characters
                if try_password_on_first_files(zip_ref, password):
                    return password
    return None

def zipunpack(file_path:str, output_dir:str, exe_path:str, password:str)-> bool:
    '''
    zip解压
    :param file_path: zip文件路径
    :param output_dir: 解压后的文件夹路径
    :param exe_path: 7z.exe 路径
    :param password: zip文件密码
    '''
    args = [exe_path, "e", file_path, f'-o{output_dir}',f'-p{password}', "-y"]
    try:
        subprocess.run(args, check=True, text=True, capture_output=True)
        remove_empty_dirs(output_dir)
        return True
    except subprocess.CalledProcessError as e:
        logger.error(e)
        return False
    except Exception as e:
        logger.error(e)
        return False

def remove_empty_dirs(directory):
    '''
    清除空文件夹
    :param directory: zip解压后的文件夹路径
    '''
    # 为了确保先从最深层的子目录开始，我们使用os.walk并设置topdown为False
    for root, dirs, files in os.walk(directory, topdown=False):
        for name in dirs:
            try:
                os.rmdir(os.path.join(root, name))
            except OSError:
                pass

def check_pdf(file_path):
    """检查 PDF 文件是否加密或损坏，1=页码不正常，2=正常，-1=加密，-2=损坏"""
    try:
        reader = pypdf.PdfReader(file_path)
        if reader.is_encrypted:
            return -1
        numPages = len(reader.pages)
        if numPages <= 10:
            return 1
        return 2
    except Exception as e:
        logger.error(e)
        return -2

# pdg转pdf
def pdg2pdf(exe_path,input_dir):
    '''
    pdg转pdf,输入文件夹的上一级是pdf 输出位置，如果已经存在pdf文件，则不进行转换，如果压缩包里边是pdf 则移动出来
    :param exe_path: pdg2pic.exe 路径
    :param input_dir: pdg文件夹路径
    '''
    app = None
    try:
        dir_path = str(input_dir).rsplit("/", 1)
        dirname, filename = dir_path[0], dir_path[-1] + ".pdf"
        pdf_path = os.path.join(dirname, filename)
        if os.path.exists(pdf_path):
            logger.debug(f"pdf文件已经存在:{pdf_path}！")
            return True

        # 检查文件夹中是否存在pdg文件
        pdg_files = [file for file in os.listdir(input_dir) if file.endswith(".pdg")]
        pdf_files = [file for file in os.listdir(input_dir) if file.endswith(".pdf")]
        if not pdg_files and not pdf_files:
            logger.debug(f"未找到pdg文件，或者是pdf文件！")
            return False

        if pdf_files:
            # 将pdf 文件移到指定位置，然后删除这个文件夹
            pdf0_path = os.path.join(input_dir, pdf_files[0])
            os.rename(pdf0_path, pdf_path)
            return True

        # 启动pdg2pic
        app = Application(backend='win32').start(exe_path)
        main_win = app.window(title_re='Pdg2Pic*', class_name_re='#32770',top_level_only=True,found_index=0)

        input_dir_status = False
        for x in range(3):
            logger.debug(f"---------------------------")
            logger.debug(f"开始第{x+1}次选择文件夹！")
            try:
                # 选择文件夹
                xuanze_but = main_win.child_window(class_name="Button", found_index=1).wait(f"exists", timeout=2)
                xuanze_but.click_input()
                # 注意：选择文件夹需要利用系统的bug，快速进行填入并且点击确定
                file_win = app.window(title="选择存放PDG文件的文件夹")
                file_win["文件夹(&F):Edit"].set_edit_text(input_dir)
                file_win["确定"].click_input()
            except Exception as e:
                logger.exception(f"选择文件夹失败：{e}")
                continue
            finally:
                logger.debug(f"开始检查文件选择窗是否正常关闭！")
                try:
                    file_win2 = app.window(title="选择存放PDG文件的文件夹").wait("exists", timeout=1)
                    if file_win2:
                        file_win2.close()
                        logger.error(f"选择文件窗未关闭，强制关闭！")
                except Exception as e:
                    pass

                logger.debug(f"开始检查文件夹是否选择完成")
                try:
                    popup_text = "\n".join([control.window_text() for control in main_win.children()])
                    if input_dir in popup_text:
                        input_dir_status = True
                        logger.success(f"文件选择成功！")
                        break
                    logger.error(f"文件选择失败，重新进行选择文件！")
                except Exception as e:
                    pass
                logger.debug("==================================")

        assert input_dir_status, "文件选择失败！"
        logger.success(f"文件选择成功！")
        # 确定，开始转换
        def click():
            try:
                start_but = main_win.child_window(title="&4、开始转换",class_name="Button").wait("exists", timeout=1)
                assert start_but, "未检测到开始转换按钮！"
                start_but.click_input()
            except Exception as e:
                logger.error(f"未知错误：{e}")

        logger.debug("开始转换！")
        t.submit(click)
        time.sleep(3)

        # 循环确定当前是否成功点击了转化
        stop_but = False
        for x in range(3):
            try:
                main_win.child_window(title="停止", class_name="Button").wait("exists", timeout=2)
                stop_but = True
                logger.debug(f"已经点击转化！")
                time.sleep(3)
                break
            except Exception as e:
                logger.error(e)
                t.submit(click)


        if not stop_but:
            logger.error(f"未检测到停止按钮！")
            return False

        start_time,max_time = int(time.time()),60*5
        status = False

        while True:
            now_time = int(time.time())
            if now_time - start_time > max_time:
                logger.debug(f"转化超时！")
                break

            logger.debug(f"正在转化中,耗时：{now_time-start_time}秒！")
            try:
                popup = app.window(title_re='Pdg2Pic*', class_name_re='#32770',found_index=0)
                logger.debug("开始检查弹窗是否存在1！")
                popup.wait('ready', timeout=10)
                logger.debug("开始检查弹窗是否存在2！")
                popup_text = "\n".join([control.window_text() for control in popup.children()])
            except Exception as e:
                logger.error(e)
                popup_text = ""

            # logger.debug(f"popup_text:{popup_text}")
            if "确定" in popup_text or "是" in popup_text or "转化完毕" in popup_text:
                if "转换完毕" in popup_text:
                    logger.debug(f"转换完毕！")
                    status = True
                    break
                logger.error(f"未知错误：{popup_text}")
                break
            time.sleep(3)
        app.kill()
        return status
    except Exception as e:
        logger.error(e)
        if app:
            app.kill()
        assert Exception(f"未知错误：{e}")

def extract_and_reduce_dpi(pdf_path, output_pdf_path, factor=0.7):
    '''
    通过降低图片的dpi进行压缩pdf，慎重使用
    '''
    doc = fitz.open(pdf_path)
    pdf_writer = fitz.open()

    output = io.BytesIO()  # 重用的字节流对象
    pil_images = []  # 存储所有的PIL Image对象

    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        base_image = page.get_pixmap(dpi=200)
        pil_image = Image.open(io.BytesIO(base_image.tobytes()))
        pil_images.append(pil_image)

    logger.debug(f"开始进行插入图片！")
    for e,pil_image in enumerate(pil_images):
        new_width = int(pil_image.width * factor)
        new_height = int(pil_image.height * factor)
        pil_image_resampled = pil_image.resize((new_width, new_height), Image.LANCZOS)

        pil_image_resampled.save(output, format="jpeg", quality=80)
        pdf_writer_page = pdf_writer.new_page(width=new_width, height=new_height)
        pdf_writer_page.insert_image(fitz.Rect(0, 0, new_width, new_height), stream=output.getvalue())
        output.seek(0)
        output.truncate()

    pdf_writer.save(output_pdf_path)
    pdf_writer.close()
    doc.close()
    logger.success(f"PDF 处理成功！")


if __name__ == '__main__':
    base_path = os.path.dirname(os.path.abspath(__file__))
    password_path = os.path.join(base_path, "../passwords", "passwords.txt")
    zip_path = os.path.join(base_path, "../7-Zip", "7z.exe")
    pdg_exe_path = os.path.join(base_path, "../Pdg2Pic", "Pdg2Pic.exe")

    zip_file_path = r"D:\Desktop\独秀\做人生路上的“不倒翁”_13594976.zip"
    output_dir = r"D:\Desktop\独秀\13594976"
    pdg2pdf(pdg_exe_path,output_dir)


