# -*- coding: utf-8 -*-

import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QListWidget, QLabel, QLineEdit, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
import os
import shutil
import uuid
import re
import posixpath
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET

# Register namespaces for ElementTree
ET.register_namespace('', "urn:oasis:names:tc:opendocument:xmlns:container")
ET.register_namespace('opf', "http://www.idpf.org/2007/opf")
ET.register_namespace('ncx', "http://www.daisy.org/z3986/2005/ncx/")

class EpubListWidget(QListWidget):
    """
    一个支持文件拖放的 QListWidget
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragDropMode(QListWidget.InternalMove)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            for url in urls:
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                    if file_path.endswith('.epub'):
                        self.addItem(file_path)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)

class EpubMergerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("EPUB 文件合并器")
        self.setGeometry(100, 100, 600, 400)
        self.initUI()
        self.temp_dir = "" # This is not strictly necessary now, but good practice.

    def initUI(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)

        # 标题标签
        title_label = QLabel("将 EPUB 文件拖动至下方")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 20px; font-weight: bold;")
        main_layout.addWidget(title_label)

        # 文件列表和操作按钮
        h_layout = QHBoxLayout()
        
        # 文件列表
        self.file_list = EpubListWidget()
        h_layout.addWidget(self.file_list, 3)

        # 按钮布局
        button_layout = QVBoxLayout()
        
        add_btn = QPushButton("添加文件...")
        add_btn.clicked.connect(self.add_files)
        button_layout.addWidget(add_btn)

        remove_btn = QPushButton("移除文件")
        remove_btn.clicked.connect(self.remove_file)
        button_layout.addWidget(remove_btn)
        
        move_up_btn = QPushButton("上移")
        move_up_btn.clicked.connect(self.move_up)
        button_layout.addWidget(move_up_btn)

        move_down_btn = QPushButton("下移")
        move_down_btn.clicked.connect(self.move_down)
        button_layout.addWidget(move_down_btn)

        button_layout.addStretch()

        h_layout.addLayout(button_layout, 1)
        main_layout.addLayout(h_layout)
        
        # 输出文件和合并按钮
        output_layout = QHBoxLayout()
        
        output_label = QLabel("输出文件:")
        output_layout.addWidget(output_label)

        self.output_name_edit = QLineEdit("合并后的文件.epub")
        output_layout.addWidget(self.output_name_edit)
        
        merge_btn = QPushButton("合并 EPUB")
        merge_btn.clicked.connect(self.merge_epubs)
        output_layout.addWidget(merge_btn)

        main_layout.addLayout(output_layout)
        
    def add_files(self):
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        file_dialog.setNameFilter("EPUB Files (*.epub)")
        if file_dialog.exec_():
            selected_files = file_dialog.selectedFiles()
            for file_path in selected_files:
                self.file_list.addItem(file_path)

    def remove_file(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))

    def move_up(self):
        row = self.file_list.currentRow()
        if row > 0:
            item = self.file_list.takeItem(row)
            self.file_list.insertItem(row - 1, item)
            self.file_list.setCurrentRow(row - 1)

    def move_down(self):
        row = self.file_list.currentRow()
        if row < self.file_list.count() - 1:
            item = self.file_list.takeItem(row)
            self.file_list.insertItem(row + 1, item)
            self.file_list.setCurrentRow(row + 1)
            
    def merge_epubs(self):
        if self.file_list.count() == 0:
            QMessageBox.warning(self, "警告", "请先添加要合并的 EPUB 文件。")
            return

        output_file_name = self.output_name_edit.text()
        if not output_file_name.strip():
            QMessageBox.warning(self, "警告", "输出文件名不能为空。")
            return
        if not output_file_name.endswith('.epub'):
            output_file_name += '.epub'

        # 创建临时处理目录
        temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp_epub_merge")
        
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)
            
            # 定义输出目录结构
            out_root_dir = os.path.join(temp_dir, 'merged_epub')
            out_image_dir = os.path.join(out_root_dir, 'image')
            out_html_dir = os.path.join(out_root_dir, 'html')
            out_css_dir = os.path.join(out_root_dir, 'css')
            
            os.makedirs(os.path.join(out_root_dir, 'META-INF'))
            os.makedirs(out_image_dir)
            os.makedirs(out_html_dir)
            os.makedirs(out_css_dir)
            
            # 处理 mimetype 和 container.xml
            with open(os.path.join(out_root_dir, 'mimetype'), 'w') as f:
                f.write("application/epub+zip")
            
            container_content = """<?xml version="1.0"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="vol.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>"""
            with open(os.path.join(out_root_dir, 'META-INF', 'container.xml'), 'w', encoding='utf-8') as f:
                f.write(container_content)

            total_pages = 0
            opf_metadata = None
            
            image_regex = re.compile(r'<img[^>]*src=["\'](?P<path>.*?\.jpg)["\'][^>]*>', re.IGNORECASE)

            # 解压并处理每个 EPUB
            for i in range(self.file_list.count()):
                epub_path = self.file_list.item(i).text()
                
                # 创建临时解压目录
                temp_epub_dir = os.path.join(temp_dir, f"epub_{i}")
                with ZipFile(epub_path, 'r') as z:
                    z.extractall(temp_epub_dir)
                
                # 找到 OPF 文件路径
                opf_path = ""
                try:
                    with open(os.path.join(temp_epub_dir, 'META-INF', 'container.xml'), 'r', encoding='utf-8') as f:
                        container_tree = ET.parse(f)
                        rootfile = container_tree.find('.//{urn:oasis:names:tc:opendocument:xmlns:container}rootfile')
                        if rootfile is not None:
                            opf_path = os.path.join(temp_epub_dir, rootfile.attrib['full-path'])
                        else:
                            QMessageBox.critical(self, "错误", f"找不到 {epub_path} 中的 OPF 文件路径。")
                            return
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"解析 {epub_path} 中的 container.xml 时出错：\n{e}")
                    return

                # 解析 OPF 文件
                opf_tree = ET.parse(opf_path)
                opf_root = opf_tree.getroot()
                opf_root_dir = os.path.dirname(opf_path)
                
                # 提取第一个 EPUB 的 metadata
                if i == 0:
                    opf_metadata = opf_root.find('{http://www.idpf.org/2007/opf}metadata')
                    
                # 遍历所有HTML文件，处理图片和创建新HTML
                for item in opf_root.find('{http://www.idpf.org/2007/opf}manifest').findall('{http://www.idpf.org/2007/opf}item'):
                    if item.attrib.get('media-type') == 'application/xhtml+xml':
                        html_path_relative = item.attrib['href']
                        html_path_absolute = os.path.join(opf_root_dir, html_path_relative)
                        if os.path.exists(html_path_absolute):
                            try:
                                with open(html_path_absolute, 'r', encoding='utf-8') as f:
                                    html_content = f.read()
                                
                                match = image_regex.search(html_content)
                                if match:
                                    img_src = match.group('path')
                                    img_src_absolute = os.path.normpath(os.path.join(os.path.dirname(html_path_absolute), img_src))
                                    
                                    total_pages += 1
                                    # 格式化文件名
                                    new_img_name = f"merge-{total_pages:06d}.jpg"
                                    new_html_name = f"mergeP-{total_pages:06d}.html"

                                    # 复制并重命名图片
                                    shutil.copy(img_src_absolute, os.path.join(out_image_dir, new_img_name))
                                    
                                    # 生成新的HTML内容
                                    new_html_content = f"""<!DOCTYPE html SYSTEM "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
 <title>第 {total_pages} 頁</title>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
 <meta name="viewport" content="width=1440, height=1920" />
 <link rel="stylesheet" type="text/css" href="../css/style.css" />
</head>
<body>
<div class="fs">
 <div>
  <img src="../image/{new_img_name}" alt="第 {total_pages} 頁" class="singlePage" kmoetag="rotate:0" />
 </div>
</div>
</body>
</html>"""
                                    # 将新的HTML写入文件
                                    with open(os.path.join(out_html_dir, new_html_name), 'w', encoding='utf-8') as f:
                                        f.write(new_html_content)
                                    
                            except Exception as e:
                                QMessageBox.critical(self, "错误", f"处理 {html_path_absolute} 时出错：\n{e}")
                                return
                
                # 复制CSS文件
                css_items = opf_root.find('{http://www.idpf.org/2007/opf}manifest').findall('{http://www.idpf.org/2007/opf}item')
                for item in css_items:
                    if item.attrib.get('media-type') == 'text/css':
                        css_path_relative = item.attrib['href']
                        css_path_absolute = os.path.join(opf_root_dir, css_path_relative)
                        if os.path.exists(css_path_absolute):
                            shutil.copy(css_path_absolute, os.path.join(out_css_dir, os.path.basename(css_path_absolute)))
                            break # 只复制第一个找到的CSS文件

            # 重建 OPF 文件
            opf_root = ET.Element('{%s}package' % "http://www.idpf.org/2007/opf",
                                      attrib={'version': '2.0', 'unique-identifier': 'uuid_id'})
            
            # 添加 metadata
            if opf_metadata is not None:
                opf_root.append(opf_metadata)
                
            # 添加 manifest
            manifest = ET.SubElement(opf_root, '{http://www.idpf.org/2007/opf}manifest')
            ET.SubElement(manifest, '{http://www.idpf.org/2007/opf}item', attrib={
                'id': 'ncx',
                'href': 'vol.ncx',
                'media-type': 'application/x-dtbncx+xml'
            })
            ET.SubElement(manifest, '{http://www.idpf.org/2007/opf}item', attrib={
                'id': 'css',
                'href': 'css/style.css',
                'media-type': 'text/css'
            })
            
            for i in range(1, total_pages + 1):
                ET.SubElement(manifest, '{http://www.idpf.org/2007/opf}item', attrib={
                    'id': f'page_{i:06d}',
                    'href': f'html/mergeP-{i:06d}.html',
                    'media-type': 'application/xhtml+xml'
                })
                ET.SubElement(manifest, '{http://www.idpf.org/2007/opf}item', attrib={
                    'id': f'img_{i:06d}',
                    'href': f'image/merge-{i:06d}.jpg',
                    'media-type': 'image/jpeg'
                })
            
            # 添加 spine
            spine = ET.SubElement(opf_root, '{http://www.idpf.org/2007/opf}spine', attrib={'toc': 'ncx'})
            for i in range(1, total_pages + 1):
                ET.SubElement(spine, '{http://www.idpf.org/2007/opf}itemref', attrib={'idref': f'page_{i:06d}'})

            opf_tree = ET.ElementTree(opf_root)
            opf_tree.write(os.path.join(out_root_dir, 'vol.opf'), encoding='utf-8', xml_declaration=True)

            # 重建 NCX 文件
            ncx_root = ET.Element('{%s}ncx' % "http://www.daisy.org/z3986/2005/ncx/",
                                      attrib={'version': '2005-1', 'xmlns': "http://www.daisy.org/z3986/2005/ncx/"})
            head = ET.SubElement(ncx_root, '{http://www.daisy.org/z3986/2005/ncx/}head')
            ET.SubElement(head, '{http://www.daisy.org/z3986/2005/ncx/}meta', attrib={
                'name': 'dtb:uid',
                'content': str(uuid.uuid4())
            })
            ET.SubElement(head, '{http://www.daisy.org/z3986/2005/ncx/}meta', attrib={
                'name': 'dtb:totalPageCount',
                'content': str(total_pages)
            })
            ET.SubElement(head, '{http://www.daisy.org/z3986/2005/ncx/}meta', attrib={
                'name': 'dtb:maxPageNumber',
                'content': str(total_pages)
            })
            
            docTitle = ET.SubElement(ncx_root, '{http://www.daisy.org/z3986/2005/ncx/}docTitle')
            text = ET.SubElement(docTitle, '{http://www.daisy.org/z3986/2005/ncx/}text')
            text.text = '合并后的EPUB'
            
            navMap = ET.SubElement(ncx_root, '{http://www.daisy.org/z3986/2005/ncx/}navMap')
            for i in range(1, total_pages + 1):
                navPoint = ET.SubElement(navMap, '{http://www.daisy.org/z3986/2005/ncx/}navPoint', attrib={
                    'id': f'Page_{i}',
                    'playOrder': str(i)
                })
                navLabel = ET.SubElement(navPoint, '{http://www.daisy.org/z3986/2005/ncx/}navLabel')
                text = ET.SubElement(navLabel, '{http://www.daisy.org/z3986/2005/ncx/}text')
                text.text = f'第 {i:03d} 頁'
                content = ET.SubElement(navPoint, '{http://www.daisy.org/z3986/2005/ncx/}content', attrib={
                    'src': f'html/mergeP-{i:06d}.html'
                })
                
            ncx_tree = ET.ElementTree(ncx_root)
            ncx_tree.write(os.path.join(out_root_dir, 'vol.ncx'), encoding='utf-8', xml_declaration=True)

            # 重新打包成 EPUB
            with ZipFile(output_file_name, 'w', ZIP_DEFLATED) as z:
                for root, dirs, files in os.walk(out_root_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        rel_path = os.path.relpath(file_path, out_root_dir)
                        if rel_path == 'mimetype':
                            z.write(file_path, rel_path, compress_type=ZIP_DEFLATED)
                        else:
                            z.write(file_path, rel_path)
            
            QMessageBox.information(self, "成功", f"文件已成功合并并保存为: {output_file_name}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"合并文件时出错：\n{e}")
        finally:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = EpubMergerApp()
    ex.show()
    sys.exit(app.exec_())
