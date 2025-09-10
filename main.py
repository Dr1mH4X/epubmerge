# -*- coding: utf-8 -*-

import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QListWidget, QLabel, QLineEdit, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
from ebooklib import epub
import os
import shutil
import uuid
import re
import posixpath

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
        self.temp_dir = ""

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

        new_book = epub.EpubBook()
        new_book.set_identifier(str(uuid.uuid4()))
        new_book.set_title('合并后的EPUB')
        new_book.set_language('zh-CN')

        items_to_add = []
        spine_ids = []
        toc_links = []
        
        # Step 1: Collect all items and build a path mapping
        path_map = {}
        source_books_items = []
        for i in range(self.file_list.count()):
            file_path = self.file_list.item(i).text()
            try:
                source_book = epub.read_epub(file_path)
                source_books_items.append((os.path.basename(file_path).replace('.epub', ''), source_book.get_items()))
                
                # Build path map without UUID
                for item in source_book.get_items():
                    new_file_name = posixpath.join(os.path.basename(file_path).replace('.epub', ''), item.file_name)
                    path_map[item.file_name] = new_file_name
            except Exception as e:
                QMessageBox.critical(self, "错误", f"处理文件 {file_path} 时出错：\n{e}")
                return
        
        # Step 2: Process items and update links and image paths
        for original_book_name, items in source_books_items:
            first_chapter_added = False
            for item in items:
                new_item_content = item.content
                if item.media_type == 'application/xhtml+xml':
                    content_str = new_item_content.decode('utf-8')
                    # Use a general regex to find and replace all href and src attributes
                    content_str = re.sub(
                        r'(href|src)=["\'](?P<path>.*?)["\']',
                        lambda m: f'{m.group(1)}="{path_map.get(m.group("path"), m.group("path"))}"',
                        content_str
                    )
                    new_item_content = content_str.encode('utf-8')

                new_file_name = posixpath.join(original_book_name, item.file_name)
                
                new_item = epub.EpubItem(
                    uid=str(uuid.uuid4()),
                    file_name=new_file_name,
                    media_type=item.media_type,
                    content=new_item_content
                )
                items_to_add.append(new_item)
                
                # If it's a content document, add it to the spine
                if new_item.media_type == 'application/xhtml+xml':
                    spine_ids.append(new_item.get_id())
                    
                    # Add a TOC link to the first chapter of each original book
                    if not first_chapter_added:
                        toc_links.append(epub.Section(original_book_name, new_item.file_name))
                        first_chapter_added = True
        
        # Step 3: Add all collected items to the new book
        for item in items_to_add:
            new_book.add_item(item)
        
        # Create a new navigation item for the book
        new_nav = epub.EpubNav(uid='nav')
        new_book.add_item(new_nav)
        
        # Set the new spine and toc
        new_book.spine = ['nav'] + spine_ids
        new_book.toc = toc_links
        
        # Write the new epub
        try:
            epub.write_epub(output_file_name, new_book, {})
            QMessageBox.information(self, "成功", f"文件已成功合并并保存为: {output_file_name}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存文件时出错：\n{e}")

    def closeEvent(self, event):
        if os.path.exists(self.temp_dir):
            try:
                shutil.rmtree(self.temp_dir)
            except OSError as e:
                print(f"删除临时文件夹失败: {e}")
        super().closeEvent(event)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = EpubMergerApp()
    ex.show()
    sys.exit(app.exec_())
