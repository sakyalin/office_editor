import os
import sys
import threading
from pathlib import Path

from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLineEdit, QFileDialog, QMessageBox, QProgressBar,
                             QWidget, QFrame)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QFont

# 导入环境变量加载模块
from env_loader import load_env_variables

# 加载环境变量
load_env_variables()

from office_processor import process_document


class ProcessThread(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(int, str)  # 进度百分比和当前处理的内容
    
    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path
    
    def progress_callback(self, percent, message):
        self.progress.emit(percent, message)
    
    def run(self):
        try:
            output_path = process_document(self.file_path, self.progress_callback)
            self.finished.emit(str(output_path))
        except Exception as e:
            self.error.emit(str(e))


class OfficeEditorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Office文档错别字检查工具")
        self.setGeometry(100, 100, 600, 400)
        self.setMinimumSize(500, 350)
        
        self.setup_ui()
    
    def setup_ui(self):
        # 创建中央部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # 标题
        title_label = QLabel("Office文档错别字检查工具")
        title_font = QFont("Arial", 16)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        main_layout.addSpacing(10)
        
        # 说明文本
        description = "上传PPT或Word文件，使用OpenAI检查错别字和病句。\n处理完成后将在相同位置生成带'修订'后缀的文件。"
        desc_label = QLabel(description)
        desc_label.setWordWrap(True)
        desc_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(desc_label)
        main_layout.addSpacing(20)
        
        # 提示信息
        api_info_label = QLabel("OpenAI API密钥已从.env文件读取")
        api_info_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(api_info_label)
        main_layout.addSpacing(10)
        
        # 文件选择框架
        file_frame = QFrame()
        file_layout = QHBoxLayout(file_frame)
        file_layout.setContentsMargins(0, 0, 0, 0)
        
        file_label = QLabel("选择文件:")
        file_layout.addWidget(file_label)
        
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("选择Office文件...")
        file_layout.addWidget(self.file_path_edit)
        
        browse_button = QPushButton("浏览")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(browse_button)
        
        main_layout.addWidget(file_frame)
        main_layout.addSpacing(20)
        
        # 处理按钮
        self.process_button = QPushButton("开始处理")
        self.process_button.clicked.connect(self.process_file)
        self.process_button.setMinimumHeight(40)
        main_layout.addWidget(self.process_button)
        main_layout.addSpacing(20)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)  # 显示进度文本
        self.progress_bar.setRange(0, 100)  # 设置为确定模式，范围0-100
        self.progress_bar.hide()
        main_layout.addWidget(self.progress_bar)
        
        # 状态标签
        self.status_label = QLabel("准备就绪")
        self.status_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.status_label)
        
        # 底部信息
        main_layout.addStretch(1)
        footer_label = QLabel("支持 .docx, .pptx 格式文件")
        footer_font = QFont("Arial", 8)
        footer_label.setFont(footer_font)
        footer_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(footer_label)
    
    def browse_file(self):
        file_filter = "Office文件 (*.docx *.pptx);;Word文档 (*.docx);;PowerPoint演示文稿 (*.pptx);;所有文件 (*.*)"
        file_path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", file_filter)
        if file_path:
            self.file_path_edit.setText(file_path)
    
    def process_file(self):
        # 获取文件路径
        file_path = self.file_path_edit.text()
        if not file_path:
            QMessageBox.critical(self, "错误", "请先选择一个文件")
            return
        
        file_path = Path(file_path)
        if not file_path.exists():
            QMessageBox.critical(self, "错误", "文件不存在")
            return
        
        if file_path.suffix.lower() not in [".docx", ".pptx"]:
            QMessageBox.critical(self, "错误", "不支持的文件格式，请选择.docx或.pptx文件")
            return
        
        # 禁用按钮并显示进度条
        self.process_button.setEnabled(False)
        self.progress_bar.show()
        self.status_label.setText("正在处理文件...")
        
        # 在新线程中处理文件
        self.process_thread = ProcessThread(file_path)
        self.process_thread.finished.connect(self.processing_complete)
        self.process_thread.error.connect(self.processing_error)
        self.process_thread.progress.connect(self.update_progress)
        self.process_thread.start()
    
    @pyqtSlot(str)
    def processing_complete(self, output_path):
        self.progress_bar.hide()
        self.process_button.setEnabled(True)
        self.status_label.setText("处理完成")
        
        QMessageBox.information(self, "完成", f"文件处理完成！\n\n已保存至: {output_path}")
    
    @pyqtSlot(int, str)
    def update_progress(self, percent, message):
        self.progress_bar.setValue(percent)
        self.status_label.setText(message)
    
    @pyqtSlot(str)
    def processing_error(self, error_message):
        self.progress_bar.hide()
        self.process_button.setEnabled(True)
        self.status_label.setText("处理出错")
        
        QMessageBox.critical(self, "错误", f"处理文件时出错:\n{error_message}")


def main():
    app = QApplication(sys.argv)
    window = OfficeEditorApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()