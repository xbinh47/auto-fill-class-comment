import sys
import os
import json
import unicodedata
import pandas as pd

from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QTextEdit, QCheckBox,
    QComboBox, QFileDialog, QMessageBox, QVBoxLayout, QHBoxLayout, QScrollArea
)
from PyQt6.QtCore import Qt, QTimer


class LessonDataUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tạo nội dung buổi học")
        self.resize(800, 900)

        self.cache_file = 'ui_cache.json'
        self.excel_data = None

        # Khởi tạo trạng thái mặc định của các checkbox
        self.checkbox_states = {
            'class_performance': True,  # Tình hình học tập của lớp
            'lesson_content': True,     # Nội dung buổi học (Excel)
            'slide_link': True,
            'video_link': True,
            'homework_result': True,
            'next_requirement': True,
            'deadline': True,
            'next_lesson_content': True
        }

        self.init_ui()

        # Sau 100ms, load cache
        QTimer.singleShot(100, self.load_cache)

    def init_ui(self):
        # Tạo layout chính của widget cha
        main_layout = QVBoxLayout(self)

        # Tạo QScrollArea để có thể cuộn khi giao diện quá dài
        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        main_layout.addWidget(scroll_area)

        # Tạo một widget chứa toàn bộ nội dung của giao diện
        container = QWidget()
        scroll_area.setWidget(container)

        # Layout cho widget chứa
        container_layout = QVBoxLayout(container)

        # --- Phần chọn file ---
        file_layout = QHBoxLayout()
        file_label = QLabel("Đường dẫn file:")
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        self.browse_button = QPushButton("Chọn file")
        self.browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_path_edit)
        file_layout.addWidget(self.browse_button)
        container_layout.addLayout(file_layout)

        # --- Chọn sheet ---
        sheet_layout = QHBoxLayout()
        sheet_label = QLabel("Tên sheet:")
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentIndexChanged.connect(self.on_sheet_or_lesson_change)
        sheet_layout.addWidget(sheet_label)
        sheet_layout.addWidget(self.sheet_combo)
        container_layout.addLayout(sheet_layout)

        # --- Chọn số bài học ---
        lesson_layout = QHBoxLayout()
        lesson_label = QLabel("Số bài học:")
        self.lesson_combo = QComboBox()
        for i in range(1, 15):
            self.lesson_combo.addItem(str(i))
        self.lesson_combo.currentIndexChanged.connect(self.on_sheet_or_lesson_change)
        lesson_layout.addWidget(lesson_label)
        lesson_layout.addWidget(self.lesson_combo)
        container_layout.addLayout(lesson_layout)

        # --- Các trường nhập liệu (mỗi trường có checkbox bật/tắt) ---
        self.lesson_content_widget = self.create_text_input_with_checkbox("Nội dung buổi học:", "lesson_content")
        container_layout.addWidget(self.lesson_content_widget)

        self.class_performance_widget = self.create_text_input_with_checkbox("Tình hình học tập của lớp:", "class_performance")
        container_layout.addWidget(self.class_performance_widget)

        self.slide_link_widget = self.create_text_input_with_checkbox("Link slide:", "slide_link")
        container_layout.addWidget(self.slide_link_widget)

        self.video_link_widget = self.create_text_input_with_checkbox("Link video:", "video_link")
        container_layout.addWidget(self.video_link_widget)

        self.homework_result_widget = self.create_text_input_with_checkbox("Kết quả bài tập về nhà:", "homework_result")
        container_layout.addWidget(self.homework_result_widget)

        self.next_requirement_widget = self.create_text_input_with_checkbox("Yêu cầu cho buổi tiếp theo:", "next_requirement")
        container_layout.addWidget(self.next_requirement_widget)

        self.deadline_widget = self.create_text_input_with_checkbox("Hạn nộp bài:", "deadline")
        container_layout.addWidget(self.deadline_widget)

        self.next_lesson_content_widget = self.create_text_input_with_checkbox("Nội dung buổi tới:", "next_lesson_content")
        container_layout.addWidget(self.next_lesson_content_widget)

        # --- Nút tạo nội dung ---
        self.generate_button = QPushButton("Tạo nội dung")
        self.generate_button.clicked.connect(self.generate_formatted_text)
        container_layout.addWidget(self.generate_button)

        # --- Khu vực hiển thị kết quả ---
        result_label = QLabel("Kết quả:")
        container_layout.addWidget(result_label)
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setMinimumHeight(200)
        container_layout.addWidget(self.result_text)

    def create_text_input_with_checkbox(self, label_text, attr_name):
        """Tạo widget gồm checkbox, nhãn và QTextEdit."""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        top_layout = QHBoxLayout()
        checkbox = QCheckBox()
        checkbox.setChecked(self.checkbox_states.get(attr_name, True))
        top_layout.addWidget(checkbox)
        label = QLabel(label_text)
        top_layout.addWidget(label)
        top_layout.addStretch()
        layout.addLayout(top_layout)
        text_edit = QTextEdit()
        text_edit.setFixedHeight(80)
        layout.addWidget(text_edit)

        setattr(self, f"{attr_name}_checkbox", checkbox)
        setattr(self, f"{attr_name}_edit", text_edit)

        checkbox.toggled.connect(lambda: self.save_cache(self.file_path_edit.text()))
        return widget

    # --------------------- Xử lý Excel ---------------------
    def read_excel_data(self, file_path, sheet_name, lesson_number):
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            lesson_column = f"Buổi {lesson_number}"
            if lesson_column not in df.columns:
                raise ValueError(f"Lesson {lesson_number} not found in sheet {sheet_name}.")
            return self._process_column_data(df[lesson_column])
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            raise

    def _process_column_data(self, column):
        data = {}

        def remove_vietnamese_accent_local(text):
            return self.remove_vietnamese_accent(text)

        column_data = []
        for x in column.tolist():
            if pd.isna(x):
                column_data.append("")
            else:
                column_data.append(str(x).strip())

        output_labels = {
            "Nội dung buổi học số": "📌Nội dung buổi học số",
            "Link slide bài giảng": "📌Link slide bài giảng",
            "Link Video hướng dẫn": "📌Link Video hướng dẫn",
            "Kết quả bài tập về nhà": "🏆Kết quả bài tập về nhà",
            "Hạn nộp bài": "⏰Hạn nộp bài",
            "Yêu cầu cho buổi tiếp theo": "🔍Yêu cầu cho buổi tiếp theo",
            "Nội dung buổi học tới": "📌 Nội dung buổi học tới"
        }

        i = 0
        while i < len(column_data):
            value = column_data[i].strip()
            clean_value = remove_vietnamese_accent_local(value)
            if "noi dung buoi hoc so" in clean_value:
                title = value
                content = []
                i += 1
                while i < len(column_data):
                    next_value = column_data[i].strip()
                    if any(remove_vietnamese_accent_local(lbl) in remove_vietnamese_accent_local(next_value)
                           for lbl in output_labels.keys()):
                        break
                    content.append(next_value)
                    i += 1
                data["📌Nội dung buổi học số"] = f"{title}\n" + "\n".join(content)
                continue

            for label, output_label in output_labels.items():
                if remove_vietnamese_accent_local(label) in remove_vietnamese_accent_local(value):
                    content = []
                    i += 1
                    while i < len(column_data):
                        next_value = column_data[i].strip()
                        if any(remove_vietnamese_accent_local(lbl) in remove_vietnamese_accent_local(next_value)
                               for lbl in output_labels.keys()):
                            break
                        content.append(next_value)
                        i += 1
                    data[output_label] = "\n".join(content)
                    break
            i += 1

        for output_label in output_labels.values():
            if output_label not in data:
                data[output_label] = ""
        print("Processed data:", data)
        return data

    def remove_vietnamese_accent(self, text):
        s1 = "ÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚÝàáâãèéêìíòóôõùúýĂăĐđĨĩŨũƠơƯưẠạẢảẤấẦầẨẩẪẫẬậẮắẰằẲẳẴẵẶặẸẹẺẻẼẽẾếỀềỂểỄễỆệỈỉỊịỌọỎỏỐốỒồỔổỖỗỘộỚớỜờỞởỠỡỢợỤụỦủỨứỪừỬửỮữỰựỲỳỴỵỶỷỸỹ"
        s0 = "AAAAEEEIIOOOOUUYaaaaeeeiioooouuyAaDdIiUuOoUuAaAaAaAaAaAaAaAaAaAaAaAaEeEeEeEeEeEeEeEeIiIiOoOoOoOoOoOoOoOoOoOoOoOoUuUuUuUuUuUuUuYyYyYyYy"
        for i in range(len(s1)):
            text = text.replace(s1[i], s0[i])
        text = unicodedata.normalize('NFKD', text)
        text = ''.join(c for c in text if not unicodedata.combining(c))
        return text.lower()

    def on_sheet_or_lesson_change(self):
        file_path = self.file_path_edit.text()
        if file_path and self.sheet_combo.currentText() and self.lesson_combo.currentText():
            try:
                lesson_data = self.read_excel_data(
                    file_path,
                    self.sheet_combo.currentText(),
                    int(self.lesson_combo.currentText())
                )
                self.fill_inputs_from_excel(lesson_data)
            except Exception as e:
                print(f"Error loading lesson data: {e}")

    def fill_inputs_from_excel(self, lesson_data):
        def remove_vietnamese_accent_local(text):
            return self.remove_vietnamese_accent(text)

        input_fields = {
            'class_performance': 'Tình hình học tập của lớp',
            'lesson_content': 'Nội dung buổi học số',
            'slide_link': 'Link slide bài giảng',
            'video_link': 'Link Video hướng dẫn',
            'homework_result': 'Kết quả bài tập về nhà',
            'next_requirement': 'Yêu cầu cho buổi tiếp theo',
            'deadline': 'Hạn nộp bài',
            'next_lesson_content': 'Nội dung buổi học tới'
        }

        for attr, label in input_fields.items():
            text_edit = getattr(self, f"{attr}_edit")
            text_edit.clear()
            for data_label, content in lesson_data.items():
                if remove_vietnamese_accent_local(label) in remove_vietnamese_accent_local(data_label):
                    text_edit.setPlainText(content)

    def generate_formatted_text(self):
        try:
            contents = {}
            fields = ['class_performance', 'lesson_content', 'slide_link', 'video_link',
                      'homework_result', 'next_requirement', 'deadline', 'next_lesson_content']
            for field in fields:
                checkbox = getattr(self, f"{field}_checkbox")
                text_edit = getattr(self, f"{field}_edit")
                if checkbox.isChecked():
                    contents[field] = text_edit.toPlainText().strip()
                else:
                    contents[field] = ""

            # Tạo nội dung HTML với định dạng (sử dụng thẻ <b> cho chữ đậm, <br> cho xuống dòng)
            html = "<p><b>- Chào cả lớp, Thầy/Em xin phép gửi nội dung buổi học vừa qua và yêu cầu cho tuần tới</b></p>"

            if contents["lesson_content"]:
                lines = contents["lesson_content"].splitlines()
                if lines:
                    html += f"<p><b>{lines[0]}</b></p>"
                    if len(lines) > 1:
                        html += "<p>" + "<br>".join(lines[1:]) + "</p>"

            if contents["class_performance"]:
                html += self.format_section("📌Tình hình học tập của lớp:", contents["class_performance"])
            if contents["slide_link"]:
                html += self.format_section("📌Link slide bài giảng:", contents["slide_link"])
            if contents["video_link"]:
                html += self.format_section("📌Link Video hướng dẫn:", contents["video_link"])
            if contents["homework_result"]:
                html += self.format_section("🏆Kết quả bài tập về nhà:", contents["homework_result"])
            if contents["next_requirement"]:
                html += self.format_section("🔍Yêu cầu cho buổi tiếp theo:", contents["next_requirement"])
            if contents["deadline"]:
                html += self.format_section("⏰Hạn nộp bài:", contents["deadline"])
            if contents["next_lesson_content"]:
                html += self.format_section("📌 Nội dung buổi học tới:", contents["next_lesson_content"])

            self.result_text.setHtml(html)
            QMessageBox.information(self, "Thành công", "Đã tạo nội dung thành công!")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi: {e}")

    def format_section(self, header, content):
        """Trả về chuỗi HTML cho mỗi phần nội dung"""
        section_html = f"<p><b>{header}</b></p>"
        formatted_content = content.replace('\n', '<br>')
        section_html += f"<p>{formatted_content}</p>"
        return section_html

    # --------------------- Cache ---------------------
    def load_cache(self):
        try:
            if os.path.exists(self.cache_file):
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    cache_data = json.load(f)
                last_file_path = cache_data.get('last_file_path', '')
                if last_file_path and os.path.exists(last_file_path):
                    if self.load_excel_file(last_file_path):
                        self.file_path_edit.setText(last_file_path)
                saved_states = cache_data.get('checkbox_states', {})
                for key, state in saved_states.items():
                    checkbox = getattr(self, f"{key}_checkbox", None)
                    if checkbox is not None:
                        checkbox.setChecked(state)
        except Exception as e:
            print(f"Error loading cache: {e}")

    def save_cache(self, file_path):
        try:
            cache_data = {
                'last_file_path': file_path,
                'checkbox_states': {
                    'class_performance': self.class_performance_checkbox.isChecked(),
                    'lesson_content': self.lesson_content_checkbox.isChecked(),
                    'slide_link': self.slide_link_checkbox.isChecked(),
                    'video_link': self.video_link_checkbox.isChecked(),
                    'homework_result': self.homework_result_checkbox.isChecked(),
                    'next_requirement': self.next_requirement_checkbox.isChecked(),
                    'deadline': self.deadline_checkbox.isChecked(),
                    'next_lesson_content': self.next_lesson_content_checkbox.isChecked()
                }
            }
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Error saving cache: {e}")

    # --------------------- Xử lý file Excel ---------------------
    def load_excel_file(self, file_path):
        try:
            self.excel_data = pd.read_excel(file_path, sheet_name=None)
            self.update_sheet_names()
            return True
        except Exception as e:
            print(f"Error loading Excel: {e}")
            QMessageBox.critical(self, "Lỗi", f"Không thể đọc file Excel: {e}")
            return False

    def update_sheet_names(self):
        if self.excel_data:
            self.sheet_combo.clear()
            for sheet in self.excel_data.keys():
                self.sheet_combo.addItem(sheet)
            if self.excel_data.keys():
                self.sheet_combo.setCurrentIndex(0)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Chọn file", "", "Excel Files (*.xlsx)")
        if file_path:
            if self.load_excel_file(file_path):
                self.file_path_edit.setText(file_path)
                self.save_cache(file_path)
                self.on_sheet_or_lesson_change()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = LessonDataUI()
    window.show()
    sys.exit(app.exec())