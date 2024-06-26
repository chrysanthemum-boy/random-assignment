import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QTextEdit, QFileDialog, QLabel, \
    QProgressBar, QMessageBox
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import random
import pandas as pd


class ExcelImportThread(QThread):
    progress_changed = pyqtSignal(int)

    def __init__(self, file_name):
        super().__init__()
        self.file_name = file_name

    def run(self):
        try:
            df = pd.read_excel(self.file_name)
            print("ok")
            self.progress_changed.emit(100)  # 通知进度值已经到达100
        except Exception as e:
            self.progress_changed.emit(-1)  # 发送错误信号
            return


class MentorAssignmentApp(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()
        self.students_df = None
        self.mentors_df = None

    def initUI(self):
        self.setWindowTitle('Mentor Assignment App')
        self.setGeometry(100, 100, 400, 500)

        self.text_edit = QTextEdit(self)
        self.text_edit.setReadOnly(True)

        self.assign_button = QPushButton('点击随机分配导师', self)
        self.assign_button.clicked.connect(self.assignMentors)

        self.import_students_button = QPushButton('导入学生信息（Excel）', self)
        self.import_students_button.clicked.connect(self.importStudentsData)

        # self.progress_label = QLabel('Progress:', self)
        # self.progress_bar = QProgressBar(self)
        # self.progress_bar.setAlignment(Qt.AlignCenter)

        self.import_mentors_button = QPushButton('导入导师信息（Excel）', self)
        self.import_mentors_button.clicked.connect(self.importMentorsData)

        # self.progress_label = QLabel('Progress:', self)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setAlignment(Qt.AlignCenter)

        # self.progress_label2 = QLabel('Progress:', self)
        self.progress_bar2 = QProgressBar(self)
        self.progress_bar2.setAlignment(Qt.AlignCenter)

        layout = QVBoxLayout()
        layout.addWidget(self.text_edit)

        layout.addWidget(self.import_students_button)
        layout.addWidget(self.progress_bar)

        layout.addWidget(self.import_mentors_button)
        layout.addWidget(self.progress_bar2)

        layout.addWidget(self.assign_button)
        self.setLayout(layout)

    def importStudentsData(self):
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("Excel files (*.xlsx)")
        file_dialog.selectNameFilter("Excel files (*.xlsx)")
        file_dialog.setDefaultSuffix("xlsx")
        file_dialog.setFileMode(QFileDialog.ExistingFiles)

        if file_dialog.exec_():
            file_names = file_dialog.selectedFiles()
            for file_name in file_names:
                self.import_thread = ExcelImportThread(file_name)
                self.import_thread.progress_changed.connect(self.updateProgressBar)
                self.import_thread.start()
                self.students_df = pd.read_excel(file_name)

    def importMentorsData(self):
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("Excel files (*.xlsx)")
        file_dialog.selectNameFilter("Excel files (*.xlsx)")
        file_dialog.setDefaultSuffix("xlsx")
        file_dialog.setFileMode(QFileDialog.ExistingFiles)

        if file_dialog.exec_():
            print("exec")
            file_names = file_dialog.selectedFiles()
            for file_name in file_names:
                self.import_thread = ExcelImportThread(file_name)
                self.import_thread.progress_changed.connect(self.updateProgressBar2)
                self.import_thread.start()
                self.mentors_df = pd.read_excel(file_name)
        # print("ok")

    def assignMentors(self):
        random.seed(100)

        students = self.students_df['姓名'].tolist()
        mentors = self.mentors_df['姓名'].tolist()
        random.shuffle(mentors)
        mentor_assignments = {mentor: [] for mentor in mentors}

        random.shuffle(students)

        res = {
            "学生姓名": [],
            "导师姓名": []
        }

        for student in students:
            res["学生姓名"].append(student)
            min_assigned_mentor = min(mentor_assignments, key=lambda mentor: len(mentor_assignments[mentor]))
            mentor_assignments[min_assigned_mentor].append(student)
            res["导师姓名"].append(min_assigned_mentor)

        print(res)
        df = pd.DataFrame(res)

        df.to_excel('学生导师分配名单.xlsx', index=False)
        self.text_edit.setPlainText("随机名单生成成功！")

    def updateProgressBar(self, value):
        if value == -1:
            QMessageBox.critical(self, 'Error', '导入失败！')
        else:
            self.progress_bar.setValue(value)
            if value == 100:
                QMessageBox.information(self, 'Success', 'Excel导入成功!')

    def updateProgressBar2(self, value):
        if value == -1:
            QMessageBox.critical(self, 'Error', '导入失败！')
        else:
            self.progress_bar2.setValue(value)
            if value == 100:
                QMessageBox.information(self, 'Success', 'Excel导入成功!')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mentor_app = MentorAssignmentApp()
    mentor_app.show()
    sys.exit(app.exec_())
