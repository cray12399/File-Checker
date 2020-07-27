from PIL import Image
from mutagen.mp3 import MP3
from mutagen.flac import FLAC
from moviepy.editor import VideoFileClip
import os
import PyPDF2
import docx2txt
import shutil
from func_timeout import func_timeout, FunctionTimedOut
import xlrd
import soundfile
import pptx
from PySide2.QtWidgets import *
from PySide2.QtCore import *
from PySide2.QtGui import *
import sys
import time


class BackgroundThread(QThread):
    thread_output = Signal(str)  # Signal used to output information to the user via the readonly textedit widget.
    thread_progress = Signal(int)  # Signal used to output progress to the progress bar widget.
    args = []
    
    def __init__(self, in_file_path, out_file_path, operation, do_extension_report):
        QThread.__init__(self)
        self.in_file_path = in_file_path
        self.out_file_path = out_file_path
        self.operation = operation
        self.do_extension_report = do_extension_report

    def __del__(self):
        self.wait()

    def run(self):
        verify_files(self, self.in_file_path, self.out_file_path, self.operation, self.do_extension_report)


class Window(QWidget):

    def __init__(self):
        super().__init__()

        width = 600
        height = 400

        self.setFixedWidth(width)
        self.setFixedHeight(height)
        self.setWindowTitle("Verify Files")
        self.create_layouts()
        self.create_ui()
        self.show()

        self.running_background = False

    def create_layouts(self):
        ######### MAIN LAYOUTS #########
        main_layout = QVBoxLayout()
        file_layout = QVBoxLayout()
        config_layout = QVBoxLayout()

        self.output_layout = QVBoxLayout()
        self.btn_layout = QHBoxLayout()

        main_layout.addLayout(file_layout)
        main_layout.addLayout(config_layout)
        main_layout.addLayout(self.output_layout)
        main_layout.addLayout(self.btn_layout)

        ######### SUB LAYOUTS #########
        self.in_file_layout = QHBoxLayout()
        self.out_file_layout = QHBoxLayout()
        self.copy_move_layout = QHBoxLayout()
        self.ext_report_layout = QHBoxLayout()

        file_layout.addLayout(self.in_file_layout)
        file_layout.addLayout(self.out_file_layout)

        config_layout.addLayout(self.copy_move_layout)
        config_layout.addLayout(self.ext_report_layout)

        self.copy_move_layout.setAlignment(Qt.AlignRight)
        self.ext_report_layout.setAlignment(Qt.AlignRight)

        self.copy_move_layout.setContentsMargins(0, 0, 8, 0)

        self.setLayout(main_layout)

    def create_ui(self):
        ######### FILE UI #########
        def in_file_function():
            in_file_field.setText(str(QFileDialog.getExistingDirectory(self, "In File Path")))

        def out_file_function():
            out_file_field.setText(str(QFileDialog.getExistingDirectory(self, "Out File Path")))

        in_file_label = QLabel("In File Path:")
        in_file_field = QLineEdit()
        in_file_btn = QPushButton("File...")

        out_file_label = QLabel("Out File Path:")
        out_file_field = QLineEdit()
        out_file_btn = QPushButton("File...")

        self.in_file_layout.addWidget(in_file_label)
        self.in_file_layout.addWidget(in_file_field)
        self.in_file_layout.addWidget(in_file_btn)

        self.out_file_layout.addWidget(out_file_label)
        self.out_file_layout.addWidget(out_file_field)
        self.out_file_layout.addWidget(out_file_btn)

        in_file_btn.setFixedWidth(45)
        in_file_btn.clicked.connect(in_file_function)

        out_file_btn.setFixedWidth(45)
        out_file_btn.clicked.connect(out_file_function)

        in_file_field.setFixedWidth(450)
        out_file_field.setFixedWidth(450)

        ######### CONFIG UI ########
        copy_radio = QRadioButton("Copy")
        move_radio = QRadioButton("Move")
        none_radio = QRadioButton("NA")
        ext_report_check = QCheckBox("Generate Extension Report")

        self.copy_move_layout.addWidget(copy_radio)
        self.copy_move_layout.addWidget(move_radio)
        self.copy_move_layout.addWidget(none_radio)

        copy_radio.setChecked(True)

        self.ext_report_layout.addWidget(ext_report_check)

        ########## OUTPUT UI ##########
        def check_for_finish():
            if self.background_thread.isFinished():
                stop_function()

        self.output_text = QTextEdit()
        self.progress_bar = QProgressBar()
        self.thread_timer = QTimer()

        self.output_layout.addWidget(self.output_text)
        self.output_layout.addWidget(self.progress_bar)

        self.output_text.setReadOnly(True)
        self.output_text.setLineWrapMode(QTextEdit.NoWrap)

        self.thread_timer.setInterval(1000)
        self.thread_timer.timeout.connect(check_for_finish)

        ######### START/STOP BUTTON UI #########
        def start_function():
            overwrite = True

            if not os.path.isdir(in_file_field.text()):
                self.update_output("ERROR: In file path does not exist!", (255, 0, 0))
            if not os.path.isdir(out_file_field.text()):
                self.update_output("ERROR: Out file path does not exist!", (255, 0, 0))
                # I wanted to make the program create the path, but doing so was very finnicky with admin privileges.

            if len(os.listdir(out_file_field.text())) != 0:
                self.update_output("ERROR: Out file path not empty!", (255, 0, 0))
                overwrite = False
                # I make the user delete the folder manually to avoid accidental overwrites.

            if os.path.isdir(in_file_field.text()) and os.path.isdir(out_file_field.text()) and overwrite:
                self.stop_thread = False
                self.running_background = True

                in_file_field.setDisabled(self.running_background)
                out_file_field.setDisabled(self.running_background)
                in_file_btn.setDisabled(self.running_background)
                out_file_btn.setDisabled(self.running_background)
                copy_radio.setDisabled(self.running_background)
                move_radio.setDisabled(self.running_background)
                none_radio.setDisabled(self.running_background)
                ext_report_check.setDisabled(self.running_background)
                btn_start.setDisabled(self.running_background)

                btn_stop.setText("Cancel")

                if copy_radio.isChecked():
                    operation = 1
                elif move_radio.isChecked():
                    operation = 2
                else:
                    operation = None

                self.background_thread = BackgroundThread(in_file_field.text(), out_file_field.text(), operation,
                                                          ext_report_check.isChecked())
                self.background_thread.thread_output.connect(self.update_output)
                self.background_thread.thread_progress.connect(self.update_progress)
                self.background_thread.start()
                self.thread_timer.start()

                self.running_background = True

        def stop_function():
            if not self.running_background:
                sys.exit()
            else:
                if not self.background_thread.isFinished():
                    alert = QMessageBox.question(
                        self, "Are you sure?", "Are you sure you wish to stop the running process?",
                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                else:
                    alert = QMessageBox.Yes

                if alert == QMessageBox.Yes:
                    self.background_thread.terminate()
                    self.thread_timer.stop()
                    self.running_background = False

                    in_file_field.setDisabled(self.running_background)
                    out_file_field.setDisabled(self.running_background)
                    in_file_btn.setDisabled(self.running_background)
                    out_file_btn.setDisabled(self.running_background)
                    copy_radio.setDisabled(self.running_background)
                    move_radio.setDisabled(self.running_background)
                    none_radio.setDisabled(self.running_background)
                    ext_report_check.setDisabled(self.running_background)
                    btn_start.setDisabled(self.running_background)

                    self.progress_bar.setValue(0)

                    if not self.background_thread.isFinished():
                        self.update_output("\nERROR: Process Canceled\n", (255, 0, 0))

                    btn_stop.setText("Close")

        btn_start = QPushButton("Start")
        btn_stop = QPushButton("Close")

        self.btn_layout.addWidget(btn_start)
        self.btn_layout.addWidget(btn_stop)

        btn_start.clicked.connect(start_function)
        btn_stop.clicked.connect(stop_function)

    def update_output(self, text_to_output, color=(0, 0, 0)):
        self.output_text.setTextColor(QColor(color[0], color[1], color[2]))
        self.output_text.append(text_to_output)
        self.output_text.setTextColor(QColor(0, 0, 0))

    def update_progress(self, progress):
        self.progress_bar.setValue(progress)


def pres_verify(*args):
    try:
        pptx.Presentation(args[0])
        args[1].append(args[0])
    except:
        args[2].append(args[0])

    return args[1], args[2]


def image_verify(*args):
    try:
        image = Image.open(args[0])  # Tries to open image and verify it.
        image.verify()

        args[1].append(args[0])  # If successful, appends to the good file list.
    except:
        args[2].append(args[0])  # If not, appends to the bad file list.

    return args[1], args[2]  # Returns updated lists


# *** All of the verification functions follow this basic algorithm. ***

# Args were used because they are required in the timeout library function that I used.

# arg[0] = The file that is being checked
# arg[1] = The good files list
# arg[2] = The bad files list


def pdf_verify(*args):
    try:
        PyPDF2.PdfFileReader(open(args[0], "rb"))
        args[1].append(args[0])
    except:
        args[2].append(args[0])

    return args[1], args[2]


def docx_verify(*args):
    try:
        docx2txt.process(args[0])
        args[1].append(args[0])
    except:
        args[2].append(args[0])

    return args[1], args[2]


def mp3_verify(*args):
    try:
        MP3(args[0])
        args[1].append(args[0])
    except:
        args[2].append(args[0])

    return args[1], args[2]


def movie_verify(*args):
    try:
        video = VideoFileClip(args[0])
        args[1].append(args[0])
        video.close()
    except:
        args[2].append(args[0])

    return args[1], args[2]


def excel_verify(*args):
    try:
        xlrd.open_workbook(args[0])
        args[1].append(args[0])
    except:
        args[2].append(args[0])

    return args[1], args[2]


def ogg_verify(*args):
    try:
        soundfile.read(args[0])
        args[1].append(args[0])
    except:
        args[2].append(args[0])

    return args[1], args[2]


def flac_verify(*args):
    try:
        FLAC(args[0])
        args[1].append(args[0])
    except:
        args[2].append(args[0])

    return args[1], args[2]


def make_reports(out_file_path, good_files, bad_files, all_files, neutral_files):
    if not os.path.isdir(f"{out_file_path}\\good_files") and len(good_files) != 0:
        os.mkdir(f"{out_file_path}\\good_files")
    if not os.path.isdir(f"{out_file_path}\\bad_files") and len(bad_files) != 0:
        os.mkdir(f"{out_file_path}\\bad_files")
    if not os.path.isdir(f"{out_file_path}\\neutral_files") and len(neutral_files) != 0:
        os.mkdir(f"{out_file_path}\\neutral_files")
    # Creating the good/bad/neutral folders so that it can put the file reports in them

    with open(f"{out_file_path}\\all_files.txt", "w+", encoding="utf-8") as all_file_report:
        for file in all_files:
            try:
                all_file_report.write(file + "\n")
            except Exception as e:
                error_log(out_file_path, "make_reports", e)

    if len(good_files) != 0:
        with open(f"{out_file_path}\\good_files\\good_files.txt", "w+", encoding="utf-8") as good_file_report:
            for file in good_files:
                try:
                    good_file_report.write(file + "\n")
                except Exception as e:
                    error_log(out_file_path, "make_reports", e)

    if len(bad_files) != 0:
        with open(f"{out_file_path}\\bad_files\\bad_files.txt", "w+", encoding="utf-8") as bad_file_report:
            for file in bad_files:
                try:
                    bad_file_report.write(file + "\n")
                except Exception as e:
                    error_log(out_file_path, "make_reports", e)

    if len(neutral_files) != 0:
        with open(f"{out_file_path}\\neutral_files\\neutral_files.txt", "w+", encoding="utf-8") as bad_file_report:
            for file in neutral_files:
                try:
                    bad_file_report.write(file + "\n")
                except Exception as e:
                    error_log(out_file_path, "make_reports", e)


def copy_files(thread, out_file_path, good_files, bad_files, neutral_files):
    output_print = lambda output_text: thread.thread_output.emit(output_text)
    output_progress = lambda percent_done: thread.thread_progress.emit(percent_done)

    delimiter = "\\"  # Since format strings do not allow the \\ symbol, I must encapsulate it in a variable
    total_files = len(good_files) + len(bad_files) + len(neutral_files)
    copied_files = 0

    output_print("\nCopying good files...\n")

    for file in good_files:
        try:
            path = [out_file_path, "good_files"]

            for folder in file.split(delimiter)[1:-1]:

                path.append(folder)

                if not os.path.isdir(delimiter.join(path)):
                    os.mkdir(delimiter.join(path))
        except Exception as e:
            error_log(out_file_path, "copy_dirs", e)

    for file in good_files:
        output_print(f"{file} - {get_multiple_name(int(os.stat(file).st_size))}")

        try:
            shutil.copy(file, f"{out_file_path}\\good_files\\{delimiter.join(file.split(delimiter)[1:])}")
        except Exception as e:
            error_log(out_file_path, "copy_files", e)

        copied_files += 1

        output_progress(int(copied_files / total_files * 100))

    output_print("\nCopying bad files...\n")

    for file in bad_files:
        try:
            path = [out_file_path, "bad_files"]

            for folder in file.split(delimiter)[1:-1]:

                path.append(folder)

                if not os.path.isdir(delimiter.join(path)):
                    os.mkdir(delimiter.join(path))
        except Exception as e:
            error_log(out_file_path, "copy_dirs", e)

    for file in bad_files:
        output_print(f"{file} - {get_multiple_name(int(os.stat(file).st_size))}")

        try:
            shutil.copy(file, f"{out_file_path}\\bad_files\\{delimiter.join(file.split(delimiter)[1:])}")
        except Exception as e:
            error_log(out_file_path, "copy_files", e)

        copied_files += 1

        output_progress(int(copied_files / total_files * 100))

    output_print("\nCopying neutral files...\n")

    for file in neutral_files:
        try:
            path = [out_file_path, "neutral_files"]

            for folder in file.split(delimiter)[1:-1]:

                path.append(folder)

                if not os.path.isdir(delimiter.join(path)):
                    os.mkdir(delimiter.join(path))
        except Exception as e:
            error_log(out_file_path, "copy_dirs", e)

    for file in neutral_files:
        output_print(f"{file} - {get_multiple_name(int(os.stat(file).st_size))}")

        try:
            shutil.copy(file, f"{out_file_path}\\neutral_files\\{delimiter.join(file.split(delimiter)[1:])}")
        except Exception as e:
            error_log(out_file_path, "copy_files", e)

        copied_files += 1

        output_progress(int(copied_files / total_files * 100))

    output_progress(0)

    output_print("Done copying!")


def calc_size(out_file_path, file_list):
    byte = 0

    for file in file_list:
        try:
            byte += int(os.stat(file).st_size)
        except Exception as e:
            error_log(out_file_path, "calc_size", e)

    return byte


def get_multiple_name(size_in_bytes):
    if size_in_bytes >= 1000000000000:
        return str(f"{round((size_in_bytes / 1000000000000), 2)} TB")
    elif size_in_bytes >= 1000000000:
        return str(f"{round((size_in_bytes / 1000000000), 2)} GB")
    elif size_in_bytes >= 1000000:
        return str(f"{round((size_in_bytes / 1000000), 2)} MB")
    elif size_in_bytes >= 1000:
        return str(f"{round((size_in_bytes / 1000), 2)} KB")
    elif size_in_bytes == 0:
        return str(f"{size_in_bytes} B")
    else:
        return "Unknown Size"


def error_log(out_file_path, comment, exception):
    try:
        if not os.path.isfile(f"{out_file_path}\\errors.log"):
            open(f"{out_file_path}\\errors.log", "w+", encoding="utf-8").close()

        with open(f"{out_file_path}\\errors.log", "a+", encoding="utf-8") as log:
            log.write(f"{comment} - {exception}\n")
    except:
        pass


def make_extension_report(out_file_path, good_files, bad_files):
    extension_counts = {"docx": [0, 0], "xlsx": [0, 0], "pdf": [0, 0], "jpg": [0, 0], "jpeg": [0, 0], "png": [0, 0],
                        "gif": [0, 0], "mp3": [0, 0], "ogg": [0, 0], "flac": [0, 0], "mpg": [0, 0], "mpeg": [0, 0],
                        "avi": [0, 0], "mp4": [0, 0], "mov": [0, 0], "xls": [0, 0], "bmp": [0, 0], "wmv": [0, 0],
                        "pptx": [0, 0]}

    for file in good_files:
        if file.split(".")[-1].lower() in extension_counts.keys():
            extension_counts[file.split(".")[-1].lower()][0] += 1

    for file in bad_files:
        if file.split(".")[-1].lower() in extension_counts.keys():
            extension_counts[file.split(".")[-1].lower()][1] += 1

    with open(f"{out_file_path}\\extension_report.txt", "w+") as file:
        file.write("Ext\tGood\tBad\n---------------------\n")

        for key, value in extension_counts.items():
            file.write(f"{key}\t{value[0]}\t{value[1]}\n")

        file.write(f"---------------------\nTotal\t{len(good_files)}\t{len(bad_files)}")


def verify_files(thread, in_file_path, out_file_path, operation, do_extension_report):
    output_print = lambda output_text: thread.thread_output.emit(output_text)
    output_progress = lambda percent_done: thread.thread_progress.emit(percent_done)

    timeout = 10

    good_files = []
    bad_files = []
    neutral_files = []
    all_files = []

    output_print("Process starting...")

    for root, dirs, files in os.walk(in_file_path):
        for file in files:

            all_files.append(os.path.join(root, file))

    output_print("Checking files...")

    for file_index, file in enumerate(all_files):
        output_print(f"{file} - {get_multiple_name(int(os.stat(file).st_size))}")

        if len(file.split(".")) > 1:

            if file.split(".")[-1].lower() in ["jpeg", "jpg", "png", "gif", "bmp"]:
                try:
                    good_files, bad_files = func_timeout(timeout, image_verify, args=(file, good_files, bad_files))
                except FunctionTimedOut:
                    bad_files.append(file)
                except Exception as e:
                    error_log(out_file_path, file, e)

            elif file.split(".")[-1].lower() == "pdf":
                try:
                    good_files, bad_files = func_timeout(timeout, pdf_verify, args=(file, good_files, bad_files))
                except FunctionTimedOut:
                    bad_files.append(file)
                except Exception as e:
                    error_log(out_file_path, file, e)

            elif file.split(".")[-1].lower() == "docx":
                try:
                    good_files, bad_files = func_timeout(timeout, docx_verify, args=(file, good_files, bad_files))
                except FunctionTimedOut:
                    bad_files.append(file)
                except Exception as e:
                    error_log(out_file_path, file, e)

            elif file.split(".")[-1].lower() == "mp3":
                try:
                    good_files, bad_files = func_timeout(timeout, mp3_verify, args=(file, good_files, bad_files))
                except FunctionTimedOut:
                    bad_files.append(file)
                except Exception as e:
                    error_log(out_file_path, file, e)

            elif file.split(".")[-1].lower() == "ogg":
                try:
                    good_files, bad_files = func_timeout(timeout, ogg_verify, args=(file, good_files, bad_files))
                except FunctionTimedOut:
                    bad_files.append(file)
                except Exception as e:
                    error_log(out_file_path, file, e)

            elif file.split(".")[-1].lower() == "flac":
                try:
                    good_files, bad_files = func_timeout(timeout, flac_verify, args=(file, good_files, bad_files))
                except FunctionTimedOut:
                    bad_files.append(file)
                except Exception as e:
                    error_log(out_file_path, file, e)

            elif file.split(".")[-1].lower() in ["mp4", "avi", "mov", "wmv", "mts", "mpg"]:
                try:
                    good_files, bad_files = func_timeout(timeout, movie_verify, args=(file, good_files, bad_files))
                except FunctionTimedOut:
                    bad_files.append(file)
                except Exception as e:
                    error_log(out_file_path, file, e)

            elif file.split(".")[-1].lower() == "xlsx":
                try:
                    good_files, bad_files = func_timeout(timeout, excel_verify, args=(file, good_files, bad_files))
                except FunctionTimedOut:
                    bad_files.append(file)
                except Exception as e:
                    error_log(out_file_path, file, e)

            elif file.split(".")[-1].lower() == "pptx":
                try:
                    good_files, bad_files = func_timeout(timeout, pres_verify, args=(file, good_files, bad_files))
                except FunctionTimedOut:
                    bad_files.append(file)
                except Exception as e:
                    error_log(out_file_path, file, e)

            else:
                neutral_files.append(file)

            output_progress(int(file_index / len(all_files) * 100))

    output_progress(0)

    if do_extension_report:
        make_extension_report(out_file_path, good_files, bad_files)

    output_print("\nGenerating report..." + "\r")

    report = f"Good Files: {len(good_files)} - {get_multiple_name(calc_size(out_file_path, good_files))}\n" \
             f"Bad Files: {len(bad_files)} - {get_multiple_name(calc_size(out_file_path, bad_files))}\n" \
             f"Neutral Files: {len(neutral_files)} - {get_multiple_name(calc_size(out_file_path, neutral_files))}\n\n" \
             f"All Files: {len(all_files)} - {get_multiple_name(calc_size(out_file_path, all_files))}\n\n" \

    with open(f"{out_file_path}\\all_files.txt", "a+") as file:
        file.write("\n-------------------------\n")
        file.write(f"{report}")
        file.write("\n-------------------------\n")

    make_reports(out_file_path, good_files, bad_files, all_files, neutral_files)

    output_print("-------------------------\n")

    if operation == 1:
        copy_files(thread, out_file_path, good_files, bad_files, neutral_files)
        output_print("\n-------------------------")

    elif operation == 2:
        copy_files(thread, out_file_path, good_files, bad_files, neutral_files)
        shutil.rmtree(in_file_path)
        time.sleep(.000000001)  # Shutil.rmtree occasionally locks the file manager, so this code unlocks it.

        output_print("\n-------------------------")

    output_print(report)


def main():
    App = QApplication(sys.argv)
    window = Window()
    sys.exit(App.exec_())


if __name__ == '__main__':
    main()
