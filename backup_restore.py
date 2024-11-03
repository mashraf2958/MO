import sys
import os
import win32com.client
import subprocess
import argparse
import psycopg2
import json
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QLineEdit, QComboBox, QStackedWidget, QFileDialog, QMessageBox, 
    QProgressBar, QTabWidget, QTimeEdit, QSpinBox, QRadioButton, 
    QCheckBox, QGroupBox, QTextEdit, QScrollArea, QStyleFactory, QComboBox, QDateEdit, QFormLayout, QDialog, QDialogButtonBox, QGridLayout,
    QListWidget, QListWidgetItem
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDate, QUrl, QSize
from PyQt5.QtGui import QIcon, QFont, QPixmap, QDesktopServices
import tempfile
import textwrap
import smtplib
import platform
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class NoScrollComboBox(QComboBox):
    def wheelEvent(self, event):
        event.ignore()

class BackupThread(QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(self, backup_type, file_extension, db_name, db_host, db_port, db_user, db_password, base_backup_dir):
        QThread.__init__(self)
        self.backup_type = backup_type
        self.file_extension = file_extension
        self.db_name = db_name
        self.db_host = db_host
        self.db_port = db_port
        self.db_user = db_user
        self.db_password = db_password
        self.base_backup_dir = base_backup_dir

    def run(self):
        try:
            if self.db_name:
                success = self.backup_database(self.backup_type, self.file_extension, self.db_name)
            else:
                success = self.backup_all_databases(self.backup_type, self.file_extension)
            self.finished.emit(success, "Backup completed successfully." if success else "Backup failed.")
        except Exception as e:
            self.finished.emit(False, f"An error occurred: {str(e)}")

    def find_pg_dump(self):
        # Get the absolute path of the script/executable
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            base_path = sys._MEIPASS
        else:
            # Running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        pg_dump_path = os.path.join(base_path, 'resources', 'bin', 'pg_dump.exe')
        
        if os.path.exists(pg_dump_path):
            return pg_dump_path
        
        # Fallback to system paths if embedded executable not found
        possible_pg_paths = [
            r"D:\SETUP PROGRAMS\PostgreSQL\16\bin\pg_dump.exe",
            r"C:\Program Files\PostgreSQL\15\bin\pg_dump.exe",
            r"C:\Program Files\PostgreSQL\14\bin\pg_dump.exe",
            r"C:\Program Files\PostgreSQL\13\bin\pg_dump.exe",
            r"C:\Program Files\PostgreSQL\12\bin\pg_dump.exe"
        ]
        
        for path in possible_pg_paths:
            if os.path.exists(path):
                return path
                
        raise FileNotFoundError("pg_dump.exe not found in embedded resources or system paths.")

    def backup_database(self, backup_type, file_extension, db_name):
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        pg_dump_path = self.find_pg_dump()
        temp_backup_file = os.path.join(self.base_backup_dir, f"{db_name}.{file_extension}")

        try:
            conn = psycopg2.connect(dbname=db_name, user=self.db_user, password=self.db_password, 
                                    host=self.db_host, port=self.db_port)
            conn.set_session(autocommit=True)
            cursor = conn.cursor()

            self.status.emit(f"Backing up roles for {db_name}")
            with open(temp_backup_file, "w", encoding='utf-8') as f:
                cursor.execute("""
                    SELECT r.rolname, r.rolsuper, r.rolinherit, r.rolcreaterole,
                           r.rolcreatedb, r.rolcanlogin, r.rolpassword
                    FROM pg_authid r JOIN pg_roles u ON r.oid = u.oid;
                """)
                roles = cursor.fetchall()
                for role in roles:
                    rolename, rolsuper, rolinherit, rolcreaterole, rolcreatedb, rolcanlogin, rolpassword = role
                    f.write(f'CREATE ROLE "{rolename}" WITH ')
                    if rolsuper:
                        f.write("SUPERUSER ")
                    if not rolinherit:
                        f.write("NOINHERIT ")
                    if rolcreaterole:
                        f.write("CREATEROLE ")
                    if rolcreatedb:
                        f.write("CREATEDB ")
                    if rolcanlogin:
                        f.write("LOGIN ")
                    if rolpassword:
                        f.write(f"ENCRYPTED PASSWORD '{rolpassword}' ")
                    f.write(";\n")

            self.status.emit(f"Backing up database {db_name}")
            os.environ['PGPASSWORD'] = self.db_password
            pg_dump_cmd = [
                pg_dump_path,
                "-h", self.db_host,
                "-p", self.db_port,
                "-U", self.db_user,
                "-d", db_name,
                "-F", "p"  # Plain text format for SQL output
            ]
            if backup_type == 'Schema':
                pg_dump_cmd.append("-s")  # Schema-only

            process = subprocess.Popen(pg_dump_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, 
                                       universal_newlines=True, encoding='utf-8')
            
            with open(temp_backup_file, "a", encoding='utf-8') as f:
                while True:
                    output = process.stdout.readline()
                    if output == '' and process.poll() is not None:
                        break
                    if output:
                        f.write(output)
                        self.progress.emit(50)  # Assuming 50% progress for simplicity

            if backup_type == 'Schema':
                backup_dir = os.path.join(self.base_backup_dir, self.db_host, f"schema_{db_name}", timestamp)
            else:
                backup_dir = os.path.join(self.base_backup_dir, self.db_host, db_name, timestamp)

            os.makedirs(backup_dir, exist_ok=True)
            final_backup_file = os.path.join(backup_dir, f"{db_name}.{file_extension}")
            os.rename(temp_backup_file, final_backup_file)

            self.progress.emit(100)
            return True

        except (psycopg2.Error, subprocess.CalledProcessError, IOError, OSError) as e:
            print(f"Error during backup of database '{db_name}': {e}")
            if os.path.exists(temp_backup_file):
                os.remove(temp_backup_file)
            return False
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def backup_all_databases(self, backup_type, file_extension):
        try:
            conn = psycopg2.connect(dbname='postgres', user=self.db_user, password=self.db_password, 
                                    host=self.db_host, port=self.db_port)
            conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
            cursor = conn.cursor()

            cursor.execute("SELECT datname FROM pg_database WHERE datistemplate = false;")
            databases = cursor.fetchall()
            total_dbs = len(databases)

            for i, db in enumerate(databases):
                db_name = db[0]
                self.status.emit(f"Backing up database: {db_name}")
                success = self.backup_database(backup_type, file_extension, db_name)
                if not success:
                    print(f"Failed to backup {db_name}")
                self.progress.emit(int((i + 1) / total_dbs * 100))

            return True
        except psycopg2.Error as e:
            print(f"Error connecting to PostgreSQL: {e}")
            return False
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

class RestoreThread(QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(self, db_host, db_port, db_user, db_password, backup_dir):
        QThread.__init__(self)
        self.db_host = db_host
        self.db_port = db_port
        self.db_user = db_user
        self.db_password = db_password
        self.backup_dir = backup_dir

    def run(self):
        try:
            self.restore_databases()
            self.finished.emit(True, "Restore completed successfully.")
        except Exception as e:
            self.finished.emit(False, f"An error occurred during restore: {str(e)}")

    def find_psql(self):
        # Get the absolute path of the script/executable
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            base_path = sys._MEIPASS
        else:
            # Running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        psql_path = os.path.join(base_path, 'resources', 'bin', 'psql.exe')
        
        if os.path.exists(psql_path):
            return psql_path
        
        # Fallback to system paths if embedded executable not found
        possible_psql_paths = [
            r"D:\SETUP PROGRAMS\PostgreSQL\16\bin\psql.exe",
            r"C:\Program Files\PostgreSQL\15\bin\psql.exe",
            r"C:\Program Files\PostgreSQL\14\bin\psql.exe",
            r"C:\Program Files\PostgreSQL\13\bin\psql.exe",
            r"C:\Program Files\PostgreSQL\12\bin\psql.exe"
        ]
        
        for path in possible_psql_paths:
            if os.path.exists(path):
                return path
                
        raise FileNotFoundError("psql.exe not found in embedded resources or system paths.")

    def restore_databases(self):
        psql_path = self.find_psql()
        os.environ['PGPASSWORD'] = self.db_password

        for root, dirs, files in os.walk(self.backup_dir):
            for file in files:
                if file.endswith('.sql') or file.endswith('.backup'):
                    db_name = os.path.splitext(file)[0]
                    backup_file = os.path.join(root, file)
                    self.status.emit(f"Restoring database: {db_name}")

                    # Create database if it doesn't exist
                    create_db_cmd = [
                        psql_path,
                        "-h", self.db_host,
                        "-p", self.db_port,
                        "-U", self.db_user,
                        "-d", "postgres",
                        "-c", f"CREATE DATABASE \"{db_name}\" WITH ENCODING 'UTF8'"
                    ]
                    subprocess.run(create_db_cmd, check=True, capture_output=True, encoding='utf-8')

                    # Restore the database
                    restore_cmd = [
                        psql_path,
                        "-h", self.db_host,
                        "-p", self.db_port,
                        "-U", self.db_user,
                        "-d", db_name,
                        "-f", backup_file
                    ]
                    process = subprocess.Popen(restore_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True, encoding='utf-8')
                    
                    while True:
                        output = process.stdout.readline()
                        if output == '' and process.poll() is not None:
                            break
                        if output:
                            self.status.emit(output.strip())
                    self.progress.emit(50)  # Update progress (you may want to adjust this)

        self.progress.emit(100)

class ModernBackupRestoreGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.dark_mode = False
        self.set_app_icon()
        self.initUI()
        self.update_statistics()
        
    def set_app_icon(self):
        app_icon = QIcon("icons/app_icon.png")
        self.setWindowIcon(app_icon)
        # Also set the taskbar icon for Windows
        if platform.system() == 'Windows':
            import ctypes
            myappid = 'mycompany.backuprestore.gui.1'  # arbitrary string
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    def initUI(self):
        self.setWindowTitle('Database Backup and Restore')
        self.setGeometry(500, 75, 650, 900)
        self.setStyleSheet(self.get_stylesheet())

        main_layout = QVBoxLayout()
        
        # Add dark mode switch button
        self.dark_mode_button = QPushButton()
        self.dark_mode_button.setFixedSize(32, 32)
        self.dark_mode_button.setStyleSheet("background-color: transparent; border: none;")
        self.dark_mode_button.clicked.connect(self.toggle_dark_mode)
        self.update_dark_mode_button()

        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(self.dark_mode_button)
        main_layout.addLayout(button_layout)

        self.stack = QStackedWidget()
        main_layout.addWidget(self.stack)

        self.main_page = self.create_main_page()
        self.backup_page = self.create_backup_page()
        self.restore_page = self.create_restore_page()

        self.stack.addWidget(self.main_page)
        self.stack.addWidget(self.backup_page)
        self.stack.addWidget(self.restore_page)

        self.setLayout(main_layout)

        QApplication.setStyle(QStyleFactory.create('Fusion'))

    def create_main_page(self):
        page = QWidget()
        layout = QVBoxLayout()

        # Add title
        title = QLabel('Database Management')
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 24px; font-weight: bold; margin-bottom: 20px;")
        layout.addWidget(title)

        # Buttons section
        buttons_layout = QVBoxLayout()
        backup_btn = self.create_button('Backup', 'backup')
        restore_btn = self.create_button('Restore', 'restore')
        exit_btn = self.create_button('Exit', 'exit')

        backup_btn.clicked.connect(lambda: self.stack.setCurrentIndex(1))
        restore_btn.clicked.connect(lambda: self.stack.setCurrentIndex(2))
        exit_btn.clicked.connect(self.close)

        buttons_layout.addWidget(backup_btn)
        buttons_layout.addWidget(restore_btn)
        buttons_layout.addWidget(exit_btn)
        layout.addLayout(buttons_layout)

        layout.addStretch(1)

        # Logo section
        logo_label = QLabel()
        logo_pixmap = QPixmap("icons/logo.png")
        scaled_pixmap = logo_pixmap.scaled(550, 550, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        logo_label.setPixmap(scaled_pixmap)
        logo_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(logo_label)

        layout.addStretch(1)

        # Copyright text
        copyright_text = QLabel('© 2024 Mohamed Ashraf. All rights reserved.')
        copyright_text.setAlignment(Qt.AlignCenter)
        copyright_text.setStyleSheet("""
            color: #666666;
            font-size: 15px;
            margin-top: 10px;
            margin-bottom: 5px;
        """)
        layout.addWidget(copyright_text)

        # Social Media Container
        social_container = QWidget()
        social_layout = QHBoxLayout(social_container)
        social_layout.setAlignment(Qt.AlignCenter)
        social_layout.setSpacing(15)
        social_layout.setContentsMargins(0, 0, 0, 0)
        
        icon_size = 24
        social_container.setFixedHeight(int(icon_size * 1.4))
        
        icon_style = """
            QPushButton {
                border: none;
                outline: none;
                background-color: transparent;
                padding: 3px;
            }
            QPushButton:hover {
                background-color: transparent;
            }
            QPushButton:focus {
                border: none;
                outline: none;
            }
        """

        class HoverButton(QPushButton):
            def __init__(self):
                super().__init__()
                self.setStyleSheet(icon_style)
                self.default_size = QSize(icon_size, icon_size)
                self.hover_size = QSize(int(icon_size * 1.2), int(icon_size * 1.2))
                self.setIconSize(self.default_size)
                self.setFixedSize(int(icon_size * 1.4), int(icon_size * 1.4))

            def enterEvent(self, event):
                self.setIconSize(self.hover_size)

            def leaveEvent(self, event):
                self.setIconSize(self.default_size)

        # Social media buttons
        social_links = [
            ("email.png", "mailto:m.ashraf55@outlook.com"),
            ("facebook.png", "https://www.facebook.com/Mohamed.Ashraf.MAS"),
            ("linkedin.png", "https://www.linkedin.com/in/Mohamed-Ashraf-MAS"),
            ("phone.png", "tel:+201550601240")
        ]

        for icon, url in social_links:
            btn = HoverButton()
            btn.setIcon(QIcon(f"icons/{icon}"))
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda checked, url=url: QDesktopServices.openUrl(QUrl(url)))
            social_layout.addWidget(btn)

        layout.addWidget(social_container)
        page.setLayout(layout)
        return page
    

    def create_backup_page(self):
        page = QWidget()
        layout = QVBoxLayout()

        title = QLabel('Database Backup')
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: bold; margin-bottom: 20px;")
        layout.addWidget(title)

        tab_widget = QTabWidget()
        tab_widget.addTab(self.create_manual_backup_tab(), "Manual Backup")
        tab_widget.addTab(self.create_scheduled_backup_tab(), "Scheduled Backup")
        tab_widget.addTab(self.create_schedule_management_tab(), "Schedule Management")
        layout.addWidget(tab_widget)

        page.setLayout(layout)
        return page

    def create_manual_backup_tab(self):
        tab = QWidget()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(tab)
        layout = QVBoxLayout()
        tab.setLayout(layout)

        self.backup_type = self.create_combobox(['Full Backup', 'Schema-only Backup'])
        self.file_extension = self.create_combobox(['.backup', '.sql'])
        layout.addWidget(QLabel('Backup Type'))
        layout.addWidget(self.backup_type)
        layout.addWidget(QLabel('File Extension'))
        layout.addWidget(self.file_extension)

        self.db_host = self.create_line_edit('localhost')
        self.db_port = self.create_line_edit('5432')
        self.db_user = self.create_line_edit('postgres')
        self.db_password = self.create_line_edit()
        self.db_password.setEchoMode(QLineEdit.Password)
        self.db_name = self.create_line_edit()
        self.backup_dir = self.create_line_edit()

        layout.addWidget(QLabel('Host'))
        layout.addWidget(self.db_host)
        layout.addWidget(QLabel('Port'))
        layout.addWidget(self.db_port)
        layout.addWidget(QLabel('User'))
        layout.addWidget(self.db_user)
        layout.addWidget(QLabel('Password'))
        layout.addWidget(self.db_password)
        layout.addWidget(QLabel('Database Name (leave blank for all)'))
        layout.addWidget(self.db_name)
        layout.addWidget(QLabel('Backup Directory'))

        backup_dir_layout = QHBoxLayout()
        backup_dir_layout.addWidget(self.backup_dir)
        browse_backup_btn = QPushButton('Browse')
        browse_backup_btn.clicked.connect(self.browse_backup_dir)
        backup_dir_layout.addWidget(browse_backup_btn)
        layout.addLayout(backup_dir_layout)

        progress_layout = QHBoxLayout()
        self.backup_progress = QProgressBar()
        self.backup_progress.setTextVisible(False)
        self.backup_percentage = QLabel('0%')
        progress_layout.addWidget(self.backup_progress)
        progress_layout.addWidget(self.backup_percentage)
        layout.addLayout(progress_layout)

        self.backup_status = QLabel('')
        layout.addWidget(self.backup_status)

        btn_layout = QHBoxLayout()
        back_btn = self.create_button('Back', 'back')
        backup_btn = self.create_button('Backup', 'start_backup')
        back_btn.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        backup_btn.clicked.connect(self.perform_manual_backup)
        btn_layout.addWidget(back_btn)
        btn_layout.addWidget(backup_btn)
        layout.addLayout(btn_layout)

        # Apply combobox style to all relevant widgets
        self.apply_combobox_style(self.db_host)
        self.apply_combobox_style(self.db_port)
        self.apply_combobox_style(self.db_user)
        self.apply_combobox_style(self.db_password)
        self.apply_combobox_style(self.db_name)
        self.apply_combobox_style(self.backup_dir)

        tab.setLayout(layout)
        return scroll

    def create_scheduled_backup_tab(self):
        tab = QWidget()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(tab)
        layout = QVBoxLayout()
        tab.setLayout(layout)

        # Task Details Group
        task_group = QGroupBox("Task Details")
        task_layout = QFormLayout()
        task_group.setLayout(task_layout)

        self.task_name_input = QLineEdit()
        self.task_name_input.setPlaceholderText("Enter task name")
        task_layout.addRow("Task Name:", self.task_name_input)

        self.description_input = QTextEdit()
        self.description_input.setPlaceholderText("Enter task description")
        self.description_input.setMaximumHeight(60)
        task_layout.addRow("Description:", self.description_input)

        layout.addWidget(task_group)

        # Schedule Group
        schedule_group = QGroupBox("Schedule")
        schedule_layout = QFormLayout()
        schedule_group.setLayout(schedule_layout)

        self.schedule_interval = self.create_combobox(['Daily', 'Weekly', 'Monthly'])
        self.schedule_interval.currentTextChanged.connect(self.update_schedule_options)
        schedule_layout.addRow("Interval:", self.schedule_interval)

        self.schedule_time = QTimeEdit()
        self.schedule_time.setDisplayFormat("HH:mm")
        schedule_layout.addRow("Time:", self.schedule_time)

        self.schedule_day = QSpinBox()
        self.schedule_day.setMinimum(1)
        self.schedule_day.setMaximum(31)
        self.schedule_day.setEnabled(False)
        schedule_layout.addRow("Day of Month:", self.schedule_day)

        self.schedule_weekday = self.create_combobox(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'])
        self.schedule_weekday.setEnabled(False)
        schedule_layout.addRow("Day of Week:", self.schedule_weekday)

        self.repetition_spinbox = QSpinBox()
        self.repetition_spinbox.setRange(0, 1440)
        self.repetition_spinbox.setSuffix(" minutes")
        schedule_layout.addRow("Repeat every:", self.repetition_spinbox)

        date_layout = QHBoxLayout()
        self.start_date = QDateEdit()
        self.start_date.setDate(QDate.currentDate())
        self.start_date.setCalendarPopup(True)
        self.end_date = QDateEdit()
        self.end_date.setDate(QDate.currentDate().addYears(1))
        self.end_date.setCalendarPopup(True)
        date_layout.addWidget(QLabel("Start:"))
        date_layout.addWidget(self.start_date)
        date_layout.addWidget(QLabel("End:"))
        date_layout.addWidget(self.end_date)
        schedule_layout.addRow("Date Range:", date_layout)

        layout.addWidget(schedule_group)

        # Options Group
        options_group = QGroupBox("Options")
        options_layout = QFormLayout()
        options_group.setLayout(options_layout)

        self.priority_combobox = self.create_combobox(["Normal", "High", "Low"])
        options_layout.addRow("Priority:", self.priority_combobox)

        self.email_notification_checkbox = QCheckBox("Send email notification")
        self.email_address_lineedit = QLineEdit()
        self.email_address_lineedit.setPlaceholderText("Enter email address")
        options_layout.addRow(self.email_notification_checkbox)
        options_layout.addRow("Email:", self.email_address_lineedit)

        self.run_on_battery = QCheckBox("Run task on battery power")
        self.run_whether_logged_on = QCheckBox("Run whether user is logged on or not")
        options_layout.addRow(self.run_on_battery)
        options_layout.addRow(self.run_whether_logged_on)

        layout.addWidget(options_group)

        # Schedule button
        schedule_btn = self.create_button('Schedule Backup', 'schedule')
        schedule_btn.clicked.connect(self.schedule_backup)
        layout.addWidget(schedule_btn)

        return scroll
    
    def create_schedule_management_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        
        # Task List Group
        task_group = QGroupBox("Scheduled Tasks")
        task_layout = QVBoxLayout()
        
        # Search
        search_layout = QHBoxLayout()
        self.task_search = QLineEdit()
        self.task_search.setPlaceholderText("Search tasks by name...")
        self.task_search.textChanged.connect(self.filter_tasks)
        search_layout.addWidget(self.task_search)
        task_layout.addLayout(search_layout)
        
        # Replace QTextEdit with QListWidget
        self.task_list = QListWidget()
        self.task_list.setSelectionMode(QListWidget.SingleSelection)
        task_layout.addWidget(self.task_list)
        
        # Task Actions
        action_layout = QHBoxLayout()
        edit_btn = self.create_button('Edit Task', 'edit')
        delete_btn = self.create_button('Delete Task', 'delete')
        disable_btn = self.create_button('Enable/Disable', 'toggle')
        run_now_btn = self.create_button('Run Now', 'run')
        
        edit_btn.clicked.connect(self.edit_selected_task)
        delete_btn.clicked.connect(self.delete_selected_task)
        disable_btn.clicked.connect(self.toggle_task_state)
        run_now_btn.clicked.connect(self.run_task_now)
        
        action_layout.addWidget(edit_btn)
        action_layout.addWidget(delete_btn)
        action_layout.addWidget(disable_btn)
        action_layout.addWidget(run_now_btn)
        
        task_layout.addLayout(action_layout)
        task_group.setLayout(task_layout)
        layout.addWidget(task_group)
        
        # Statistics Group
        stats_group = QGroupBox("Task Statistics")
        stats_layout = QGridLayout()
        stats_group.setLayout(stats_layout)

        # Create statistics labels with initial styling
        stats_style = """
            QLabel {
                font-size: 14px;
                padding: 5px;
                border-radius: 3px;
                background-color: %s;
                color: %s;
            }
        """ % (('#424242' if self.dark_mode else '#f0f0f0'),
            ('#ffffff' if self.dark_mode else '#000000'))

        self.total_tasks_label = QLabel("Total Tasks: 0")
        self.active_tasks_label = QLabel("Active Tasks: 0")
        self.completed_tasks_label = QLabel("Completed Tasks: 0")
        self.failed_tasks_label = QLabel("Failed Tasks: 0")

        # Apply styling to all statistics labels
        for label in [self.total_tasks_label, self.active_tasks_label, 
                    self.completed_tasks_label, self.failed_tasks_label]:
            label.setStyleSheet(stats_style)

        # Add labels to the grid layout
        stats_layout.addWidget(self.total_tasks_label, 0, 0)
        stats_layout.addWidget(self.active_tasks_label, 0, 1)
        stats_layout.addWidget(self.completed_tasks_label, 1, 0)
        stats_layout.addWidget(self.failed_tasks_label, 1, 1)

        # Add the statistics group to the main layout
        layout.addWidget(stats_group)
        
        # Refresh button
        refresh_btn = self.create_button('View Tasks', 'view') 
        refresh_btn.clicked.connect(self.refresh_task_list)
        layout.addWidget(refresh_btn)
        
        tab.setLayout(layout)
        return tab

    def create_restore_page(self):
        page = QWidget()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(page)
        layout = QVBoxLayout()
        page.setLayout(layout)

        title = QLabel('Database Restore')
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: bold; margin-bottom: 20px;")
        layout.addWidget(title)

        self.restore_db_host = self.create_line_edit('localhost')
        self.restore_db_port = self.create_line_edit('5432')
        self.restore_db_user = self.create_line_edit('postgres')
        self.restore_db_password = self.create_line_edit()
        self.restore_db_password.setEchoMode(QLineEdit.Password)
        self.restore_backup_dir = self.create_line_edit()

        layout.addWidget(QLabel('Host'))
        layout.addWidget(self.restore_db_host)
        layout.addWidget(QLabel('Port'))
        layout.addWidget(self.restore_db_port)
        layout.addWidget(QLabel('User'))
        layout.addWidget(self.restore_db_user)
        layout.addWidget(QLabel('Password'))
        layout.addWidget(self.restore_db_password)
        layout.addWidget(QLabel('Backup Directory'))

        restore_dir_layout = QHBoxLayout()
        restore_dir_layout.addWidget(self.restore_backup_dir)
        browse_restore_btn = QPushButton('Browse')
        browse_restore_btn.clicked.connect(self.browse_restore_dir)
        restore_dir_layout.addWidget(browse_restore_btn)
        layout.addLayout(restore_dir_layout)

        progress_layout = QHBoxLayout()
        self.restore_progress = QProgressBar()
        self.restore_progress.setTextVisible(False)
        self.restore_percentage = QLabel('0%')
        progress_layout.addWidget(self.restore_progress)
        progress_layout.addWidget(self.restore_percentage)
        layout.addLayout(progress_layout)

        self.restore_status = QLabel('')
        layout.addWidget(self.restore_status)

        btn_layout = QHBoxLayout()
        back_btn = self.create_button('Back', 'back')
        restore_btn = self.create_button('Restore', 'start_restore')
        back_btn.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        restore_btn.clicked.connect(self.perform_restore)
        btn_layout.addWidget(back_btn)
        btn_layout.addWidget(restore_btn)
        layout.addLayout(btn_layout)

        # Apply combobox style to all relevant widgets
        self.apply_combobox_style(self.restore_db_host)
        self.apply_combobox_style(self.restore_db_port)
        self.apply_combobox_style(self.restore_db_user)
        self.apply_combobox_style(self.restore_db_password)
        self.apply_combobox_style(self.restore_backup_dir)

        page.setLayout(layout)
        return scroll

    def perform_manual_backup(self):
        backup_type = 'Schema' if self.backup_type.currentText() == 'Schema-only Backup' else 'Data'
        file_extension = self.file_extension.currentText().strip('.')
        db_host = self.db_host.text()
        db_port = self.db_port.text()
        db_user = self.db_user.text()
        db_password = self.db_password.text()
        db_name = self.db_name.text()
        base_backup_dir = self.backup_dir.text()

        if not base_backup_dir:
            QMessageBox.warning(self, 'Warning', 'Please select a backup directory.')
            return

        self.backup_thread = BackupThread(backup_type, file_extension, db_name, db_host, db_port, db_user, db_password, base_backup_dir)
        self.backup_thread.progress.connect(self.update_backup_progress)
        self.backup_thread.status.connect(self.update_backup_status)
        self.backup_thread.finished.connect(self.backup_finished)
        self.backup_thread.start()

    def schedule_backup(self):
        # Get the task name from the input field
        task_name = self.task_name_input.text()
        if not task_name:
            QMessageBox.warning(self, 'Warning', 'Please enter a task name.')
            return

        # Get the description
        description = self.description_input.toPlainText()

        # Check if the task already exists
        if self.task_exists(task_name):
            QMessageBox.warning(self, 'Warning', f'A task with the name "{task_name}" already exists. Please choose a different name.')
            return

        interval = self.schedule_interval.currentText()
        time = self.schedule_time.time().toString("HH:mm")
        day = self.schedule_day.value() if interval == 'Monthly' else None
        weekday = self.schedule_weekday.currentText() if interval == 'Weekly' else None

        # Get start and end dates
        start_date = self.start_date.date().toString("yyyy/MM/dd")
        end_date = self.end_date.date().toString("yyyy/MM/dd")

        # Get all settings from manual backup tab
        backup_type = 'Schema' if self.backup_type.currentText() == 'Schema-only Backup' else 'Data'
        file_extension = self.file_extension.currentText().strip('.')
        db_host = self.db_host.text()
        db_port = self.db_port.text()
        db_user = self.db_user.text()
        db_password = self.db_password.text()
        db_name = self.db_name.text() or "all_databases"
        base_backup_dir = self.backup_dir.text()

        if not base_backup_dir:
            QMessageBox.warning(self, 'Warning', 'Please select a backup directory in the Manual Backup tab.')
            return

        # New options
        repetition = self.repetition_spinbox.value()
        priority = self.priority_combobox.currentText()
        email_notification = self.email_notification_checkbox.isChecked()
        email_address = self.email_address_lineedit.text() if email_notification else None

        # Create a backup script
        script_content = self.create_backup_script(db_host, db_port, db_user, db_password, db_name, base_backup_dir, backup_type, file_extension)
        script_path = self.save_backup_script(script_content, task_name)

        try:
            self.schedule_with_task_scheduler(script_path, interval, time, day, weekday, repetition, priority, task_name, email_notification, email_address, description, start_date, end_date)

            QMessageBox.information(self, 'Backup Scheduled', 
                f'Backup task "{task_name}" scheduled {interval} at {time}\n'
                f'Description: {description}\n'
                f'Type: {backup_type}\n'
                f'File Extension: {file_extension}\n'
                f'Database: {db_name or "All databases"}\n'
                f'Backup Directory: {base_backup_dir}\n'
                f'Repetition: Every {repetition} minutes\n'
                f'Priority: {priority}\n'
                f'Email Notification: {"Yes" if email_notification else "No"}\n'
                f'Start Date: {start_date}\n'
                f'End Date: {end_date}')

            # Send email notification
            if email_notification and email_address:
                self.send_email_notification(email_address, "Backup Task Scheduled", f'A new backup task "{task_name}" has been scheduled.')

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to schedule task: {str(e)}")
            # Clean up the created script file if scheduling fails
            if os.path.exists(script_path):
                os.remove(script_path)

    def task_exists(self, task_name):
        try:
            result = subprocess.run(['schtasks', '/query', '/tn', task_name], 
                                    capture_output=True, text=True, check=True)
            return True
        except subprocess.CalledProcessError:
            return False            

    def create_backup_script(self, db_host, db_port, db_user, db_password, db_name, backup_dir, backup_type, file_extension):
        script = textwrap.dedent(r"""
        import os
        import subprocess
        import sys
        import psycopg2
        from datetime import datetime

        def find_pg_dump():
            possible_paths = [
                r"C:\Program Files\PostgreSQL\16\bin\pg_dump.exe",
                r"C:\Program Files\PostgreSQL\15\bin\pg_dump.exe",
                r"C:\Program Files\PostgreSQL\14\bin\pg_dump.exe",
                r"C:\Program Files\PostgreSQL\13\bin\pg_dump.exe",
                r"C:\Program Files\PostgreSQL\12\bin\pg_dump.exe",
                r"D:\SETUP PROGRAMS\PostgreSQL\16\bin\pg_dump.exe",
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    return path
            return "pg_dump"  # Default to just the command name if not found

        def perform_backup(db_host, db_port, db_user, db_password, db_name, backup_dir, backup_type, file_extension):
            os.environ['PGPASSWORD'] = db_password
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            
            if db_name != "all_databases":
                backup_single_database(db_host, db_port, db_user, db_name, backup_dir, backup_type, file_extension, timestamp)
            else:
                backup_all_databases(db_host, db_port, db_user, backup_dir, backup_type, file_extension, timestamp)

        def backup_single_database(db_host, db_port, db_user, db_name, backup_dir, backup_type, file_extension, timestamp):
            if backup_type == 'Schema':
                backup_subdir = os.path.join(backup_dir, db_host, f"schema_{db_name}", timestamp)
            else:
                backup_subdir = os.path.join(backup_dir, db_host, db_name, timestamp)
            
            os.makedirs(backup_subdir, exist_ok=True)
            backup_file = os.path.join(backup_subdir, f"{db_name}.{file_extension}")
            
            pg_dump_path = find_pg_dump()
            cmd = [
                pg_dump_path,
                "-h", db_host,
                "-p", db_port,
                "-U", db_user,
                "-F", "p",  # Plain text format
                "-d", db_name
            ]
            
            if backup_type == 'Schema':
                cmd.append("-s")
            
            try:
                with open(backup_file, 'w') as f:
                    subprocess.run(cmd, stdout=f, check=True)
                print(f"Backup created successfully: {backup_file}")
            except subprocess.CalledProcessError as e:
                print(f"Error during backup of {db_name}: {e}")

        def backup_all_databases(db_host, db_port, db_user, backup_dir, backup_type, file_extension, timestamp):
            try:
                conn = psycopg2.connect(dbname='postgres', user=db_user, host=db_host, port=db_port)
                conn.autocommit = True
                cursor = conn.cursor()
                
                cursor.execute("SELECT datname FROM pg_database WHERE datistemplate = false;")
                databases = [row[0] for row in cursor.fetchall()]
                
                for db_name in databases:
                    backup_single_database(db_host, db_port, db_user, db_name, backup_dir, backup_type, file_extension, timestamp)
                
            except psycopg2.Error as e:
                print(f"Error connecting to PostgreSQL: {e}")
            finally:
                if cursor:
                    cursor.close()
                if conn:
                    conn.close()

        if __name__ == "__main__":
            if len(sys.argv) != 9:
                print("Usage: python script.py <db_host> <db_port> <db_user> <db_password> <db_name> <backup_dir> <backup_type> <file_extension>")
                sys.exit(1)
            
            perform_backup(*sys.argv[1:])
        """)
        return script

    def save_backup_script(self, script_content, task_name):
        script_dir = os.path.join(tempfile.gettempdir(), 'db_backup_scripts')
        os.makedirs(script_dir, exist_ok=True)
        safe_task_name = ''.join(c for c in task_name if c.isalnum() or c in (' ', '_')).rstrip()
        script_path = os.path.join(script_dir, f'db_backup_{safe_task_name}.py')
        with open(script_path, 'w', newline='') as f:
            f.write(script_content)
        return script_path


    def schedule_with_task_scheduler(self, script_path, interval, time, day, weekday, repetition, priority, task_name, email_notification, email_address, description, start_date, end_date):
        schedule_type = "/sc "
        if interval == 'Daily':
            schedule_type += "DAILY"
        elif interval == 'Weekly':
            schedule_type += f"WEEKLY /d {self.schedule_weekday.currentText()[:3].upper()}"
        elif interval == 'Monthly':
            schedule_type += f"MONTHLY /d {self.schedule_day.value()}"

        python_path = sys.executable

        safe_task_name = ''.join(c for c in task_name if c.isalnum() or c in (' ', '_')).rstrip()
        batch_file_name = f"run_backup_{safe_task_name}.bat"
        python_file_name = f"db_backup_{safe_task_name}.py"

        script_dir = os.path.dirname(script_path)
        batch_file_path = os.path.join(script_dir, batch_file_name)
        new_script_path = os.path.join(script_dir, python_file_name)

        os.rename(script_path, new_script_path)

        # Get the required arguments from the GUI
        db_host = self.db_host.text()
        db_port = self.db_port.text()
        db_user = self.db_user.text()
        db_password = self.db_password.text()
        db_name = self.db_name.text() or "all_databases"
        backup_dir = self.backup_dir.text()
        backup_type = 'Schema' if self.backup_type.currentText() == 'Schema-only Backup' else 'Data'
        file_extension = self.file_extension.currentText().strip('.')

        with open(batch_file_path, 'w') as batch_file:
            batch_file.write('@echo off\n')
            batch_file.write(r'set PATH=%PATH%;C:\Program Files\PostgreSQL\16\bin;C:\Program Files\PostgreSQL\15\bin;C:\Program Files\PostgreSQL\14\bin;D:\SETUP PROGRAMS\PostgreSQL\16\bin' + '\n')
            batch_file.write(f'cd /d "%~dp0"\n')
            batch_file.write(f'start /B "" "{python_path}" "{new_script_path}" "{db_host}" "{db_port}" "{db_user}" "{db_password}" "{db_name}" "{backup_dir}" "{backup_type}" "{file_extension}" >> "{new_script_path[:-3]}_log.txt" 2>&1\n')

        # Convert date format from yyyy/MM/dd to MM/dd/yyyy
        start_date = datetime.strptime(start_date, "%Y/%m/%d").strftime("%m/%d/%Y")
        end_date = datetime.strptime(end_date, "%Y/%m/%d").strftime("%m/%d/%Y")

        cmd = f'schtasks /create /tn "{task_name}" /tr "{batch_file_path}" {schedule_type} /st {time} /sd {start_date} /ed {end_date}'

        # Add repetition
        if repetition > 0:
            cmd += f" /ri {repetition}"

        # Add additional options
        if self.run_on_battery.isChecked():
            cmd += " /rl HIGHEST"

        # Run whether user is logged on or not
        if self.run_whether_logged_on.isChecked():
            cmd += " /ru \"SYSTEM\""

        try:
            result = subprocess.run(cmd, shell=True, check=True, capture_output=True, text=True)
            
            # Set priority and description using Windows API
            scheduler = win32com.client.Dispatch('Schedule.Service')
            scheduler.Connect()
            root_folder = scheduler.GetFolder('\\')
            task = root_folder.GetTask(task_name)
            task_definition = task.Definition
            
            # Set priority
            if priority == "High":
                task_definition.Settings.Priority = 7  # Highest priority
            elif priority == "Low":
                task_definition.Settings.Priority = 1  # Lowest priority
            else:
                task_definition.Settings.Priority = 4  # Normal priority
            
            # Set description
            task_definition.RegistrationInfo.Description = description
            
            root_folder.RegisterTaskDefinition(
                task_name,
                task_definition,
                6,  # Update the existing task
                None,  # No user
                None,  # No password
                3  # Run whether user is logged on or not
            )

            QMessageBox.information(self, "Task Scheduled", f"Backup task '{task_name}' has been scheduled successfully.")
            
        
            # Send email notification
            if email_notification and email_address:
                self.send_email_notification(email_address, "Backup Task Scheduled", f"A new backup task '{task_name}' has been scheduled.\nDescription: {description}")
        except subprocess.CalledProcessError as e:
            error_message = f"Failed to schedule task. Error: {e.stderr}"
            QMessageBox.critical(self, "Error", error_message)
            print(error_message)
        except Exception as e:
            error_message = f"An unexpected error occurred: {str(e)}"
            QMessageBox.critical(self, "Error", error_message)
            print(error_message)
        finally:
            # Print the full command for debugging purposes
            print(f"Executed command: {cmd}")

    def filter_tasks(self):
        search_text = self.task_search.text().lower()
        
        try:
            scheduler = win32com.client.Dispatch('Schedule.Service')
            scheduler.Connect()
            root_folder = scheduler.GetFolder('\\')
            tasks = root_folder.GetTasks(0)
            
            self.task_list.clear()
            
            for task in tasks:
                task_name = task.Name
                if not task_name.startswith("Backup_"):
                    continue
                    
                # Apply search filter
                if search_text and search_text not in task_name.lower():
                    continue
                    
                # Create list item without checkbox
                item = QListWidgetItem()
                
                # Format task information
                status = 'Running' if task.State == 4 else 'Not Running'
                last_run = task.LastRunTime.strftime("%Y-%m-%d %H:%M:%S") if task.LastRunTime else "Never"
                next_run = task.NextRunTime.strftime("%Y-%m-%d %H:%M:%S") if task.NextRunTime else "Not Scheduled"
                
                item_text = f"Task: {task_name}\nStatus: {status}\nLast Run: {last_run}\nNext Run: {next_run}"
                item.setText(item_text)
                
                # Store task name as item data
                item.setData(Qt.UserRole, task_name)
                
                self.task_list.addItem(item)
                
            self.update_statistics()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to filter tasks: {str(e)}")       

    def update_statistics(self):
        try:
            scheduler = win32com.client.Dispatch('Schedule.Service')
            scheduler.Connect()
            root_folder = scheduler.GetFolder('\\')
            tasks = root_folder.GetTasks(0)
            
            total_tasks = 0
            active_tasks = 0
            completed_tasks = 0
            failed_tasks = 0
            
            for task in tasks:
                if not task.Name.startswith("Backup_"):
                    continue
                    
                total_tasks += 1
                
                # Check task state
                if task.State == 4:  # TASK_STATE_RUNNING
                    active_tasks += 1
                
                # Check last run result
                if task.LastTaskResult == 0:  # S_OK (Success)
                    completed_tasks += 1
                elif task.LastTaskResult != 0 and task.LastRunTime:  # Error occurred
                    failed_tasks += 1
                
                # Update labels with formatted text
                self.total_tasks_label.setText(f"Total Tasks: {total_tasks}")
                self.active_tasks_label.setText(f"Active Tasks: {active_tasks}")
                self.completed_tasks_label.setText(f"Completed Tasks: {completed_tasks}")
                self.failed_tasks_label.setText(f"Failed Tasks: {failed_tasks}")
                
                # Apply styling to the statistics labels
                stats_style = """
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border-radius: 3px;
                        background-color: %s;
                        color: %s;
                    }
                """ % (('#424242' if self.dark_mode else '#f0f0f0'),
                    ('#ffffff' if self.dark_mode else '#000000'))
                
                for label in [self.total_tasks_label, self.active_tasks_label, 
                            self.completed_tasks_label, self.failed_tasks_label]:
                    label.setStyleSheet(stats_style)
                    
        except Exception as e:
            print(f"Failed to update statistics: {str(e)}")

    def get_selected_task_name(self):
        selected_items = self.task_list.selectedItems()
        if selected_items:
            return selected_items[0].data(Qt.UserRole)
        return None
    
    def get_selected_tasks(self):
        return [item.data(Qt.UserRole) for item in self.task_list.selectedItems()]

    def edit_selected_task(self):
        # Get selected item from QListWidget
        selected_items = self.task_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select a task to edit")
            return
            
        # Get task name from the selected item's text
        task_text = selected_items[0].text()
        task_name = task_text.split('\n')[0].replace('Task: ', '')
        
        try:
            scheduler = win32com.client.Dispatch('Schedule.Service')
            scheduler.Connect()
            task = scheduler.GetFolder('\\').GetTask(task_name)
            
            # Create edit dialog
            dialog = QDialog(self)
            dialog.setWindowTitle(f"Edit Task - {task_name}")
            layout = QFormLayout()
            
            # Add edit fields
            time_edit = QTimeEdit()
            time_edit.setTime(datetime.strptime(task.Definition.Triggers[0].StartBoundary, 
                            "%Y-%m-%dT%H:%M:%S").time())
            
            priority_combo = self.create_combobox(['Low', 'Normal', 'High'])
            priority_combo.setCurrentText(self.get_task_priority(task))
            
            email_checkbox = QCheckBox()
            email_checkbox.setChecked(bool(task.Definition.Actions.Count > 1))
            
            layout.addRow("Run Time:", time_edit)
            layout.addRow("Priority:", priority_combo)
            layout.addRow("Email Notification:", email_checkbox)
            
            # Add buttons
            buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            buttons.accepted.connect(dialog.accept)
            buttons.rejected.connect(dialog.reject)
            layout.addRow(buttons)
            
            dialog.setLayout(layout)
            
            if dialog.exec_() == QDialog.Accepted:
                self.update_task(task_name, time_edit.time(), priority_combo.currentText(), 
                            email_checkbox.isChecked())
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to edit task: {str(e)}")

    def update_task(self, task_name, new_time, new_priority, email_enabled):
        try:
            # Connect to Task Scheduler
            scheduler = win32com.client.Dispatch('Schedule.Service')
            scheduler.Connect()
            root_folder = scheduler.GetFolder('\\')
            task = root_folder.GetTask(task_name)
            task_def = task.Definition
            
            # Update time
            for trigger in task_def.Triggers:
                # Parse existing date from trigger
                current_datetime = datetime.strptime(trigger.StartBoundary, "%Y-%m-%dT%H:%M:%S")
                # Create new datetime with updated time
                new_datetime = current_datetime.replace(
                    hour=new_time.hour(),
                    minute=new_time.minute(),
                    second=0
                )
                # Update trigger
                trigger.StartBoundary = new_datetime.strftime("%Y-%m-%dT%H:%M:%S")
            
            # Update priority
            priority_levels = {
                "Low": 0,    # IDLE
                "Normal": 4, # NORMAL
                "High": 7    # HIGHEST
            }
            task_def.Settings.Priority = priority_levels.get(new_priority, 4)
            
            # Update email notification settings
            if email_enabled and self.email_address_lineedit.text():
                # Add or update email action if enabled
                if task_def.Actions.Count <= 1:  # Only backup action exists
                    email_action = task_def.Actions.Create(0)  # Create email action
                    email_action.From = "firetiger555@gmail.com"
                    email_action.To = self.email_address_lineedit.text()
                    email_action.Subject = f"Backup Task '{task_name}' Completed"
                    email_action.Body = "The scheduled backup task has been completed."
            elif task_def.Actions.Count > 1:  # Remove email action if disabled
                task_def.Actions.Remove(2)
                
            # Register the updated task
            root_folder.RegisterTaskDefinition(
                task_name,
                task_def,
                6,  # Update existing task
                None,  # No user
                None,  # No password
                3  # Run whether user is logged on or not
            )
            
            
            # Refresh the task list
            self.refresh_task_list()
            
            QMessageBox.information(self, "Success", f"Task '{task_name}' has been updated successfully!")
            
        except Exception as e:
            raise Exception(f"Failed to update task: {str(e)}")

    def get_task_priority(self, task):
        # Map Windows Task Scheduler priority levels to our priority options
        priority_map = {
            0: "Low",      # IDLE
            4: "Normal",   # NORMAL
            7: "High"      # HIGHEST
        }
        
        task_priority = task.Definition.Settings.Priority
        return priority_map.get(task_priority, "Normal")  # Default to Normal if priority not found

    def refresh_task_list(self):
        self.task_list.clear()
        
        try:
            scheduler = win32com.client.Dispatch('Schedule.Service')
            scheduler.Connect()
            root_folder = scheduler.GetFolder('\\')
            tasks = root_folder.GetTasks(0)
            
            filter_text = self.task_search.text().lower()
            
            for task in tasks:
                task_name = task.Name
                if filter_text in task_name.lower():
                    item = QListWidgetItem()
                    
                    status = "Enabled" if task.Enabled else "Disabled"
                    last_run = task.LastRunTime.strftime("%Y-%m-%d %H:%M:%S") if task.LastRunTime else "Never"
                    next_run = task.NextRunTime.strftime("%Y-%m-%d %H:%M:%S") if task.NextRunTime else "Not Scheduled"
                    
                    # Add separator line between task details
                    item_text = f"Task: {task_name}\nStatus: {status}\nLast Run: {last_run}\nNext Run: {next_run}\n{'-' * 50}"
                    item.setText(item_text)
                    item.setData(Qt.UserRole, task_name)
                    
                    self.task_list.addItem(item)
                    self.update_statistics()
                        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to refresh task list: {str(e)}")
            
    def run_task_now(self):
        # Get the currently selected item from the QListWidget
        selected_items = self.task_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select a task to run")
            return
            
        # Get the task name from the selected item
        selected_item = selected_items[0]
        task_name = selected_item.data(Qt.UserRole)  # Get the stored task name
        
        try:
            scheduler = win32com.client.Dispatch('Schedule.Service')
            scheduler.Connect()
            task = scheduler.GetFolder('\\').GetTask(task_name)
            task.Run(0)
            
            
            QMessageBox.information(self, "Success", 
                f'Task "{task_name}" has been started.')
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to run task: {str(e)}")

    def toggle_task_state(self):
        selected_items = self.task_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select a task to toggle state")
            return
            
        selected_item = selected_items[0]
        task_name = selected_item.data(Qt.UserRole)
        
        try:
            scheduler = win32com.client.Dispatch('Schedule.Service')
            scheduler.Connect()
            task = scheduler.GetFolder('\\').GetTask(task_name)
            
            # Get current state
            enabled = task.Enabled
            
            # Toggle state
            task_definition = task.Definition
            task_definition.Settings.Enabled = not enabled
            
            # Update task
            scheduler.GetFolder('\\').RegisterTaskDefinition(
                task_name,
                task_definition,
                6,  # Update existing task
                None,  # No user
                None,  # No password
                3  # Run whether user is logged on or not
            )
            
            # Log state change
            new_state = "Enabled" if not enabled else "Disabled"
                        
            
            # Refresh the task list
            self.refresh_task_list()
            
            QMessageBox.information(self, "Success", 
                f'Task "{task_name}" has been {new_state.lower()}.')
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to toggle task state: {str(e)}")

    def delete_selected_task(self):
        selected_items = self.task_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select a task to delete")
            return
            
        selected_item = selected_items[0]
        task_name = selected_item.data(Qt.UserRole)
        
        # Confirm deletion
        reply = QMessageBox.question(self, 'Confirm Delete', 
                                f'Are you sure you want to delete task "{task_name}"?',
                                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            try:
                # Delete the task using schtasks
                subprocess.run(['schtasks', '/delete', '/tn', task_name, '/f'], 
                            check=True, capture_output=True, text=True)
                
                # Delete associated script and batch files
                script_dir = os.path.join(tempfile.gettempdir(), 'db_backup_scripts')
                safe_task_name = ''.join(c for c in task_name if c.isalnum() or c in (' ', '_')).rstrip()
                
                # File paths
                python_script = os.path.join(script_dir, f'db_backup_{safe_task_name}.py')
                batch_file = os.path.join(script_dir, f'run_backup_{safe_task_name}.bat')
                log_file = os.path.join(script_dir, f'db_backup_{safe_task_name}_log.txt')
                
                # Remove files if they exist
                for file_path in [python_script, batch_file, log_file]:
                    if os.path.exists(file_path):
                        try:
                            os.remove(file_path)
                            print(f"Deleted file: {file_path}")
                        except Exception as e:
                            print(f"Error deleting file {file_path}: {str(e)}")
                
                
                # Refresh the task list
                self.refresh_task_list()
                self.update_statistics()
                
                QMessageBox.information(self, "Success", 
                    f'Task "{task_name}" and associated files have been deleted.')
                
            except subprocess.CalledProcessError as e:
                QMessageBox.critical(self, "Error", f"Failed to delete task: {e.stderr}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")


    def send_email_notification(self, email_address, subject, message):
        # Email configuration
        sender_email = "firetiger555@gmail.com"  # Replace with your Gmail address
        sender_password = "jqfs zlyt rprs ehtb"   # Replace with your Gmail app password
        smtp_server = "smtp.gmail.com"
        smtp_port = 587

        # Create detailed schedule message
        schedule_details = f"""
        Backup Schedule Details:
        -----------------------
        Schedule Type: {self.schedule_interval.currentText()}
        Time: {self.schedule_time.time().toString("HH:mm")}
        Start Date: {self.start_date.date().toString("yyyy-MM-dd")}
        End Date: {self.end_date.date().toString("yyyy-MM-dd")}
        Repetition: Every {self.repetition_spinbox.value()} minutes
        Priority: {self.priority_combobox.currentText()}
        
        Database Configuration:
        ----------------------
        Host: {self.db_host.text()}
        Port: {self.db_port.text()}
        Database: {self.db_name.text() or "All Databases"}
        Backup Type: {self.backup_type.currentText()}
        Backup Directory: {self.backup_dir.text()}
        """

        # Create the email message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = email_address
        msg['Subject'] = subject
        msg.attach(MIMEText(message + "\n\n" + schedule_details, 'plain'))

        # Send email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)

        # Add body to email
        msg.attach(MIMEText(message, 'plain'))

        try:
            # Create SMTP session
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()  # Enable TLS
                server.login(sender_email, sender_password)
                
                # Send email
                server.send_message(msg)
            
            print(f"Email notification sent successfully to {email_address}")
        except Exception as e:
            print(f"Failed to send email notification: {str(e)}")

    def update_backup_progress(self, value):
        self.backup_progress.setValue(value)
        self.backup_percentage.setText(f'{value}%')

    def update_backup_status(self, status):
        self.backup_status.setText(status)

    def backup_finished(self, success, message):
        if success:
            QMessageBox.information(self, 'Success', message)
        else:
            QMessageBox.critical(self, 'Error', message)
        self.backup_progress.setValue(0)
        self.backup_percentage.setText('0%')

    def perform_restore(self):
        db_host = self.restore_db_host.text()
        db_port = self.restore_db_port.text()
        db_user = self.restore_db_user.text()
        db_password = self.restore_db_password.text()
        backup_dir = self.restore_backup_dir.text()

        if not backup_dir:
            QMessageBox.warning(self, 'Warning', 'Please select a restore directory.')
            return

        self.restore_thread = RestoreThread(db_host, db_port, db_user, db_password, backup_dir)
        self.restore_thread.progress.connect(self.update_restore_progress)
        self.restore_thread.status.connect(self.update_restore_status)
        self.restore_thread.finished.connect(self.restore_finished)
        self.restore_thread.start()

    def update_restore_progress(self, value):
        self.restore_progress.setValue(value)
        self.restore_percentage.setText(f'{value}%')

    def update_restore_status(self, status):
        self.restore_status.setText(status)

    def restore_finished(self, success, message):
        if success:
            QMessageBox.information(self, 'Success', message)
        else:
            QMessageBox.critical(self, 'Error', message)
        self.restore_progress.setValue(0)
        self.restore_percentage.setText('0%')

    def toggle_dark_mode(self):
        self.dark_mode = not self.dark_mode
        self.setStyleSheet(self.get_stylesheet())
        self.update_dark_mode_button()
        self.update_icons()
        
        # Reapply styles to all widgets
        for page in [self.main_page, self.backup_page, self.restore_page]:
            for child in page.findChildren((QComboBox, QLineEdit, QTimeEdit, QSpinBox, QDateEdit)):
                self.apply_combobox_style(child)

    def update_combobox_style(self, combobox):
        items = [combobox.itemText(i) for i in range(combobox.count())]
        index = combobox.currentIndex()
        new_combobox = self.create_combobox(items)
        new_combobox.setCurrentIndex(index)
        
        parent_layout = combobox.parent().layout()
        if parent_layout:
            for i in range(parent_layout.count()):
                if parent_layout.itemAt(i).widget() == combobox:
                    parent_layout.replaceWidget(combobox, new_combobox)
                    combobox.deleteLater()
                    break

    def update_dark_mode_button(self):
        icon_name = 'light_mode' if self.dark_mode else 'dark_mode'
        icon_path = f'icons/{icon_name}.png'
        if os.path.exists(icon_path):
            self.dark_mode_button.setIcon(QIcon(icon_path))
        else:
            print(f"Dark mode icon not found: {icon_path}")

    def update_icons(self):
        for page in [self.main_page, self.backup_page, self.restore_page]:
            for child in page.findChildren(QPushButton):
                if child.objectName().endswith('_btn'):
                    icon_name = child.objectName().replace('_btn', '')
                    self.update_button_icon(child, icon_name)

    def get_stylesheet(self):
        base_style = """
        QScrollArea {
            border: none;
        }
        QScrollBar:vertical {
            border: none;
            background: #f0f0f0;
            width: 10px;
            margin: 0px 0px 0px 0px;
        }
        QScrollBar::handle:vertical {
            background: #c0c0c0;
            min-height: 20px;
            border-radius: 5px;
        }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            border: none;
            background: none;
        }
        QComboBox QAbstractItemView {
            border: 1px solid #ccc;
            selection-background-color: #f0f0f0;
            max-height: 200px;
        }
        """
        
        if self.dark_mode:
            return base_style + """
            QWidget {
                background-color: #2b2b2b;
                color: #ffffff;
                font-family: Arial, sans-serif;
            }
            """ + self.dark_mode_styles()
        else:
            return base_style + """
            QWidget {
                background-color: #f0f0f0;
                font-family: Arial, sans-serif;
            }
            """ + self.light_mode_styles()

    def dark_mode_styles(self):
        return """
        QPushButton {
            background-color: #0d47a1;
            color: white;
            border: none;
            padding: 10px 20px;
            text-align: center;
            text-decoration: none;
            font-size: 16px;
            margin: 4px 2px;
            border-radius: 4px;
        }
        QPushButton:hover {
            background-color: #1565c0;
        }
        QLineEdit, QComboBox, QSpinBox, QTimeEdit {
            background-color: #424242;
            color: #ffffff;
            padding: 8px;
            margin: 4px 0;
            border: 1px solid #555;
            border-radius: 4px;
        }
        QLabel {
            font-size: 14px;
            margin-top: 8px;
        }
        QProgressBar {
            border: 2px solid #555;
            border-radius: 5px;
            text-align: center;
        }
        QProgressBar::chunk {
            background-color: #0d47a1;
            width: 10px;
            margin: 0.5px;
        }
        QTabWidget::pane {
            border: 1px solid #555;
            background-color: #2b2b2b;
        }
        QTabBar::tab {
            background-color: #1e1e1e;
            color: #ffffff;
            padding: 8px 12px;
            margin-right: 2px;
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
        }
        QTabBar::tab:selected {
            background-color: #0d47a1;
        }
        QGroupBox {
            border: 1px solid #555;
            border-radius: 6px;
            margin-top: 6px;
            padding-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 7px;
            padding: 0 3px 0 3px;
            color: #ffffff;
        }
        QCheckBox, QRadioButton {
            color: #ffffff;
        }
        QTextEdit {
            background-color: #424242;
            color: #ffffff;
            border: 1px solid #555;
            border-radius: 4px;
        }
        
        QSpinBox {
        background-color: #424242;
        color: #ffffff;
        padding: 8px;
        margin: 4px 0;
        border: 1px solid #555;
        border-radius: 4px;
        }
    
        QSpinBox::up-button, QSpinBox::down-button {
            background-color: #555;
            border: none;
            border-radius: 2px;
            margin: 1px;
        }
        
        QSpinBox::up-button:hover, QSpinBox::down-button:hover {
            background-color: #666;
        }
        
        QSpinBox::up-arrow {
            image: url(icons/up_arrow_dark.png);
            width: 8px;
            height: 8px;
        }
        
        QSpinBox::down-arrow {
            image: url(icons/down_arrow_dark.png);
            width: 8px;
            height: 8px;
        }
        """
        

    def light_mode_styles(self):
        return """
        QPushButton {
            background-color: #4CAF50;
            color: white;
            border: none;
            padding: 10px 20px;
            text-align: center;
            text-decoration: none;
            font-size: 16px;
            margin: 4px 2px;
            border-radius: 4px;
        }
        QPushButton:hover {
            background-color: #45a049;
        }
        QLineEdit, QComboBox, QSpinBox, QTimeEdit {
            background-color: #ffffff;
            color: #000000;
            padding: 8px;
            margin: 4px 0;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
        QLabel {
            font-size: 14px;
            margin-top: 8px;
            color: #000000;
        }
        QProgressBar {
            border: 2px solid #ddd;
            border-radius: 5px;
            text-align: center;
        }
        QProgressBar::chunk {
            background-color: #4CAF50;
            width: 10px;
            margin: 0.5px;
        }
        QTabWidget::pane {
            border: 1px solid #ddd;
            background-color: #f0f0f0;
        }
        QTabBar::tab {
            background-color: #e0e0e0;
            color: #000000;
            padding: 8px 12px;
            margin-right: 2px;
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
        }
        QTabBar::tab:selected {
            background-color: #4CAF50;
            color: #ffffff;
        }
        QGroupBox {
            border: 1px solid #ccc;
            border-radius: 6px;
            margin-top: 6px;
            padding-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 7px;
            padding: 0 3px 0 3px;
            color: #000000;
        }
        QCheckBox, QRadioButton {
            color: #000000;
        }
        QTextEdit {
            background-color: #ffffff;
            color: #000000;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
        
        QSpinBox {
        background-color: #ffffff;
        color: #000000;
        padding: 8px;
        margin: 4px 0;
        border: 1px solid #ddd;
        border-radius: 4px;
        }
        
        QSpinBox::up-button, QSpinBox::down-button {
            background-color: #f0f0f0;
            border: none;
            border-radius: 2px;
            margin: 1px;
        }
        
        QSpinBox::up-button:hover, QSpinBox::down-button:hover {
            background-color: #e0e0e0;
        }
        
        QSpinBox::up-arrow {
            image: url(icons/up_arrow_light.png);
            width: 8px;
            height: 8px;
        }
        
        QSpinBox::down-arrow {
            image: url(icons/down_arrow_light.png);
            width: 8px;
            height: 8px;
        }
        """

    def create_button(self, text, icon_name):
        button = QPushButton(text)
        button.setObjectName(f"{icon_name}_btn")
        self.update_button_icon(button, icon_name)
        return button

    def update_button_icon(self, button, icon_name):
        icon_suffix = '_dark' if self.dark_mode else '_light'
        icon_path = f'icons/{icon_name}{icon_suffix}.png'
        if os.path.exists(icon_path):
            button.setIcon(QIcon(icon_path))
        else:
            print(f"Icon not found: {icon_path}")

    def create_combobox(self, items):
        combobox = NoScrollComboBox()
        combobox.addItems(items)
        self.apply_combobox_style(combobox)
        return combobox

    def apply_combobox_style(self, widget):
        arrow_icon = "dropdown_arrow_dark.png" if self.dark_mode else "dropdown_arrow_light.png"
        
        if isinstance(widget, (QComboBox, QSpinBox, QTimeEdit)):
            widget.setStyleSheet(f"""
                QComboBox, QSpinBox, QTimeEdit {{
                    border: 1px solid {'#555' if self.dark_mode else '#ccc'};
                    border-radius: 3px;
                    padding: 5px;
                    min-width: 6em;
                    color: {'#ffffff' if self.dark_mode else '#000000'};
                    background-color: {'#424242' if self.dark_mode else '#ffffff'};
                }}
                QComboBox::drop-down, QSpinBox::drop-down, QTimeEdit::drop-down {{
                    subcontrol-origin: padding;
                    subcontrol-position: top right;
                    width: 20px;
                    border-left-width: 1px;
                    border-left-color: {'#555' if self.dark_mode else '#ccc'};
                    border-left-style: solid;
                    border-top-right-radius: 3px;
                    border-bottom-right-radius: 3px;
                }}
                QComboBox::down-arrow, QSpinBox::down-arrow, QTimeEdit::down-arrow {{
                    image: url(icons/{arrow_icon});
                    width: 12px;
                    height: 12px;
                }}
                QComboBox QAbstractItemView {{
                    border: 1px solid {'#555' if self.dark_mode else '#ccc'};
                    selection-background-color: {'#0d47a1' if self.dark_mode else '#e0e0e0'};
                    background-color: {'#424242' if self.dark_mode else '#ffffff'};
                    color: {'#ffffff' if self.dark_mode else '#000000'};
                }}
            """)
        elif isinstance(widget, QDateEdit):
            widget.setStyleSheet(f"""
                QDateEdit {{
                    border: 1px solid {'#555' if self.dark_mode else '#ccc'};
                    border-radius: 3px;
                    padding: 5px;
                    min-width: 6em;
                    color: {'#ffffff' if self.dark_mode else '#000000'};
                    background-color: {'#424242' if self.dark_mode else '#ffffff'};
                }}
                QDateEdit::drop-down {{
                    subcontrol-origin: padding;
                    subcontrol-position: top right;
                    width: 20px;
                    border-left-width: 1px;
                    border-left-color: {'#555' if self.dark_mode else '#ccc'};
                    border-left-style: solid;
                    border-top-right-radius: 3px;
                    border-bottom-right-radius: 3px;
                }}
                QDateEdit::down-arrow {{
                    image: url(icons/{arrow_icon});
                    width: 12px;
                    height: 12px;
                }}
                QCalendarWidget QToolButton {{
                    height: 30px;
                    width: 90px;
                    color: {'#ffffff' if self.dark_mode else '#000000'};
                    background-color: {'#424242' if self.dark_mode else '#f0f0f0'};
                    font-size: 14px;
                    icon-size: 16px, 16px;
                }}
                QCalendarWidget QMenu {{
                    color: {'#ffffff' if self.dark_mode else '#000000'};
                    background-color: {'#424242' if self.dark_mode else '#ffffff'};
                }}
                QCalendarWidget QSpinBox {{
                    color: {'#ffffff' if self.dark_mode else '#000000'};
                    background-color: {'#424242' if self.dark_mode else '#ffffff'};
                    selection-background-color: {'#0d47a1' if self.dark_mode else '#e0e0e0'};
                    selection-color: {'#ffffff' if self.dark_mode else '#000000'};
                }}
                QCalendarWidget QAbstractItemView:enabled {{
                    color: {'#ffffff' if self.dark_mode else '#000000'};
                    background-color: {'#2b2b2b' if self.dark_mode else '#ffffff'};
                    selection-background-color: {'#0d47a1' if self.dark_mode else '#e0e0e0'};
                    selection-color: {'#ffffff' if self.dark_mode else '#000000'};
                }}
                QCalendarWidget QAbstractItemView:disabled {{
                    color: {'#666666' if self.dark_mode else '#999999'};
                }}
            """)
        elif isinstance(widget, QLineEdit):
            widget.setStyleSheet(f"""
                QLineEdit {{
                    border: 1px solid {'#555' if self.dark_mode else '#ccc'};
                    border-radius: 3px;
                    padding: 5px;
                    background-color: {'#424242' if self.dark_mode else '#ffffff'};
                    color: {'#ffffff' if self.dark_mode else '#000000'};
                }}
            """)
        else:
            print(f"Unsupported widget type for styling: {type(widget)}")

    def create_line_edit(self, placeholder_text=''):
        line_edit = QLineEdit()
        line_edit.setPlaceholderText(placeholder_text)
        return line_edit

    def browse_backup_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Select Backup Directory")
        if dir_path:
            self.backup_dir.setText(dir_path)

    def browse_restore_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Select Restore Directory")
        if dir_path:
            self.restore_backup_dir.setText(dir_path)

    def update_schedule_options(self):
        interval = self.schedule_interval.currentText()
        
        if interval == 'Monthly':
            self.schedule_day.setEnabled(True)
            self.schedule_weekday.setEnabled(False)
        elif interval == 'Weekly':
            self.schedule_day.setEnabled(False)
            self.schedule_weekday.setEnabled(True)
        else:  # Daily
            self.schedule_day.setEnabled(False)
            self.schedule_weekday.setEnabled(False)

def main():
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("icons/app_icon.png"))
    ex = ModernBackupRestoreGUI()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()

