from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QTableWidget, QTableWidgetItem,
    QSpacerItem, QSizePolicy, QMenu, QSplitter, QStatusBar, QWidget,
    QPushButton, QGridLayout, QFrame, QVBoxLayout, QHBoxLayout, QLabel, QMessageBox, QListWidget, QCheckBox, QLineEdit
)
from PySide6.QtGui import QFont, QColor, QIcon, QCursor, QKeySequence, QShortcut
from PySide6.QtCore import Signal, QSettings, Qt, QTimer, Slot
from PySide6 import QtCore, QtWidgets, QtGui

from pathlib import Path
import pandas as pd
import webbrowser
import logging
import sys
import os

from gui.log_window import LogWindow
from XTB_converter import CashOperationXLSXReader
from gui.update_checker import UpdateChecker

logging.basicConfig(level=logging.NOTSET, filename="log.log", filemode="w", format="%(asctime)s - %(lineno)d - %(levelname)s - %(message)s")
settings = QSettings("PP", "Portfolio Performance")


VERSION = "0.9.0"


class MyMainWindow(QMainWindow):
    def __init__(self, argv_path=None):
        super().__init__()
        self.base_path = self._get_base_path()

        self.file_paths = []

        self.settings = settings
        self.dark_mode_enabled = self.settings.value("DarkMode", False, type=bool)

        self.log_window = LogWindow(self)

        self._configure_window()
        self._init_status_bar()
        self._init_ui()

        self._connect_option_logic()
        self._load_export_path()
        self.version_checker()

        self.setup_file_list_actions()

        logging.debug("Main application initialized.")

    # Setup
    def _get_base_path(self) -> Path:
        if getattr(sys, 'frozen', False):
            return Path(sys.executable).parent
        else:
            import __main__
            return Path(__main__.__file__).resolve().parent

    def resource(self, *parts):
        #print(f"Accessing resource: {self.base_path.joinpath(*parts)}")
        return str(self.base_path.joinpath(*parts))

    def _configure_window(self):
        self.setWindowIcon(QIcon(self.resource("gui", "Stylesheets", "GML.ico")))
        self.setWindowTitle(f"XTB to Portfolio Performance v{VERSION}")
        self.setMinimumSize(800, 500)
        self.setAcceptDrops(True)

    def version_checker(self):  # Check if the current version is up to date.
        result = UpdateChecker("RybarskiDominik/XTB-TO-PORTFOLIO-PERFORMANCE").check_app_update_status(VERSION)
        if result is True:
            print("Dostępna jest aktualizacja")
            self.update_status_bar("🚀 Dostępna jest nowa aktualizacja.", 10000, "red")
        logging.info(f"Running version {VERSION}")

    # Init UI
    def _init_ui(self):
        """Initialize main user interface layout."""

        # ===== CENTRAL WIDGET =====
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        # ===== MAIN CONTENT (LEFT / RIGHT) =====
        content_layout = QHBoxLayout()
        content_layout.setSpacing(5)

        # ==========================================================
        # LEFT PANEL — FILE INPUT
        # ==========================================================
        left_panel_layout = QVBoxLayout()

        self.drop_info_label = QLabel(
            'Drag & drop the .xlsx file exported from XTB\n'
            '(Report: "Cash Operations")'
        )
        self.drop_info_label.setAlignment(Qt.AlignCenter)
        self.drop_info_label.setStyleSheet(
            "font-size: 16px; color: gray; "
            "border: 2px dashed #cccccc; padding: 20px;"
        )

        self.file_list_widget = QListWidget()

        left_panel_layout.addWidget(self.drop_info_label)
        left_panel_layout.addWidget(self.file_list_widget)

        content_layout.addLayout(left_panel_layout, 1)


        # ==========================================================
        # RIGHT PANEL — EXPORT SETTINGS
        # ==========================================================
        right_panel_layout = QVBoxLayout()

        settings_frame = QFrame()
        settings_frame.setFrameShape(QFrame.StyledPanel)

        settings_layout = QGridLayout(settings_frame)
        settings_layout.setSpacing(6)

        # ===== EXPORT DIRECTORY =====
        export_dir_label = QLabel("Export Directory:")
        export_dir_label.setStyleSheet("font-weight: bold;")

        self.export_path_input = QLineEdit()
        self.export_path_input.setPlaceholderText("Select destination folder...")

        self.browse_export_button = QPushButton("Browse")
        self.browse_export_button.setFixedWidth(100)
        self.browse_export_button.clicked.connect(self._browse_export_directory)

        settings_layout.addWidget(export_dir_label, 0, 0, 1, 3)
        settings_layout.addWidget(self.export_path_input, 1, 0, 1, 4)
        settings_layout.addWidget(self.browse_export_button, 1, 5)

        settings_layout.setColumnStretch(0, 1)
        settings_layout.setColumnStretch(1, 1)
        settings_layout.setColumnStretch(2, 0)

        # ===== DEFAULT EXPORT OPTION =====
        default_section_label = QLabel("Default processing method")
        default_section_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        settings_layout.addWidget(default_section_label, 2, 0, 1, 3)

        self.default_export_checkbox = QCheckBox(
            "Perform default export (no additional processing)"
        )
        self.default_export_checkbox.setChecked(True)

        settings_layout.addWidget(self.default_export_checkbox, 3, 0, 1, 3)

        # ===== ADVANCED OPTIONS =====
        advanced_section_label = QLabel("Alternative Advanced Processing Options")
        advanced_section_label.setStyleSheet(
            "font-weight: bold; font-size: 14px; margin-top: 10px;"
        )
        settings_layout.addWidget(advanced_section_label, 4, 0, 1, 3)

        self.include_open_positions_checkbox = QCheckBox("Include open positions")
        self.include_closed_positions_checkbox = QCheckBox("Include closed positions")
        self.simplified_deposit_checkbox = QCheckBox("Use simplified deposit format")

        settings_layout.addWidget(self.include_open_positions_checkbox, 5, 0, 1, 3)
        settings_layout.addWidget(self.include_closed_positions_checkbox, 6, 0, 1, 3)
        settings_layout.addWidget(self.simplified_deposit_checkbox, 7, 0, 1, 3)

        settings_layout.setRowStretch(8, 1)

        right_panel_layout.addWidget(settings_frame)

        self.export_button = QPushButton("Export to CSV")
        self.export_button.setFixedWidth(100)
        self.export_button.clicked.connect(self.process_files)
        self.export_button.setStyleSheet("padding: 8px; font-weight: bold;")

        settings_layout.addWidget(self.export_button, 9, 5, 1, 1)

        content_layout.addLayout(right_panel_layout, 4)
        main_layout.addLayout(content_layout)

    def _connect_option_logic(self):
        """Connect export option logic."""

        # Default
        self.default_export_checkbox.stateChanged.connect(
            self._handle_default_checkbox
        )

        # Advanced — uncheck default when any advanced option is selected
        for checkbox in [
            self.include_open_positions_checkbox,
            self.include_closed_positions_checkbox,
            self.simplified_deposit_checkbox,
        ]:
            checkbox.stateChanged.connect(
                lambda: self.default_export_checkbox.setChecked(False)
            )

    def _handle_default_checkbox(self):
        if self.default_export_checkbox.isChecked():
            self.include_open_positions_checkbox.setChecked(False)
            self.include_closed_positions_checkbox.setChecked(False)
            self.simplified_deposit_checkbox.setChecked(False)
            self.default_export_checkbox.setChecked(True)

    def _browse_export_directory(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Export Directory"
        )

        if folder:
            self.export_path_input.setText(folder)
            self.settings.setValue("ExportPath", folder)

    def _load_export_path(self):
        saved_path = self.settings.value("ExportPath", "", type=str)

        if saved_path and Path(saved_path).exists():
            self.export_path_input.setText(saved_path)
        else:
            # Jeśli ścieżka nie istnieje → wyczyść pole
            self.export_path_input.clear()

    # Main functions
    def process_files(self):
        account_currency = ""
        data = pd.DataFrame()
        open_positions, closed_positions, simplified_deposit = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        if not self.file_paths:
            QMessageBox.warning(self, "No file", "Please add at least one .xlsx file to process.")
            return

        export_path = self.export_path_input.text().strip()
        if not export_path:
            QMessageBox.warning(self, "No export directory", "Please select an export directory.")
            return
        
        if not self.default_export_checkbox.isChecked() and not (self.include_open_positions_checkbox.isChecked() or self.include_closed_positions_checkbox.isChecked() or self.simplified_deposit_checkbox.isChecked()):
            QMessageBox.warning(self, "No export options", "Please select at least one export option.")
            return

        for file_path in self.file_paths:
            ac = CashOperationXLSXReader(file_path, 3).read_header()
            account_currency = ac.get("Currency", "")
            
            if self.default_export_checkbox.isChecked():
                converter = CashOperationXLSXReader(file_path, 3)
                data = converter.export_default_cash_operations()
            else:
                if self.include_open_positions_checkbox.isChecked():
                    cash_open_operations = CashOperationXLSXReader(file_path, sheet_index=1)
                    open_positions = cash_open_operations.export_open_operations()
                if self.include_closed_positions_checkbox.isChecked():
                    cash_closed_operations = CashOperationXLSXReader(file_path, sheet_index=0)
                    closed_positions = cash_closed_operations.export_closed_operations()
                if self.simplified_deposit_checkbox.isChecked():
                    cash_deposit = CashOperationXLSXReader(file_path, sheet_index=3)
                    simplified_deposit = cash_deposit.export_simplified_deposit_of_operation()
                
                data = pd.concat([open_positions, closed_positions, simplified_deposit], ignore_index=True)

            data.to_csv(Path(export_path) / f"{Path(file_path).stem}_XTB_{account_currency}.csv", index=False)

    # Status bar.
    def _init_status_bar(self):
        self.statusbar = QStatusBar()
        self.setStatusBar(self.statusbar)

        self.default_color = "white" if self.dark_mode_enabled else "black"

        container = QWidget()
        container.setMaximumHeight(20)

        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.log_button = QPushButton()
        self.log_button.setFixedSize(20, 20)

        icon_path = (
            self.resource("gui", "Stylesheets", "images_dark-light", "Data-light.svg")
            if self.dark_mode_enabled
            else self.resource("gui", "Stylesheets", "images_dark-light", "Data-dark.svg")
        )

        self.log_button.setIcon(QIcon(icon_path))
        self.log_button.setIconSize(QtCore.QSize(20, 20))
        self.log_button.setToolTip("Application logs")
        self.log_button.clicked.connect(self.log_window.show)

        self.status_label = QLabel()
        self.status_label.setStyleSheet(f"color: {self.default_color}; font-size: 14px;")

        layout.addWidget(self.log_button)
        layout.addStretch()
        layout.addWidget(self.status_label)
        layout.addStretch()

        self.statusbar.addWidget(container, 1)

    @Slot(str)
    def update_status_bar(self, message: str, duration: int = 10000, color: str = None):
        color = color or self.default_color

        self.status_label.setStyleSheet(f"color: {color}; font-size: 14px;")
        self.status_label.setText(message)

        if duration:
            QTimer.singleShot(duration, self.clear_status_bar)

    def clear_status_bar(self):
        self.status_label.setStyleSheet(f"color: {self.default_color}; font-size: 14px;")
        self.status_label.clear()

    # External Links
    def open_url(self, url: str):
        try:
            webbrowser.open(url)
        except webbrowser.Error:
            edge_path = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
            webbrowser.register("edge", None, webbrowser.BackgroundBrowser(edge_path))
            webbrowser.get("edge").open(url)

    def open_donation_page(self):
        self.open_url("https://www.paypal.com/donate/?hosted_button_id=DVJJ5QVHCN2X6")

    def open_github(self):
        self.open_url("https://github.com/RybarskiDominik/XTB-TO-PORTFOLIO-PERFORMANCE")

    def store_file_(self, file_path):
        """Dodaje plik do listy i wyświetla jego nazwę w QListWidget."""
        if file_path not in self.file_paths:
            self.file_paths.append(file_path)
            file_name = file_path.split("/")[-1]  # tylko nazwa pliku
            self.file_list_widget.addItem(file_name)
            print(f"File stored: {file_path}")
        else:
            print(f"File already in list: {file_path}")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().endswith(".xlsx"):
                    event.acceptProposedAction()
                    return     

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            print(f'Dropped GML file: {file_path}')
            self.store_file_(file_path)


    def setup_file_list_actions(self):
        """Dodaje menu kontekstowe i klawisz Delete do QListWidget."""
        
        # Menu kontekstowe
        self.file_list_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.file_list_widget.customContextMenuRequested.connect(self.show_file_context_menu)
        
        # Obsługa klawisza Delete
        self.file_list_widget.keyPressEvent = self.file_list_key_press

    def show_file_context_menu(self, pos):
        menu = QMenu()
        remove_action = menu.addAction("Usuń")
        action = menu.exec(self.file_list_widget.mapToGlobal(pos))
        
        if action == remove_action:
            self.remove_selected_files()

    def file_list_key_press(self, event):
        if event.key() == Qt.Key_Delete:
            self.remove_selected_files()
        else:
            # domyślna obsługa innych klawiszy
            QListWidget.keyPressEvent(self.file_list_widget, event)

    def remove_selected_files(self):
        """Usuwa zaznaczone pliki z QListWidget i z self.file_paths."""
        selected_items = self.file_list_widget.selectedItems()
        for item in selected_items:
            row = self.file_list_widget.row(item)
            self.file_list_widget.takeItem(row)
            # Usuń pełną ścieżkę odpowiadającą nazwie pliku
            # Szukamy w self.file_paths pliku z końcówką matching item.text()
            for path in self.file_paths:
                if path.endswith(item.text()):
                    self.file_paths.remove(path)
                    break
        print(f"Remaining files: {self.file_paths}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    dark_mode = settings.value("DarkMode", False, type=bool)

    try:
        if dark_mode:
            stylesheet_path = Path("gui/Stylesheets/Darkmode.qss")
            app.setStyleSheet(stylesheet_path.read_text())
        else:
            app.setStyleSheet("""
                QGraphicsView {
                    border: none;
                    background: transparent;
                }
            """)
    except Exception as e:
        logging.exception(e)

    window = MyMainWindow()
    window.show()

    sys.exit(app.exec())