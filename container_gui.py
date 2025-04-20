import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QTextEdit, QPushButton, 
                           QFileDialog, QMessageBox, QTabWidget,
                           QFrame, QRadioButton, QButtonGroup)
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
from file_con import create_summary

class DropArea(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.StyledPanel | QFrame.Sunken)
        self.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 2px dashed #dee2e6;
                border-radius: 5px;
                padding: 20px;
                min-height: 100px;
            }
        """)
        
        layout = QVBoxLayout(self)
        self.label = QLabel("여기에 ASC 파일을 드래그 앤 드롭하세요")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)
        
        self.file_path = None
        
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            
    def dropEvent(self, event: QDropEvent):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        if files:
            self.file_path = files[0]
            self.label.setText(os.path.basename(self.file_path))
            self.label.setStyleSheet("color: #28a745;")  # Green color for success

class ContainerTab(QWidget):
    def __init__(self, title, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)
        
        label = QLabel(f"{title} (한 줄에 하나의 컨테이너 번호):")
        self.text_edit = QTextEdit()
        self.text_edit.setMinimumHeight(200)
        self.text_edit.setStyleSheet("""
            QTextEdit {
                background-color: white;
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        
        layout.addWidget(label)
        layout.addWidget(self.text_edit)
        layout.addStretch()
        
    def get_container_list(self):
        return [line.strip() for line in self.text_edit.toPlainText().split('\n') if line.strip()]

class ContainerAnalyzerGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('Container Analyzer')
        self.setGeometry(100, 100, 1000, 800)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        
        # Operation type selection with radio buttons
        op_layout = QHBoxLayout()
        op_label = QLabel('Operation Type:')
        op_layout.addWidget(op_label)
        
        self.op_group = QButtonGroup(self)
        self.discharge_radio = QRadioButton('Discharge')
        self.load_radio = QRadioButton('Load')
        self.discharge_radio.setChecked(True)  # Default selection
        
        self.op_group.addButton(self.discharge_radio)
        self.op_group.addButton(self.load_radio)
        
        op_layout.addWidget(self.discharge_radio)
        op_layout.addWidget(self.load_radio)
        op_layout.addStretch()
        
        main_layout.addLayout(op_layout)
        
        # ASC file drop area
        main_layout.addWidget(QLabel('ASC File:'))
        self.drop_area = DropArea()
        main_layout.addWidget(self.drop_area)
        
        # Tab widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #dee2e6;
                border-radius: 4px;
                background: white;
            }
            QTabBar::tab {
                padding: 8px 16px;
                margin: 2px;
                background: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 4px;
            }
            QTabBar::tab:selected {
                background: white;
                border-bottom: none;
            }
        """)
        
        # Create tabs
        self.tpf_tab = ContainerTab('TPF Containers')
        self.truck_tab = ContainerTab('Truck Containers')
        self.local_tab = ContainerTab('Local')
        self.same_ts_tab = ContainerTab('Same TS')
        self.external_ts_tab = ContainerTab('External TS')
        
        self.tab_widget.addTab(self.tpf_tab, 'TPF Containers')
        self.tab_widget.addTab(self.truck_tab, 'Truck Containers')
        self.tab_widget.addTab(self.local_tab, 'Local')
        self.tab_widget.addTab(self.same_ts_tab, 'Same TS')
        self.tab_widget.addTab(self.external_ts_tab, 'External TS')
        
        main_layout.addWidget(self.tab_widget)
        
        # Process button
        self.process_btn = QPushButton('Create Summary')
        self.process_btn.setStyleSheet("""
            QPushButton {
                background-color: #007bff;
                color: white;
                padding: 10px 20px;
                font-size: 14px;
                border: none;
                border-radius: 5px;
                min-height: 40px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """)
        self.process_btn.clicked.connect(self.process_data)
        main_layout.addWidget(self.process_btn)

    def get_operation_type(self, tab_name: str) -> tuple:
        """Get operation type and truck flag based on radio selection and tab"""
        is_discharge = self.discharge_radio.isChecked()
        
        if tab_name == 'Local':
            return ('DIS' if is_discharge else 'LOD', False)
        elif tab_name == 'Same TS':
            return ('TSD' if is_discharge else 'TSL', False)
        elif tab_name == 'External TS':
            return ('TSD' if is_discharge else 'TSL', True)
        return (None, False)
        
    def process_data(self):
        try:
            # Get ASC file path
            if not self.drop_area.file_path:
                QMessageBox.warning(self, 'Error', 'ASC 파일을 드래그 앤 드롭 해주세요.')
                return
                
            asc_file = self.drop_area.file_path
            if not os.path.exists(asc_file):
                QMessageBox.warning(self, 'Error', 'ASC 파일을 찾을 수 없습니다.')
                return

            # Get container lists from tabs
            tpf_containers = self.tpf_tab.get_container_list()
            truck_containers = self.truck_tab.get_container_list()
            
            # Add containers from operation tabs based on rules
            local_containers = self.local_tab.get_container_list()
            same_ts_containers = self.same_ts_tab.get_container_list()
            external_ts_containers = self.external_ts_tab.get_container_list()
            
            # Add External TS containers to truck containers
            truck_containers.extend(external_ts_containers)
            
            # Get operation type based on radio selection
            operation_type = 'DIS' if self.discharge_radio.isChecked() else 'LOD'
            
            # Get output file
            output_file, _ = QFileDialog.getSaveFileName(
                self, 'Save Summary File', 
                'container_summary.xlsx',
                'Excel Files (*.xlsx)'
            )
            
            if output_file:
                # Create summary
                create_summary(
                    asc_file=asc_file,
                    operation_type=operation_type,
                    tpf_containers=tpf_containers,
                    truck_containers=truck_containers,
                    output_file=output_file
                )
                
                QMessageBox.information(
                    self, 
                    'Success', 
                    f'Summary가 성공적으로 생성되었습니다:\n{output_file}'
                )
                
        except Exception as e:
            QMessageBox.critical(self, 'Error', str(e))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Modern style
    window = ContainerAnalyzerGUI()
    window.show()
    sys.exit(app.exec_()) 