import sys
from pathlib import Path
import os
import pandas as pd
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, 
                             QLineEdit, QPushButton, QComboBox, QTextEdit, QProgressBar, QStatusBar, QTableView,
                             QFileDialog, QMessageBox, QSizePolicy, QDialog, QTreeView, QFileSystemModel, QListWidget, QMenu)
from PySide6.QtGui import QAction, QCloseEvent, QIcon, QDropEvent, QStandardItemModel
from PySide6.QtCore import QThread, Signal, QObject, QDir, Qt, QAbstractTableModel, QModelIndex


class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return str(self._data.iloc[index.row(), index.column()])
        return None
    

class DraggableListWidget(QListWidget):
    def __init__(self, program_output: QTextEdit, table_view: QTableView, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)  # Enable dropping on QLineEdit
        self.items_list = []
        self.program_output = program_output
        self.table_view = table_view
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu_listbox)

    def dragEnterEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()  # Accept the drag event
        else:
            event.ignore()

    def dragMoveEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()  # Accept the drag move event
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            # Extract the file path from the drop event
            file_path = event.mimeData().urls()[0].toLocalFile()

            # Append the full file path to items_list for later use
            if file_path not in self.items_list:
                self.items_list.append(file_path)
                event.acceptProposedAction()  # Accept the drop event   

                # Get the filename from the full path
                filename = os.path.basename(file_path)

                # Add the filename to the list display, but keep the full path in items_list
                if not self.findItems(filename, Qt.MatchFixedString | Qt.MatchCaseSensitive):
                    self.addItem(filename)  

                self.update_listbox_items_count()
            else:
                self.program_output.setText("File already exists in the list.")
        else:
            event.ignore()
            
    def show_context_menu_listbox(self, position):
        context_menu = QMenu(self)
        delete_action = QAction("Delete Selected", self)
        delete_all_action = QAction("Delete All", self)
        load_dataframe_action = QAction("Load Selected Dataframe", self)

        context_menu.addAction(delete_action)
        context_menu.addAction(delete_all_action)
        context_menu.addSeparator()
        context_menu.addAction(load_dataframe_action)

        delete_action.triggered.connect(self.remove_selected_item)
        delete_all_action.triggered.connect(self.remove_all_items)
        load_dataframe_action.triggered.connect(self.load_excel_dataframe_into_table)
        # Show the context menu at the cursor's current position
        context_menu.exec(self.mapToGlobal(position))
    
    def remove_selected_item(self):
        try:
            current_selected_item = self.currentRow()
            if current_selected_item != -1:
                item_to_remove = self.takeItem(current_selected_item)
                self.items_list.pop(current_selected_item)
                self.update_listbox_items_count()
                self.program_output.append(f"Removed item: {item_to_remove.text()} at row {current_selected_item}")
            else:
                self.program_output.append("No item selected to delete.")
        
        except IndexError:
            self.program_output.append("Nothing to delete.")
        except Exception as ex:
            message = f"An exception of type {type(ex).__name__} occurred. Arguments: {ex.args!r}"
            self.program_output.setText(f"Error removing selected item from list: {message}")
            
    def remove_all_items(self):
        try:
            if self.count() > 0:
                self.items_list.clear()
                self.clear()
                self.program_output.setText("Deleted all items from the list.")
            else:
                self.program_output.setText("No items to delete.")
                
            self.update_listbox_items_count()
            
        except Exception as ex:
            message = f"An exception of type {type(ex).__name__} occurred. Arguments: {ex.args!r}"
            self.program_output.setText(f"Error removing selected all items from list: {message}")
    
    
    def load_excel_dataframe_into_table(self):
        try:
            self.selected_item = self.items_list[self.currentRow()]
            self.excel_filepath = self.selected_item

            try:
                df = pd.read_excel(self.excel_filepath)

                # Update the QTableView with the merged data
                model = PandasModel(df)
                self.table_view.setModel(model)

                QMessageBox.information(self,"Loaded Dataframe", f"Loaded the following excel file into the table: '{os.path.basename(self.excel_filepath)}'")

            except Exception as ex:
                self.program_output.setText(f"Exception occurred: {str(ex)}")
                
        except IndexError as ix:
            QMessageBox.critical(self, "Error", f"Error occurred while trying to load dataframe: {str(ix)}")
            
        
    def update_listbox_items_count(self):
        self.counter = self.count()
        if self.counter != 0:
            self.program_output.append(f"Total number of items in List: {self.counter}")
    
            
class MainWindow(QMainWindow):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Merger")
        self.setWindowIcon(QIcon("icon\\merge.ico"))
        self.setGeometry(500, 250, 1000, 700)
        self.saveGeometry()
        self.program_output = QTextEdit()
        self.table_view = QTableView()
        self.table_view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.initUI()
        
        # Create the menu bar
        self.create_menu_bar()
        
        # Apply the custom dark theme
        self.apply_custom_dark_theme()
        
    def initUI(self):
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # ListBox and Tree View
        listbox_and_treeview_layout = QHBoxLayout()

        # Left side: ListBox Element
        listbox_vertical_layout = QVBoxLayout()
        listbox_horizontal_layout = QHBoxLayout()
        self.listbox_label = QLabel("List Box View:")
        self.listbox = DraggableListWidget(self.program_output, self.table_view)
        listbox_vertical_layout.addWidget(self.listbox_label)
        listbox_vertical_layout.addWidget(self.listbox)
        listbox_vertical_layout.addLayout(listbox_horizontal_layout)
        self.input_path_main_excel_file = QLineEdit()
        self.input_path_main_excel_file.setPlaceholderText("Select the excel files to which all data will be merged to...")
        self.button_browse_main_excel_file = QPushButton("Browse")
        self.button_browse_main_excel_file.clicked.connect(self.browse_path_to_main_excel_file)
        self.button_merge = QPushButton("Merge")
        self.button_merge.clicked.connect(lambda: self.start_merging_data(self.input_path_main_excel_file.text(), self.get_listbox_items()))
        listbox_horizontal_layout.addWidget(self.input_path_main_excel_file)
        listbox_horizontal_layout.addWidget(self.button_browse_main_excel_file)
        listbox_vertical_layout.addWidget(self.button_merge)
        
        program_output_layout = QVBoxLayout()
        
        self.program_output_label = QLabel("Program Output:")
        self.program_output.setReadOnly(True)
        self.table_view_label = QLabel("Table View:")
        program_output_layout.addWidget(self.table_view_label)
        program_output_layout.addWidget(self.table_view)
        program_output_layout.addWidget(self.program_output_label)
        program_output_layout.addWidget(self.program_output)

        # Right side: Tree View
        tree_view_vertical_layout = QVBoxLayout()
        self.tree_view_label = QLabel("Tree View:")
        self.tree_view = QTreeView(self)
        self.tree_view.setDragEnabled(True)  # Enable dragging
        tree_view_horizontal_layout = QHBoxLayout()
        self.new_filesystem_input = QLineEdit()
        self.new_filesystem_input.setPlaceholderText("Select a path to the excel files for merging...")
        self.button_browse_new_filesystem = QPushButton("Browse")
        self.button_browse_new_filesystem.clicked.connect(self.load_new_filesystem_path)
        tree_view_vertical_layout.addWidget(self.tree_view_label)
        tree_view_vertical_layout.addWidget(self.tree_view)
        tree_view_horizontal_layout.addWidget(self.new_filesystem_input)
        tree_view_horizontal_layout.addWidget(self.button_browse_new_filesystem)
        tree_view_vertical_layout.addLayout(tree_view_horizontal_layout)

        # Add both vertical layouts to the horizontal layout
        listbox_and_treeview_layout.addLayout(listbox_vertical_layout)
        listbox_and_treeview_layout.addLayout(tree_view_vertical_layout)

        # Add the horizontal layout to the main vertical layout
        layout.addLayout(listbox_and_treeview_layout)
        layout.addLayout(program_output_layout) # Added program output as last element
        
        # Set up the file system model
        self.file_system_model = QFileSystemModel(self)
        self.file_system_model.setRootPath("")  # Set root path to the filesystem's root
        self.file_system_model.setFilter(QDir.NoDotAndDotDot | QDir.AllDirs | QDir.Files)  # Show all dirs and files

        # Set the model to the tree view
        self.tree_view.setModel(self.file_system_model)
        self.tree_view.setRootIndex(self.file_system_model.index(""))  # Set root index to the user's home directory

        # Optional: Customize the view
        self.tree_view.setColumnWidth(0, 250)  # Adjust column width
        self.tree_view.setHeaderHidden(False)   # Show the header
        self.tree_view.setSortingEnabled(True)  # Enable sorting
        
     # ============ Methods for the MainWindow Class ============ # 
     
    def create_menu_bar(self):
        menu_bar = self.menuBar()
        
         # File Menu
        file_menu = menu_bar.addMenu("&File")
        clear_action = QAction("Clear Output", self)
        clear_action.setStatusTip("Clear the program output")
        clear_action.triggered.connect(self.clear_output)
        file_menu.addAction(clear_action)
        file_menu.addSeparator()
        exit_action = QAction("E&xit", self)
        exit_action.setStatusTip("Exit the application")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Open Menu
        open_menu = menu_bar.addMenu("&Open")
        open_input_action = QAction("Open Input Folder", self)
        open_input_action.setStatusTip("Opens the input folder")
        #open_input_action.triggered.connect(self.open_input_folder)
        open_menu.addAction(open_input_action)
        open_output_action = QAction("Open Output Folder ", self)
        open_output_action.setStatusTip("Opens the output folder")
        #open_output_action.triggered.connect(self.open_output_folder)
        open_menu.addAction(open_output_action)   
        
    def closeEvent(self, event: QCloseEvent):
        reply = QMessageBox.question(self, "Exit Program", "Are you sure you want to exit the   program?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
     
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
    
    # Method - Load new filesystem path into the TreeView element.
    def load_new_filesystem_path(self):
        
        folder = QFileDialog.getExistingDirectory(self, "Select Directory")
        if folder:
            self.new_filesystem_input.setText(folder)
            
        try:
            folder_path = self.new_filesystem_input.text()
            if folder_path and QDir(folder_path).exists():
                self.file_system_model.setRootPath(folder_path)
                self.tree_view.setRootIndex(self.file_system_model.index(folder_path))
                self.program_output.setText(f"Tree View updated to the following path: {folder_path}")
                self.new_filesystem_input.clear()
            else:
                pass
        except Exception as ex:
            self.program_output.setText(f"An exception of type {type(ex).__name__} occurred. Arguments: {ex.args!r}")
    
    def browse_path_to_main_excel_file(self):
        
        options = QFileDialog.Options()
        
        file, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file:
            self.input_path_main_excel_file.setText(file)
    
    def clear_output(self):
        self.program_output.clear()
        
    def get_listbox_items(self):
        # Directly access the items_list from the DraggableListWidget instance
        items = self.listbox.items_list
        return items
        
    def start_merging_data(self, main_excel_file_path, list_of_excel_files_to_merge):
        try:
            self.program_output.setText(f"Merging data to {os.path.basename(main_excel_file_path)}...")
            if self.listbox.count() > 0:
                if os.path.isfile(main_excel_file_path):
                    main_df = pd.read_excel(main_excel_file_path)

                    if "Profile" not in main_df.columns:
                        raise ValueError("Main file must have a 'Profile' column.")

                    for excel_file in list_of_excel_files_to_merge: 
                        excel_file_df = pd.read_excel(excel_file)

                        # Ensure 'Profile' and the value column exist in the month file
                        if 'Profile' not in excel_file_df.columns:
                            raise ValueError(f"File {excel_file} must have a 'Profile' column")
                        if 'Value' not in excel_file_df.columns:  # Assuming 'value' is the column with 1 or 0
                            raise ValueError(f"File {excel_file} must have a 'Value' column (1 or 0)")

                        excel_filename = Path(excel_file).stem
                        print(excel_filename)

                        # Merge the main file with the current month file on 'Profile'
                        merged_df = main_df.merge(excel_file_df[['Profile', 'Value']], on='Profile', how='left')

                        # Replace the 1/0 values with 'X' and '-' and fill NaN with '?'
                        merged_df[excel_filename] = merged_df['Value'].apply(lambda x: 'x' if x == 1 else '-')
                        merged_df.fillna({excel_filename: "?"}, inplace=True)

                        # Drop the temporary 'Value' column after merging
                        main_df = merged_df.drop(columns=['Value'])

                    updated_file_path = os.path.splitext(main_excel_file_path)[0] + '_updated.xlsx'
                    main_df.to_excel(updated_file_path, index=False)

                    self.program_output.setText(f"Updated file saved as: {updated_file_path}")
                else:
                    self.program_output.setText("No excel file selected.")
            else:
                self.program_output.setText("Listbox does not contain any items.")
                
        except Exception as ex:
            message = f"An exception of type {type(ex).__name__} occurred. Arguments: {ex.args!r}"
            self.program_output.setText(f"Error merging data: {message}")
            
    def apply_custom_dark_theme(self):
        self.setStyleSheet("""
        QMainWindow, QWidget {
            background-color: #1e2329;
            color: #e0e0e0;
            font-family: 'Segoe UI', 'Roboto', sans-serif;
            font-size: 12px;
        }
        QLabel {
            color: #e0e0e0;
            font-weight: 500;
            font: bold;
        }
        QLineEdit, QTextEdit, QTreeView {
            background-color: #2a3038;
            border: 1px solid #3a4149;
            border-radius: 6px;
            padding: 8px;
            color: #e0e0e0;
        }
        QLineEdit:focus, QTextEdit:focus, QTreeView:focus {
            border-color: #4caf50;
            background-color: #2f363f;
        }
        QPushButton {
            background-color: #4caf50;
            color: #ffffff;
            border-radius: 6px;
            padding: 10px 16px;
            font-weight: 600;
            min-width: 100px;
        }
        QPushButton:hover {
            background-color: #45a049;
            color: #000000;
        }
        QPushButton:pressed {
            background-color: #3d8b40;
        }
        QPushButton:disabled {
            background-color: #5c6370;
            color: #9da5b4;
        }
        QTreeView::item:selected, QListWidget::item:selected {
            background-color: #4caf50;
        }
        QMenuBar {
            border-bottom: 2px solid #4caf50;
            background-color: #1e2329;
            color: #e0e0e0;
        }
        QMenuBar::item:selected {
            background-color: #2a3038;
        }
        QMenu {
            background-color: #1e2329;
            color: #e0e0e0;
            border: 1px solid #3a4149;
            border-radius: 6px;
        }
        QMenu::item:selected {
            background-color: #4caf50;
        }
        QStatusBar {
            background-color: #171b20;
            color: #e0e0e0;
        }
        QListWidget {
            background-color: #2a3038;
            color: #e0e0e0;
        }
        QListWidget::item {
            height: 32px;
        }
        QScrollBar:vertical {
            border: none;
            background-color: #2a3038;
            width: 10px;
            margin: 0px 0px 0px 0px;
        }
        QScrollBar::handle:vertical {
            background-color: #4caf50;
            min-height: 30px;
            border-radius: 5px;
        }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }
        QScrollBar:horizontal {
            border: none;
            background-color: #2a3038;
            height: 10px;
            margin: 0px 0px 0px 0px;
        }
        QScrollBar::handle:horizontal {
            background-color: #4caf50;
            min-width: 30px;
            border-radius: 5px;
        }
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
            width: 0px;
        }
        """)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
