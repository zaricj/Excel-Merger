import sys
import os
from pathlib import Path
import pandas as pd
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, 
                             QLineEdit, QPushButton, QComboBox, QTextEdit, QProgressBar, QStatusBar, QTableView, QGroupBox,
                             QFileDialog, QMessageBox, QSizePolicy, QDialog, QTreeView, QFileSystemModel, QListWidget, QMenu)
from PySide6.QtGui import QAction, QCloseEvent, QIcon, QDropEvent, QStandardItemModel
from PySide6.QtCore import QThread, Signal, QObject, QDir, Qt, QAbstractTableModel, QModelIndex, QFile, QTextStream


class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]
    
    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                # Return the column header
                return str(self._data.columns[section])
            elif orientation == Qt.Vertical:
                # Return the row number
                return str(section + 1)
        return None

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return str(self._data.iloc[index.row(), index.column()])
        return None
    

class DraggableListWidget(QListWidget):
    def __init__(self, table_view: QTableView, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)  # Enable dropping on QLineEdit
        self.items_list = []

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
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()

                # Append the full file path to items_list for later use
                if file_path not in self.items_list:
                    self.items_list.append(file_path)

                    # Get the filename from the full path
                    filename = os.path.basename(file_path)

                    # Add the filename to the list display, but keep the full path in items_list
                    if not self.findItems(filename, Qt.MatchFixedString | Qt.MatchCaseSensitive):
                        self.addItem(filename)

            event.acceptProposedAction()  # Accept the drop event
        else:
            event.ignore()

            
    def show_context_menu_listbox(self, position):
        context_menu = QMenu(self)
        delete_action = QAction("Delete Selected", self)
        delete_all_action = QAction("Delete All", self)
        load_dataframe_action = QAction("Load Into Table View", self)

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
                self.takeItem(current_selected_item)
                self.items_list.pop(current_selected_item)
        except IndexError:
            QMessageBox.information(self, "No Item Selected", "Nothing to delete, fool.")
        except Exception as ex:
            message = f"An exception of type {type(ex).__name__} occurred. Arguments: {ex.args!r}"
            QMessageBox.critical(self, "Exception Error", f"Error removing selected item from list: {message}")
            
    def remove_all_items(self):
        try:
            if self.count() > 0:
                self.items_list.clear()
                self.clear()
                QMessageBox.information(self, "Delete All", "All items have been deleted from the list.")
        except Exception as ex:
            message = f"An exception of type {type(ex).__name__} occurred. Arguments: {ex.args!r}"
            QMessageBox.critical(self, "Exception Error", f"Error removing selected all items from list: {message}")
    
    def load_excel_dataframe_into_table(self):
        try:
            if self.currentRow() != -1:
                self.selected_item = self.items_list[self.currentRow()]
                self.excel_filepath = self.selected_item
            
                try:
                    df = pd.read_excel(self.excel_filepath)

                    # Update the QTableView with the merged data
                    model = PandasModel(df)
                    self.table_view.setModel(model)

                    # QMessageBox.information(self,"Loaded Dataframe", f"Loaded the following excel file into the table: '{os.path.basename(self.excel_filepath)}'")

                except Exception as ex:
                    QMessageBox.critical(self, "Exception Error", f"An exception occurred: {str(ex)}")
            else:
                QMessageBox.information(self,"No File Selected", "Nothing to load, fool.")
                
        except IndexError as ix:
            QMessageBox.critical(self, "Error", f"An Error occurred while trying to load dataframe: {str(ix)}")
            
class MainWindow(QMainWindow):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Merger")
        self.setWindowIcon(QIcon("_internal\\icon\\merge.ico"))
        self.setGeometry(500, 250, 1200, 900)
        self.saveGeometry()
        self.table_view = QTableView()
        self.table_view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        # Theme stuff
        self.light_mode = QIcon("_internal\\imgs\\light.png")
        self.dark_mode = QIcon("_internal\\imgs\\dark.png")
        # Current theme to load on startup
        self.current_theme = "_internal\\themes\\dark.qss"
        
        self.initUI()
        
        # Create the menu bar
        self.create_menu_bar()
        
        # Apply the custom dark theme
        self.initialize_theme(self.current_theme)

        
    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # Left side layout
        left_layout = QVBoxLayout()

        # ListBox GroupBox
        listbox_group = QGroupBox("List Box View")
        listbox_layout = QVBoxLayout()
        self.listbox = DraggableListWidget(self.table_view)
        listbox_layout.addWidget(self.listbox)
        listbox_group.setLayout(listbox_layout)

        # Tree View GroupBox
        treeview_group = QGroupBox("Tree View")
        treeview_layout = QVBoxLayout()
        self.tree_view = QTreeView(self)
        self.tree_view.setDragEnabled(True)  # Enable dragging
        treeview_layout.addWidget(self.tree_view)

        treeview_input_layout = QHBoxLayout()

        self.button_browse_new_filesystem = QPushButton("Browse")
        self.button_browse_new_filesystem.clicked.connect(self.load_new_filesystem_path)
        treeview_input_layout.addWidget(self.button_browse_new_filesystem)
        treeview_layout.addLayout(treeview_input_layout)
        treeview_group.setLayout(treeview_layout)
        
        # Columns GroupBox
        columns_group = QGroupBox("Excel Columns")
        columns_layout_v = QVBoxLayout()
        columns_layout_one = QHBoxLayout()
        columns_layout_two = QHBoxLayout()
        columns_merge_layout = QHBoxLayout()
        find_and_replace_layout = QHBoxLayout()
        excel_input_layout = QHBoxLayout()

        self.input_path_main_excel_file = QLineEdit()
        self.input_path_main_excel_file.setPlaceholderText("Select the excel files to which all data will be merged to...")
        self.button_browse_main_excel_file = QPushButton("Browse")
        self.button_browse_main_excel_file.clicked.connect(self.browse_path_to_main_excel_file)
        excel_input_layout.addWidget(self.input_path_main_excel_file)
        excel_input_layout.addWidget(self.button_browse_main_excel_file)

        self.button_merge = QPushButton("Merge and Apply Replacements")
        self.button_merge.clicked.connect(lambda: self.start_merging_data(self.input_path_main_excel_file.text(), self.get_listbox_items()))

        # Main Excel Widgets
        self.main_excel_label = QLabel("Compare and match values from column")
        self.main_excel_column_combobox = QComboBox()
        self.main_excel_column_combobox.setEditable(True)
        self.main_excel_label_two = QLabel("Compare and match values against column")
        self.to_compare_excel_column_combobox = QComboBox()
        self.to_compare_excel_column_combobox.setEditable(True)
        
        self.main_excel_load_column_button = QPushButton("Load Main Columns")
        self.main_excel_load_column_button.clicked.connect(self.load_main_excel_columns)
        self.to_compare_load_columns_button = QPushButton("Load Other Columns")
        self.to_compare_load_columns_button.clicked.connect(self.load_to_compare_excel_columns)
        
        self.merge_value_label = QLabel("Merge all values from column of to merge file")
        self.merge_value_combobox = QComboBox()
        self.merge_value_combobox.setEditable(True)
        self.merge_value_label_on_main = QLabel("on the main excel file.")
        
        # Find and Replace
        self.find_input = QLineEdit()
        self.find_input.setPlaceholderText(" Values (comma-separated) to find...")
        self.replace_input = QLineEdit()
        self.replace_input.setPlaceholderText("Values (comma-separated) to replace...")

        # Labels for the inputs
        find_label = QLabel("Find:")
        replace_label = QLabel("Replace:")

        columns_layout_one.addWidget(self.main_excel_label)
        columns_layout_one.addWidget(self.main_excel_column_combobox)
        columns_layout_one.addWidget(self.main_excel_load_column_button)
        columns_layout_two.addWidget(self.main_excel_label_two)
        columns_layout_two.addWidget(self.to_compare_excel_column_combobox)
        columns_layout_two.addWidget(self.to_compare_load_columns_button)

        columns_merge_layout.addWidget(self.merge_value_label)
        columns_merge_layout.addWidget(self.merge_value_combobox)
        columns_merge_layout.addWidget(self.merge_value_label_on_main)
        
        find_and_replace_layout.addWidget(find_label)
        find_and_replace_layout.addWidget(self.find_input)
        find_and_replace_layout.addWidget(replace_label)
        find_and_replace_layout.addWidget(self.replace_input)

        columns_layout_v.addLayout(columns_layout_one)
        columns_layout_v.addLayout(columns_layout_two)
        columns_layout_v.addLayout(columns_merge_layout)
        columns_layout_v.addLayout(find_and_replace_layout)
        columns_layout_v.addLayout(excel_input_layout)
        columns_layout_v.addWidget(self.button_merge)
        columns_group.setLayout(columns_layout_v)

        # Add GroupBoxes to left layout
        left_layout.addWidget(treeview_group)
        left_layout.addWidget(listbox_group)
        left_layout.addWidget(columns_group)

        # Add left layout to main layout
        main_layout.addLayout(left_layout)

        # Right side: QTableView GroupBox
        table_group = QGroupBox("Table View")
        table_layout = QVBoxLayout()
        table_layout.addWidget(self.table_view)
        table_group.setLayout(table_layout)

        # Add the table group to a vertical layout
        right_layout = QVBoxLayout()
        right_layout.addWidget(table_group)

        # Add right layout to the main layout
        main_layout.addLayout(right_layout)

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
        
        # Change Theme Toggle
        self.toggle_theme_action = menu_bar.addAction(self.light_mode, "Toggle Theme")
        self.toggle_theme_action.triggered.connect(self.change_theme)
        
    def change_theme(self):
        if self.current_theme == "_internal\\themes\\dark.qss":
            self.current_theme = "_internal\\themes\\light.qss"
            self.toggle_theme_action.setIcon(self.dark_mode)
            self.toggle_theme_action.setText("Toggle Theme")
        else:
            self.current_theme = "_internal\\themes\\dark.qss"
            self.toggle_theme_action.setIcon(self.light_mode)
            self.toggle_theme_action.setText("Toggle Theme")
            
        self.initialize_theme(self.current_theme)
        
    def closeEvent(self, event: QCloseEvent):
        reply = QMessageBox.question(self, "Exit Program", "Are you sure you want to exit the program?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
     
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
    
    # Method - Load new filesystem path into the TreeView element.
    def load_new_filesystem_path(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Directory")
        if folder:
            try:
                self.file_system_model.setRootPath(folder)
                self.tree_view.setRootIndex(self.file_system_model.index(folder))
            except Exception as ex:
                QMessageBox.critical(self, "Exception Error", f"An exception of type {type(ex).__name__} occurred. Arguments: {ex.args!r}")
        else:
            pass
        
    def browse_path_to_main_excel_file(self):
        options = QFileDialog.Options()
        
        file, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file:
            self.input_path_main_excel_file.setText(file)
        
    def get_listbox_items(self):
        # Directly access the items_list attribute from the DraggableListWidget instance
        items = self.listbox.items_list
        if items:
            return items
        else:
            None

    def initialize_theme(self, theme_file):
        try:
            file = QFile(theme_file)
            if file.open(QFile.ReadOnly | QFile.Text):
                stream = QTextStream(file)
                stylesheet = stream.readAll()
                self.setStyleSheet(stylesheet)
            file.close()
        except Exception as ex:
            QMessageBox.critical(self, "Theme load error", f"Failed to load theme: {str(ex)}")
            
    # Methods for Loading Columns from Excel Files
    # Load all Columns from the Main Excel Files which the values will be merged to
    def load_main_excel_columns(self):
        try:
            options = QFileDialog.Options()
            file, _ = QFileDialog.getOpenFileName(self, "Select Main Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
            excel_file = file
            
            if Path(excel_file).is_file():
                main_df = pd.read_excel(excel_file)
                self.main_excel_column_combobox.clear()
                self.main_excel_column_combobox.addItems(main_df.columns)
                QMessageBox.information(self,"Columns Loaded","Successfully loaded all columns into the ComboBox.")
            else:
                QMessageBox.information(self, "Cannot Load File","Cannot load excel file.")
            
        except Exception as ex:
            QMessageBox.critical(self, "Exception Error", f"An exception occurred: {str(ex)}")
            
    def load_to_compare_excel_columns(self):
        try:
            options = QFileDialog.Options()
            file, _ = QFileDialog.getOpenFileName(self, "Select Main Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
            excel_file = file
            
            if Path(excel_file).is_file():
                main_df = pd.read_excel(excel_file)
                self.to_compare_excel_column_combobox.clear()
                self.merge_value_combobox.clear()
                self.to_compare_excel_column_combobox.addItems(main_df.columns)
                self.merge_value_combobox.addItems(main_df.columns)
                QMessageBox.information(self,"Columns Loaded","Successfully loaded all columns into the ComboBox.")
            else:
                QMessageBox.information(self, "Cannot Load File","Cannot load excel file.")
        except Exception as ex:
            QMessageBox.critical(self, "Exception Error", f"An exception occurred: {str(ex)}")
        
    def start_merging_data(self, main_excel_file_path=None, list_of_excel_files_to_merge=None):
        """_summary_

        Args:
            main_excel_file_path (_type_, optional): _description_. Defaults to None.
            list_of_excel_files_to_merge (_type_, optional): _description_. Defaults to None.

        Raises:
            ValueError: _description_
            ValueError: _description_
        """
        try:
            find_patterns = [p.strip() for p in self.find_input.text().split(',') if p.strip()]
            replace_patterns = [p.strip() for p in self.replace_input.text().split(',') if p.strip()]

            if self.listbox.count() > 0 and os.path.isfile(main_excel_file_path):
                main_df = pd.read_excel(main_excel_file_path)

                # Get selected key columns
                main_df_key = self.main_excel_column_combobox.currentText()
                to_compare_df_key = self.to_compare_excel_column_combobox.currentText()
                value_to_merge_key = self.merge_value_combobox.currentText()

                if not main_df_key or not to_compare_df_key or not value_to_merge_key:
                    QMessageBox.warning(self, "Warning", "Please select valid keys for the columns from both files!")
                    return

                for excel_file in list_of_excel_files_to_merge: 
                    excel_file_df = pd.read_excel(excel_file, dtype={to_compare_df_key: str, value_to_merge_key: str})
                    excel_file_df[to_compare_df_key] = excel_file_df[to_compare_df_key].fillna("?")
                    excel_file_df[value_to_merge_key] = excel_file_df[value_to_merge_key].fillna("?")
                    excel_filename = Path(excel_file).stem

                    if to_compare_df_key not in excel_file_df.columns:
                        raise ValueError(f"File {excel_file} column mismatch.")
                    if value_to_merge_key not in excel_file_df.columns:
                        raise ValueError(f"File {excel_file} column mismatch.")

                    # Perform the merge
                    merged_df = main_df.merge(excel_file_df[[to_compare_df_key, value_to_merge_key]], on=main_df_key, how="left")

                    # Create a new column based on the filename to store the merged values
                    merged_df[excel_filename] = merged_df[value_to_merge_key].copy()

                    # Apply find-replace on the new column
                    for find_pattern, replace_pattern in zip(find_patterns, replace_patterns):
                        # Use a vectorized replace across the new column
                        merged_df[excel_filename] = merged_df[excel_filename].replace(find_pattern, replace_pattern)

                    # Fill any remaining NaN values with '?'
                    merged_df.fillna({excel_filename: "?"}, inplace=True)

                    # Drop the temporary merge column after processing
                    main_df = merged_df.drop(columns=[value_to_merge_key])

                # Save the updated file
                updated_file_path = os.path.splitext(main_excel_file_path)[0] + '_updated.xlsx'
                main_df.to_excel(updated_file_path, index=False)

                QMessageBox.information(self, "Merging Finished", f"Updated file saved as: {updated_file_path}")
            else:
                 QMessageBox.information(self, "No Excel File", "No excel files selected.")
                
        except Exception as ex:
            message = f"An exception of type {type(ex).__name__} occurred. Arguments: {ex.args!r}"
            QMessageBox.critical(self, "Exception Error", f"Error merging data: {message}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
