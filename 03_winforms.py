import clr
import math
import System

# Verify these are needed.
clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference("System.Windows.Forms")

# Excel
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel

# Windows Forms Elements
from System.Drawing import Point, Icon, Color
from System.Windows import Forms
from System.Windows.Forms import Application, Form, ControlPaint, MessageBox, MessageBoxButtons, MessageBoxIcon
from System.Windows.Forms import DialogResult, GroupBox, FormBorderStyle, ComboBox, BorderStyle, Label, ComboBoxStyle, DataGridViewRowHeadersWidthSizeMode
from System.Windows.Forms import Button, OpenFileDialog, DataGridView, DataGridViewRow, DataGridViewColumnSortMode, ListBox

class excel_paramter:
    def __init__(self, r, a = None, b = None, c = 0, d = None):
        self.raw_data = r
        self.index_title_row = a
        self.index_unit_row = b
        self.index_data_start_row = c
        self.index_column_identifier = d
        
class form_interface(Form):
    def __init__(self, title):
        #output
        self.excel_parameters = []
        self.excel_file_parameters = []
        #window Settings
        self.Text = title
        self.TopMost = True
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MinimizeBox = False
        self.MaximizeBox = False
        self.ShowIcon = False
        self.Height = 510
        self.Width = 820
        self.CenterToScreen() 
        
        #dynamic input
        self.collection_controls = []
        self.counter = 0
        
        #add proceed
        self.button_proceed = self.__add_button(150, 80, 640, 350)
        self.button_proceed.Text = "Proceed"
        self.button_proceed.Click += self.__click_proceed
        self.button_proceed.Enabled = False
        
        ##add cancel
        self.button_cancel = self.__add_button(150, 25, 640, 435)
        self.button_cancel.Text = "Cancel"
        self.button_cancel.Click += self.__click_cancel
                
        #fixed input
        self.combo_file_num = self.__add_combo_box(600, 20, 20, 20)
        self.combo_file_num .SelectedIndexChanged += self.__update_file_num_selection
        self.__set_file_num()
        self.label_file_num = self.__add_label(150, 20, 640, 20, "Number of Excel Files")
        self.label_file_num.Font = System.Drawing.Font(self.label_file_num.Font.FontFamily, 9.5, System.Drawing.FontStyle.Bold)
        self.label_file_num.BackColor = ControlPaint.LightLight(ControlPaint.LightLight(self.label_file_num.ForeColor))
        
        self.show()
        
    def __click_proceed(self, sender, event):
        eps = []
        if [] in self.excel_file_parameters:
            MessageBox.Show("Have you selected all excel files?", "ERROR: Please revise selection", MessageBoxButtons.OK, MessageBoxIcon.Error)
        else:
            for list in self.excel_file_parameters:
                ep = excel_paramter(list[0], list[1], list[2], list[3], list[4])
                eps.append(ep)
            self.excel_parameters = eps
            self.DialogResult = DialogResult.OK
        
    def __click_cancel(self, sender, event):
        self.excel_file_parameters = []
        self.DialogResult = DialogResult.Cancel
        
    def __click_select_excel(self, sender, event):
        button_index = int(sender.Text.split(" > ")[0])
        self.excel_file_parameters[button_index] = []
        form_excel = form_excel_preview(sender.Text)
        if form_excel.DialogResult == DialogResult.OK:
            sender.Text = str(button_index) + " > " + form_excel.button_select.Text
            parameter_list = [
            form_excel.raw_excel_data,
            form_excel.index_title_row,
            form_excel.index_unit_row,
            form_excel.index_data_start_row,
            form_excel.index_column_identifier,
            ]
            self.excel_file_parameters[button_index] = parameter_list
            self.button_proceed.Enabled = True
  
    def __update_file_num_selection(self, sender, event):
        #clear controls
        self.excel_file_parameters = []
        self.counter = 0
        self.button_proceed.Enabled = False
        self.excel_file_parameters = []
        for control in self.collection_controls:
            self.Controls.Remove(control)
        current_num = self.combo_file_num.SelectedIndex
        buttons_labels = []
        #add controls
        for i in range(current_num + 1):
            colour = Color.OrangeRed
            text_label = "Other Excel File"
            if i == 0:
                text_label = "Master Excel File"
                colour = Color.SeaGreen
            text_button = str(self.counter) + " > Select Excel File"
            button = self.__add_button(600, 25, 20, 60 + i*30, text_button)
            button.Click += self.__click_select_excel
            label = self.__add_label(150, 25, 640, 60 + i*30, text_label)
            label.Font = System.Drawing.Font(label.Font.FontFamily, 9.5, System.Drawing.FontStyle.Bold)
            label.ForeColor = colour
            label.BackColor = ControlPaint.LightLight(ControlPaint.LightLight(label.ForeColor))
            buttons_labels.append(button)
            buttons_labels.append(label)
            self.counter += 1
            self.excel_file_parameters.append([])
            #buttons_labels.append(label)
        self.collection_controls = buttons_labels
        
    def __entry_combo_box (self, combo_box, list):
        combo_box.Items.Clear()
        for item in list:
           combo_box.Items.Add(item)
           
    def __set_file_num(self):
        combo = self.combo_file_num
        self.__entry_combo_box(combo, range(10)[1:])
        combo.SelectedIndex = 0
    
    def __add_button(self, width, height, x, y, text = None):
        button = Button()
        button.Text = text
        button.Width = width
        button.Height = height
        button.Location = Point(x, y)
        self.Controls.Add(button)
        return button
        
    def __add_label(self, width, height, x, y, text = None):
        label = Label()
        label.Text = text
        label.Font = System.Drawing.Font(label.Font.FontFamily, 9.5, System.Drawing.FontStyle.Regular)
        label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        label.Width = width
        label.Height = height
        label.Location = Point(x, y)
        self.Controls.Add(label)
        return label
    
    def __add_combo_box(self, width, height, x, y):
        combo_box = ComboBox()
        combo_box.DropDownStyle = ComboBoxStyle.DropDownList
        combo_box.Text = "None"
        combo_box.Width = width
        combo_box.Height = height
        combo_box.Location = Point(x, y)
        self.Controls.Add(combo_box)
        return combo_box
        
    def __add_label_button_groups(self):
        button_text_lists = ["Select Master Excel File"]
        self.__add_button
    
    def show(self):
        self.ShowDialog()

class form_options(Form):
    def __init__(self, option_list):
        #output
        self.selected_index = None
        #settings
        self.Text = None
        self.TopMost = True
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow
        self.MinimizeBox = False
        self.MaximizeBox = False
        self.ShowIcon = False
        self.Height = 300
        self.Width = 300
        self.CenterToParent()
        
        self.box_items = self.__add_list_box(270, 200, 7.5, 0)
        self.button_confirm = self.__add_button(270, 50, 7.5, 205)
        self.button_confirm.Text = "Please select 1 option"
        self.box_items.SelectedIndexChanged += self.__update_selected_index
        self.button_confirm.Click += self.click_button_confirm
        
        self.add_items(option_list)
        self.box_items.SelectedIndex = 0
        
    def __add_list_box(self, width, height, x, y):
        list_box = ListBox()
        list_box.Width = width
        list_box.Height = height
        list_box.Location = Point(x, y)
        self.Controls.Add(list_box)
        return list_box
        
    def __add_button(self, width, height, x, y):
        button = Button()
        button.Width = width
        button.Height = height
        button.Location = Point(x, y)
        self.Controls.Add(button)
        return button
        
    def __add_items(self, list_box, item_list):
        list_box.Items.Clear()
        for item in item_list:
            list_box.Items.Add(item)
    
    def add_items(self, item_list):
        self.__add_items(self.box_items, item_list)
        
    def click_button_confirm(self, sender, event):
        if self.box_items.SelectedIndex >= 0:
            self.selected_index = self.box_items.SelectedIndex
            self.DialogResult = DialogResult.OK
        else:
           self.button_confirm.Text = "Nothing selected"
    
    def __update_selected_index(self, sender, event):
        self.button_confirm.Text = "Confirm selection"
        
    def show(self):
        self.ShowDialog()
        
class form_excel_preview(Form):
    def __init__(self, title):
        #output
        self.raw_excel_data = None
        self.index_title_row = None
        self.index_unit_row = None
        self.index_data_start_row = None
        self.index_column_identifier = None
        
        #window Settings
        self.Text = title
        self.TopMost = True
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow
        self.MinimizeBox = False
        self.MaximizeBox = False
        self.ShowIcon = False
        self.Height = 820
        self.Width = 1020
        self.CenterToScreen()
        
        #add controls
        self.button_select = self.add_button(1000, 30, 2.5, 600)
        self.button_select.Text = "Select an Excel File"
        self.button_select.Click += self.click_select_excel
        
        #add excel preview
        self.grid_excel = self.add_grid(1000, 600, 2.5, 0)
        
        #add_combo_label
        self.combo_boxes = None
        self.add_combo_label()
        
        #add instruction label
        self.label_instruction = self.add_label(340, 120, 480, 650)
        self.label_instruction.Text = ""
        self.label_instruction.Font = System.Drawing.Font(self.label_instruction.Font.FontFamily, 9.5, System.Drawing.FontStyle.Regular)
        self.label_instruction.TextAlign = System.Drawing.ContentAlignment.TopLeft
        
        #add proceed
        self.button_proceed = self.add_button(150, 80, 840, 650)
        self.button_proceed.Text = "Proceed"
        self.button_proceed.Click += self.click_proceed
        self.button_proceed.Enabled = False
        
        ##add cancel
        self.button_cancel = self.add_button(150, 25, 840, 735)
        self.button_cancel.Text = "Cancel"
        self.button_cancel.Click += self.click_cancel
        
        #show
        self.show()
        
        
    def click_proceed(self, sender, event):
        id_0 = self.combo_boxes[0].SelectedIndex
        id_1 = self.combo_boxes[1].SelectedIndex
        id_2 = self.combo_boxes[2].SelectedIndex
        id_3 = self.combo_boxes[3].SelectedIndex

        summary_text = "Tile Row # = " + self.display_number(id_0) + "\nUnit Row # = " + self.display_number(id_1) + "\nData Start from Row # " + self.display_number(id_2) + "\nIdentifier Column # = " + self.display_number(id_3, True)
        
        ids = [id_0, id_1, id_2]
        id_set = []
        for id in ids:
            if id > 0:
                id_set.append(id)
        if len(id_set) != len(set(id_set)):
            MessageBox.Show("The selected rows must NOT share number!\n\n" + summary_text, "ERROR: Please revise selection", MessageBoxButtons.OK, MessageBoxIcon.Error)            
        else:
            self.index_title_row = id_0
            self.index_unit_row = id_1
            self.index_data_start_row = id_2
            if self.index_data_start_row == None or self.index_data_start_row == 0:
                self.index_data_start_row = 1
            self.index_column_identifier = self.display_number(id_3, True)
            MessageBox.Show(summary_text, "SUCCESSFUL", MessageBoxButtons.OK, MessageBoxIcon.Information)
            self.DialogResult = DialogResult.OK
    
    def display_number(self, number, alpha = False):
        if number == 0:
            return "None"
        if alpha == False:
            return str(number)
        else:
            return self.base_transformer(number-1)

    def click_cancel(self, sender, event):
        self.DialogResult = DialogResult.Cancel
    
    def colour_data_grid_clear(self, data_grid):
        for i in range(data_grid.Rows.Count):
            row = data_grid.Rows[i]
            for cell in row.Cells:
                cell.Style.BackColor = Color.White
                cell.Style.ForeColor = Color.Black

    def colour_data_grid_row(self, data_grid, row_index, colour):
        for cell in data_grid.Rows[row_index].Cells:
            cell.Style.BackColor = ControlPaint.LightLight(colour)
            cell.Style.ForeColor = ControlPaint.Dark(colour)
            
    def colour_data_grid_column(self, data_grid, column_index):
        for i in range(data_grid.Rows.Count):
            current_back_colour = data_grid.Rows[i].Cells[column_index].Style.BackColor
            current_fore_colour = data_grid.Rows[i].Cells[column_index].Style.ForeColor
            data_grid.Rows[i].Cells[column_index].Style.BackColor = ControlPaint.Dark(current_back_colour)
            data_grid.Rows[i].Cells[column_index].Style.ForeColor = Color.White
            
    def update_combo_selection(self, sender, event):
        #selected clear
        self.grid_excel.ClearSelection()
        #clear colour
        self.colour_data_grid_clear(self.grid_excel)
        #data rows
        data_combo = self.combo_boxes[2]
        data_selected_index = data_combo.SelectedIndex
        if data_selected_index > 0:
            data_row_index = data_selected_index - 1
            for i in range(self.grid_excel.Rows.Count)[data_row_index:]:
                self.colour_data_grid_row(self.grid_excel, i, ControlPaint.Light(data_combo.ForeColor))
        #title unit rows
        for combo in self.combo_boxes[:2]:
            if combo.SelectedIndex > 0:
                row_index = combo.SelectedIndex - 1
                self.colour_data_grid_row(self.grid_excel, row_index, ControlPaint.Light(combo.ForeColor))
        #column
        column_combo = self.combo_boxes[3]
        if column_combo.SelectedIndex > 0:
            column_index = column_combo.SelectedIndex - 1
            self.colour_data_grid_column(self.grid_excel, column_index)
            
    def add_combo_label(self):
        label_texts = ["Title / Parameter Row Number ( If Exist )", "Unit Row Number ( If Exist )", "Data Start Row Number", "Required for non-master sheets: Select The Identifier Column"]
        label_colours = [Color.OrangeRed, Color.SeaGreen, Color.CornflowerBlue, Color.DarkGray]
        added_combo_boxes = []
        for i in range(len(label_texts)):
            label = self.add_label(250, 20, 200, 650 + i*30)
            label.Text = label_texts[i]
            label.BackColor = ControlPaint.LightLight(label_colours[i])
            label.ForeColor = ControlPaint.Dark(label_colours[i])
           
            combo = self.add_combo_box(150, 20, 20, 650 + i*30)
            combo.BackColor = label.BackColor
            combo.ForeColor = ControlPaint.Dark(label_colours[i])
            combo.Enabled = False
            added_combo_boxes.append(combo)
        self.combo_boxes = added_combo_boxes
    
    def add_combo_box(self, width, height, x, y):
        combo_box = ComboBox()
        combo_box.DropDownStyle = ComboBoxStyle.DropDownList
        combo_box.Text = "None"
        combo_box.Width = width
        combo_box.Height = height
        combo_box.Location = Point(x, y)
        self.Controls.Add(combo_box)
        return combo_box
        
    def after_excel_preview(self):
        displayed_row_num = ["None"] + [str(x+1) for x in range(self.grid_excel.Rows.Count)]
        displayed_column_num = ["Optional"] + [self.base_transformer(alpha) for alpha in range(self.grid_excel.Columns.Count)]
        for combo in self.combo_boxes[:3]:
            combo.Enabled = True
            self.entry_combo_box(combo, displayed_row_num)
            combo.SelectedIndex = 0
            combo.SelectedIndexChanged += self.update_combo_selection
        
        column_combo = self.combo_boxes[3]
        column_combo.Enabled = True
        self.entry_combo_box(column_combo, displayed_column_num)
        column_combo.SelectedIndex = 0
        column_combo.SelectedIndexChanged += self.update_combo_selection
        #enable proceed
        self.button_proceed.Enabled = True
        
    def entry_combo_box (self, combo_box, list):
        combo_box.Items.Clear()
        for item in list:
           combo_box.Items.Add(item)
           
    def add_label(self, width, height, x, y):
        label = Label()
        label.Font = System.Drawing.Font(label.Font.FontFamily, 9.5, System.Drawing.FontStyle.Regular)
        label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        label.Width = width
        label.Height = height
        label.Location = Point(x, y)
        self.Controls.Add(label)
        return label
    
    def add_button(self, width, height, x, y):
        button = Button()
        button.Width = width
        button.Height = height
        button.Location = Point(x, y)
        self.Controls.Add(button)
        return button
    
    def add_grid(self, width, height, x, y):
        data_grid_view = DataGridView()
        data_grid_view.ReadOnly = True
        data_grid_view.AllowUserToAddRows = False
        data_grid_view.DefaultCellStyle.SelectionBackColor = Color.Transparent
        data_grid_view.DefaultCellStyle.SelectionForeColor = Color.Transparent
        data_grid_view.BorderStyle = BorderStyle.Fixed3D
        data_grid_view.Width = width
        data_grid_view.Height = height
        data_grid_view.Location = Point(x, y)
        data_grid_view.RowHeadersVisible = True
        data_grid_view.BackgroundColor = Color.White
        self.Controls.Add(data_grid_view)
        return data_grid_view
    
    def disable_things(self):
        for item in self.combo_boxes + [self.button_proceed]:
            item.Enabled = False
    
    def click_select_excel(self, sender, event):
        ofd = OpenFileDialog()
        ofd.Filter = "Excel File|*.xls;*.xlsx;*.xlsm"
        if ofd.ShowDialog() == DialogResult.OK:
            self.disable_things()
            self.grid_excel.Rows.Clear()
            self.grid_excel.Columns.Clear()
            excel_data_rows = self.excel_session(ofd.FileName)
            if excel_data_rows == None:
               return None
            else:
                self.button_select.Text = ofd.FileName
                self.raw_excel_data = excel_data_rows
                self.entry_data_grid(self.grid_excel, excel_data_rows)
                self.after_excel_preview()
        else:
            pass
    
    def entry_data_grid(self, data_grid, list_of_rows):
        column_count = len(list_of_rows[0])
        row_count = len(list_of_rows)
        #add columns and define header
        for i in range(column_count):
            data_grid.Columns.Add(str(i), self.base_transformer(i).upper())
        #add rows
        for i in range(row_count):
            row_as_array = System.Array[System.String]([x for x in list_of_rows[i]])
            data_grid.Rows.Add(row_as_array)
            data_grid.Rows[i].HeaderCell.Value = str(i+1)
        #disable sorting
        for column in data_grid.Columns:
            column.SortMode =  DataGridViewColumnSortMode.NotSortable
        #resize
        data_grid.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)
            
    def excel_session(self, path):
        excelApp = Excel.ApplicationClass()
        excelApp.Visible = False
        excelApp.DisplayAlerts = False
        wBooks = excelApp.Workbooks
        wBook = wBooks.Open(path)
        #output
        data_out = None
        #select sheet
        sheet_names = []
        for sheet in wBook.Sheets:
            sheet_names.append(sheet.Name)
        selected_sheet_form = form_options(sheet_names)
        selected_sheet_form.Text = "Select 1 Work Sheet"
        selected_sheet_form.show()
        if selected_sheet_form.DialogResult == DialogResult.OK:
            selected_sheet_index = selected_sheet_form.selected_index
            if selected_sheet_index != None:
                #extract data
                wSheet = wBook.Sheets[sheet_names[int(selected_sheet_index)]]
                used_range = wSheet.UsedRange
                text_data_rows = []
                for row in used_range.Rows:
                    text_data_row = []
                    for cell in row.Cells:
                        text_data_row.append(cell.Text)
                    text_data_rows.append(text_data_row)
                    data_out = text_data_rows
        else:
             self.button_select.Text = "Select an Excel File"
        #close app
        wBook.Close(True)
        excelApp.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(wBooks)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
        return data_out
        
    def base_transformer(self, column_index):
        num_int = column_index + 1
        atoz = "abcdefghijklmnopqrstuvwxyz"
        alpha = []
        while num_int > 0:
            mod = num_int % 26
            if mod == 0:
                mod = 26
            num_int = int((num_int - mod) / 26)
            alpha.append(atoz[mod - 1])
        alpha.reverse()
        return ''.join(alpha).upper()

    def show(self):
        self.ShowDialog()

# ----- Dynamo Input -----
if IN[0] == False:
    sys.exit("Operation Aborted")

excel_parameters = []
form_start = form_interface("Select Excel Files to Read")
if form_start.DialogResult == DialogResult.OK:
    excel_parameters = form_start.excel_parameters

OUT = excel_parameters
