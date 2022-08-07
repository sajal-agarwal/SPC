import logging
import tkinter as tk
from main import *
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import os
from threading import Thread

bottom_bar_y = 150
in_filenames = []
in_file_last_dir = ''
win_width = 0
win_height = 0
worker_thread = 0  # generate excel
worker_thread2 = 0  # update df
generation_in_progress = False
progress_bar_update_interval = 500
status_display_time = 3000
status_cancel_id = -1
out_file_name = ''

app_data_path = os.path.join(os.getenv('APPDATA'), 'SPC')  # School Performance Calculator
if not os.path.exists(app_data_path):
    os.makedirs(app_data_path)
file_path_cache = os.path.join(app_data_path, 'paths.txt')
avg_sel_cache = os.path.join(app_data_path, 'avg_sel.txt')
rem_sel_cache = os.path.join(app_data_path, 'rem_sel.txt')


def load_cache():
    global file_path_cache
    global in_file_last_dir
    global out_file_name
    try:
        with open(file_path_cache, 'r') as file1:
            in_file_last_dir = file1.readline()
            in_file_last_dir = in_file_last_dir.strip('\n')
            out_file_name = file1.readline()
    except Exception as err1:
        logging.error(err1)


def load_sel_cache():
    try:
        with open(avg_sel_cache, 'r') as file1:
            items_str = file1.read()
            items = items_str.split('-+-')
            set_average_cols_str(items)
            update_avg_sel_view()

    except Exception as err1:
        logging.error(err1)

    try:
        with open(rem_sel_cache, 'r') as file1:
            items_str = file1.read()
            items = items_str.split('-+-')
            set_deleted_cols_str(items)
            update_rem_sel_view()
    except Exception as err1:
        logging.error(err1)


def set_input_file_names(text):
    scroll_txt.configure(state='normal')
    scroll_txt.insert(END, text)
    scroll_txt.configure(state='disabled')


def generate_out_excel():
    if len(in_filenames) == 0:
        status.config(fg='red')
        status_text.set('Add input file(s)!')
        return

    global out_file_name
    out_file_name = out_file_text.get()
    out_file_name = out_file_name.strip('"')
    if len(out_file_name) == 0:
        status.config(fg='red')
        status_text.set('Provide an output file name!')
        return

    dir_name = os.path.dirname(out_file_name)
    if not os.path.exists(dir_name):
        status.config(fg='red')
        status_text.set('Output file directory does not exist!')
        return

    if not os.access(dir_name, os.W_OK | os.X_OK):
        status.config(fg='red')
        status_text.set('No permission!')
        return

    pb.place(x=70, y=bottom_bar_y + 1, width=win_width - 110, height=24)
    status.place(x=(win_width - 40), y=bottom_bar_y, width=30, height=24)
    status_text.set('')
    status.config(fg='black')
    btn2.configure(state='disabled')

    set_deleted_cols(listbox.curselection())
    set_average_cols(listbox2.curselection())
    set_remove_if_str(scroll_txt2.get(0.1, END))
    set_include_if_str(scroll_txt3.get(0.1, END))

    global generation_in_progress
    generation_in_progress = True
    global worker_thread
    set_progress(0)
    worker_thread = Thread(target=do_work, args=(in_filenames, out_file_name))
    worker_thread.start()

    window.after(progress_bar_update_interval, update_progress_fun)


def update_cache():
    try:
        with open(file_path_cache, 'w') as f:
            global in_file_last_dir
            f.write(in_file_last_dir + '\n')
            f.write(out_file_name)
    except Exception as err:
        logging.error(err)

    try:
        with open(avg_sel_cache, 'w') as f:
            cols = ''
            for col in get_avg_columns():
                cols += ('' if len(cols) == 0 else '-+-') + col
            f.write(cols)
    except Exception as err:
        logging.error(err)

    try:
        with open(rem_sel_cache, 'w') as f:
            cols = ''
            for col in get_deleted_cols():
                cols += ('' if len(cols) == 0 else '-+-') + col
            f.write(cols)
    except Exception as err:
        logging.error(err)


def update_progress_fun():
    p = get_progress()
    pb['value'] = p
    status_text.set(str(int(p)) + '%')
    if p < 100:
        window.after(progress_bar_update_interval, update_progress_fun)
    else:
        pb.stop()
        pb.place(x=70, y=52, width=0, height=0)
        status.place(x=70, y=bottom_bar_y, width=380, height=24)
        err = get_last_error()
        if err == '':
            status.config(fg='green')
            status_text.set('Completed!')
            update_cache()
        else:
            status.config(fg='red')
            status_text.set(err)

        btn2.configure(state='normal')
        global generation_in_progress
        generation_in_progress = False


def update_columns():
    if get_df_updated():
        err = get_last_error()
        if err == '':
            index = 1
            cols = get_columns()
            for col in cols:
                listbox.insert(index, col)
                listbox2.insert(index, col)
                index += 1
            select_all_numeric_cols()
        else:
            status.config(fg='red')
            status_text.set(err)

        btn2.configure(state='normal')
        load_sel_cache()
    else:
        window.after(progress_bar_update_interval, update_columns)


def browse_in_excel():
    global in_file_last_dir
    filenames = filedialog.askopenfilenames(initialdir=in_file_last_dir,
                                            title="Select a File",
                                            filetypes=(("Excel files",
                                                        "*.xls*"),
                                                       ))

    if len(filenames) > 0:
        status_text.set('')
    else:
        return

    global in_filenames
    entry_text.set(filenames)
    for filename in filenames:
        set_input_file_names(filename + '\n')
        in_filenames.append(filename)

    in_file_last_dir = os.path.dirname(in_filenames[-1])

    listbox.delete(0, END)
    listbox2.delete(0, END)

    btn2.configure(state='disabled')

    global worker_thread2
    worker_thread2 = Thread(target=update_df, args=(in_filenames,))
    worker_thread2.start()

    window.after(progress_bar_update_interval, update_columns)


def clear_in_files():
    in_filenames.clear()
    scroll_txt.configure(state='normal')
    scroll_txt.delete('1.0', END)
    scroll_txt.configure(state='disabled')
    listbox.delete(0, END)
    listbox2.delete(0, END)
    clear()


def top_window_resized(event):
    global win_width
    global win_height
    if event.widget == window and (win_width != event.width or win_height != event.height):
        win_width = event.width
        win_height = event.height
        tabControl.place(x=10, y=40, width=event.width - 20, height=event.height - 140)
        scroll_txt.pack(expand=1, fill="both")
        listbox.pack(side=LEFT, expand=1, fill="both")
        listbox2.pack(side=LEFT, expand=1, fill="both")
        scroll_txt2.pack(expand=1, fill="both")
        scroll_txt3.pack(expand=1, fill="both")
        lbl2.place(x=10, y=event.height - 80)
        out_file_edit.place(x=100, y=event.height - 80, width=event.width - 110)
        global bottom_bar_y
        bottom_bar_y = event.height - 50
        btn2.place(x=10, y=bottom_bar_y)
        if generation_in_progress:
            pb.place(x=70, y=bottom_bar_y + 1, width=win_width - 110, height=24)
            status.place(x=(win_width - 40), y=bottom_bar_y, width=30, height=24)
        else:
            status.place(x=70, y=bottom_bar_y, width=(win_width - 80), height=24)
        btn1.place(x=event.width - 60, y=10, width=50)
        btn3.place(x=event.width - 120, y=10, width=50)
        btn4.place(x=event.width - 240, y=10, width=110)


def update_avg_sel_view():
    average_cols = get_avg_columns()
    cols = get_columns()
    for col in average_cols:
        index = cols.index(col)
        listbox2.selection_set(index)
    on_avg_listbox_selection_changed()


def update_rem_sel_view():
    rem_cols = get_deleted_cols()
    cols = get_columns()
    for col in rem_cols:
        listbox.selection_set(cols.index(col))
    on_rem_listbox_selection_changed()


def select_all_numeric_cols_in_list():
    listbox2.selection_clear(0, 'end')
    select_all_numeric_cols()
    average_cols = get_avg_columns()
    cols = get_columns()
    for col in average_cols:
        listbox2.selection_set(cols.index(col))
    on_avg_listbox_selection_changed()


def clear_status():
    global status_cancel_id
    status_cancel_id = -1
    status.config(fg='black')
    status_text.set('')


def set_temp_status(txt, col):
    status.config(fg=col)
    status_text.set(txt)
    global status_cancel_id
    if status_cancel_id != -1:
        window.after_cancel(status_cancel_id)
    status_cancel_id = window.after(status_display_time, clear_status)


def on_avg_listbox_selection_changed():
    sel = listbox2.curselection()
    sel2 = listbox.curselection()
    clear_items = ''
    count = 0
    for item in sel:
        if item in sel2:
            clear_items += (', ' if len(clear_items) > 0 else '') + get_columns()[item]
            count += 1
        listbox.selection_clear(item)

    if len(clear_items) != 0:
        status.config(fg='orange')
        status_text.set(str(count) + ' item' + ('s' if (count > 1) else '') +
                        ' removed from Remove Column list: ' + clear_items)


def on_rem_listbox_selection_changed():
    sel = listbox.curselection()
    sel2 = listbox2.curselection()
    clear_items = ''
    count = 0
    for item in sel:
        if item in sel2:
            clear_items += (', ' if len(clear_items) > 0 else '') + get_columns()[item]
            count += 1
        listbox2.selection_clear(item)

    if len(clear_items) != 0:
        status.config(fg='orange')
        status_text.set(str(count) + ' item' + ('s' if (count > 1) else '') +
                        ' removed from Average Column list: ' + clear_items)


def on_listbox_selection_changed1(evt):
    if evt.widget != listbox:
        return

    on_rem_listbox_selection_changed()


def on_listbox_selection_changed2(evt):
    if evt.widget != listbox2:
        return

    on_avg_listbox_selection_changed()


def tab_changed(evt):
    if evt.widget != tabControl:
        return

    cur_tab = tabControl.index(tabControl.select())

    status.config(fg='black')
    status_text.set('')

    if cur_tab == 2:
        btn4.configure(state='normal')
    elif cur_tab == 3:
        if len(scroll_txt3.get(0.1, END).strip()) == 0:
            scroll_txt2.configure(state='normal')
        else:
            scroll_txt2.configure(state='disabled')
            status.config(fg='orange')
            status_text.set('Remove all the conditions from "Include If" tab to add condition(s) here!')
    elif cur_tab == 4:
        if len(scroll_txt2.get(0.1, END).strip()) == 0:
            scroll_txt3.configure(state='normal')
        else:
            scroll_txt3.configure(state='disabled')
            status.config(fg='orange')
            status_text.set('Remove all the conditions from "Remove If" tab to add condition(s) here!')
    else:
        btn4.configure(state='disabled')


def remove_if_text_changed(evt):
    if evt.widget != scroll_txt2:
        return

    scroll_txt2.edit_modified(False)  # reset to detect next change


def include_if_text_changed(evt):
    if evt.widget != scroll_txt3:
        return

    scroll_txt3.edit_modified(False)  # reset to detect next change


load_cache()

window = Tk()

window.bind("<Configure>", top_window_resized)

entry_text = tk.StringVar()

btn1 = Button(window, text="Add", fg='blue', command=browse_in_excel)
btn3 = Button(window, text="Clear", fg='blue', command=clear_in_files)
btn4 = Button(window, text="Select All Numeric", fg='blue', command=select_all_numeric_cols_in_list)

tabControl = ttk.Notebook(window)
tabControl.bind('<<NotebookTabChanged>>', tab_changed)
tab1 = ttk.Frame(tabControl)
tab2 = ttk.Frame(tabControl)
tab3 = ttk.Frame(tabControl)
tab4 = ttk.Frame(tabControl)
tab5 = ttk.Frame(tabControl)
tabControl.add(tab1, text='Input Excel')
tabControl.add(tab2, text='Remove Column')
tabControl.add(tab3, text='Average Column')
tabControl.add(tab4, text='Remove If')
tabControl.add(tab5, text='Include If')

scroll_txt = ScrolledText(tab1, wrap="none")
scroll_txt.config(state=DISABLED)

scrollbar = Scrollbar(tab2)
scrollbar.pack(side=RIGHT, fill=Y)
listbox = Listbox(tab2, selectmode=MULTIPLE, exportselection=0, yscrollcommand=scrollbar.set)
listbox.bind('<<ListboxSelect>>', on_listbox_selection_changed1)
scrollbar.config(command=listbox.yview)

scrollbar2 = Scrollbar(tab3)
scrollbar2.pack(side=RIGHT, fill=Y)
listbox2 = Listbox(tab3, selectmode=MULTIPLE, exportselection=0, yscrollcommand=scrollbar2.set)
listbox2.bind('<<ListboxSelect>>', on_listbox_selection_changed2)
scrollbar2.config(command=listbox2.yview)

scroll_txt2 = ScrolledText(tab4, wrap="none")
scroll_txt2.bind("<<Modified>>", remove_if_text_changed)

scroll_txt3 = ScrolledText(tab5, wrap="none")
scroll_txt3.bind("<<Modified>>", include_if_text_changed)

lbl2 = Label(window, text="Output Excel")
out_file_text = tk.StringVar()
if len(out_file_name) > 0 and os.path.exists(out_file_name):
    out_file_text.set(out_file_name)
else:
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    if os.path.isdir(desktop):
        os.path.join(desktop, 'Result.xlsx')
        out_file_text.set(desktop)
out_file_edit = Entry(window, textvariable=out_file_text)

btn2 = Button(window, text="Generate", fg='blue', command=generate_out_excel)
pb = ttk.Progressbar(
    window,
    orient='horizontal',
    mode='determinate',
    length=0
)
status_text = tk.StringVar()
status = Label(window, text="", textvariable=status_text, anchor='w')

window.title('SPC')
window.geometry("600x400")
window.mainloop()
set_exit_flag(True)

if worker_thread != 0:
    worker_thread.join()

if worker_thread2 != 0:
    worker_thread2.join()
