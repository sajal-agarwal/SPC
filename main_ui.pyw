import logging
from main import *
from tkinter import *
from tkinterdnd2 import *
from tkinter import filedialog
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import os
from threading import Thread
import json
from tkinter import messagebox
import tkinter.font as tk_font

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
cur_profile = {}
profile_dir = ''
generation_aborted = False
clear_msg_on_tab_change = False
summary_window = None
rsf_index = 0  # Right Side Frame index
app_help = {}
global do_not_show_summary_var
cur_settings = {}
rule_err = ''
rem_inc_if_err = ''

app_data_path = os.path.join(os.getenv('APPDATA'), 'SPC')  # School Performance Calculator
if not os.path.exists(app_data_path):
    os.makedirs(app_data_path)
default_profile_file = os.path.join(app_data_path, 'default_profile.json')
settings_file = os.path.join(app_data_path, 'settings.json')

help_file_name = './help/help.json'
try:
    with open(help_file_name) as f:
        app_help = json.load(f)
except Exception as err:
    logging.info(err)

try:
    with open(default_profile_file, 'r') as f:
        cur_profile = json.load(f)
except Exception as err:
    logging.error(err)

try:
    with open(settings_file, 'r') as f:
        cur_settings = json.load(f)
except Exception as err:
    logging.error(err)


def load_paths_from_default_profile():
    global in_file_last_dir
    global out_file_name
    global cur_profile
    in_file_last_dir = cur_profile.get('in_file_last_dir', '')
    out_file_name = cur_profile.get('out_file_name', '')


def load_selection_info_from_default_profile():
    global cur_profile

    set_average_cols_str(cur_profile.get('avg_sel_cache', []))
    update_avg_sel_view()

    set_deleted_cols_str(cur_profile.get('rem_sel_cache', []))
    update_rem_sel_view()

    set_sheet_cols_str(cur_profile.get('sheet_sel_cache', []))
    update_sheet_sel_view()

    update_highlight_column_sel_view_from_cur_profile()


def load_rem_inc_if_from_default_profile():
    global cur_profile

    set_remove_if_str(cur_profile.get('remove_if_str', ''))
    set_include_if_str(cur_profile.get('include_if_str', ''))

    str4 = get_remove_if_str()
    if len(str4) > 0:
        scroll_txt2.configure(state=NORMAL)
        scroll_txt2.insert(END, str4)
        scroll_txt3.delete('1.0', END)
        scroll_txt3.configure(state=DISABLED)
    else:
        str5 = get_include_if_str()
        if len(str5) > 0:
            scroll_txt3.configure(state=NORMAL)
            scroll_txt3.insert(END, str5)
            scroll_txt2.delete('1.0', END)
            scroll_txt2.configure(state=DISABLED)


def load_rules_from_profile():
    global cur_profile

    rs = cur_profile.get('rules', '')
    scroll_txt4.delete('1.0', END)
    scroll_txt4.insert(END, rs)


def load_profile_from_file(file_name):
    global cur_profile
    try:
        with open(file_name) as file:
            cur_profile = json.load(file)
    except Exception as e:
        messagebox.showerror('Load Profile', str(e))
        return False

    load_paths_from_default_profile()
    load_selection_info_from_default_profile()
    load_rem_inc_if_from_default_profile()
    load_rules_from_profile()

    return True


def append_input_file_names(text):
    listbox_infile.insert(END, text)


def summary_ok_clicked():
    trigger_generation()
    global summary_window
    summary_window.destroy()


def get_summary():
    tbf_normal = tk_font.nametofont('TkTextFont')
    tbf_heading = tk_font.Font(family=tbf_normal.cget("family"), size=tbf_normal.cget("size") + 2, weight=tk_font.BOLD)

    global summary_window
    summary_window = Toplevel(window)
    summary_window.title("Confirmation")
    summary_window.geometry("600x600")
    summary_window.grab_set()

    btn_frame = Frame(summary_window)
    btn_frame.pack(fill=X, side=BOTTOM, pady=16)
    btn = Button(btn_frame, text='OK', fg='blue', command=summary_ok_clicked)
    btn.pack(side=RIGHT, padx=16, ipadx=10)
    btn = Button(btn_frame, text='Cancel', fg='blue', command=lambda: summary_window.destroy())
    btn.pack(side=RIGHT, ipadx=10)
    do_not_show_cb = Checkbutton(btn_frame, text='Do not show again', fg='blue', variable=do_not_show_summary_var)
    do_not_show_cb.pack(side=LEFT)
    st = ScrolledText(summary_window)
    st.pack(expand=1, fill=BOTH, side=BOTTOM)
    st.tag_config('heading', font=tbf_heading)
    st.insert(END, 'Input files:', 'heading')

    for file in in_filenames:
        st.insert(END, '\n' + file)

    st.insert(END, '\n\nColumns to be removed:', 'heading')
    for col in get_deleted_cols():
        st.insert(END, '\n' + col)

    st.insert(END, '\n\nColumns for which average needs to be calculated:', 'heading')
    for col in get_avg_columns():
        st.insert(END, '\n' + col)

    st.insert(END, '\n\nConditions for row removal:\n', 'heading')
    st.insert(END, get_remove_if_str())

    st.insert(END, '\n\nConditions for row inclusion:\n', 'heading')
    st.insert(END, get_include_if_str())

    st.insert(END, '\n\nRules to be applied:', 'heading')
    if len(rule_err) == 0:
        st.insert(END, '\n' + scroll_txt4.get(0.1, END).strip())

    st.insert(END, '\n\nColumns for dividing sheets:', 'heading')
    for col in get_sheet_columns():
        st.insert(END, '\n' + col)


def enable_all():
    btn1.config(state=NORMAL)
    cur_tab = tabControl.index(tabControl.select())
    btn3.config(state=(DISABLED if (cur_tab == 1 or cur_tab == 2 or cur_tab == 6) else NORMAL))
    btn4.config(state=(NORMAL if (cur_tab == 2) else DISABLED))
    btn5.config(state=(NORMAL if (cur_tab == 0) and (len(listbox_infile.curselection()) > 0) else DISABLED))
    scroll_txt2.config(state=NORMAL if (len(scroll_txt3.get(0.1, END).strip()) == 0) and
                                       (not generation_in_progress) else DISABLED)
    scroll_txt3.config(state=NORMAL if (len(scroll_txt2.get(0.1, END).strip()) == 0) and
                                       (not generation_in_progress) else DISABLED)
    scroll_txt4.config(state=NORMAL)


def disable_all():
    btn1.config(state=DISABLED)
    btn3.config(state=DISABLED)
    btn4.config(state=DISABLED)
    btn5.config(state=DISABLED)
    scroll_txt2.config(state=DISABLED)
    scroll_txt3.config(state=DISABLED)
    scroll_txt4.config(state=DISABLED)


def trigger_generation():
    global generation_in_progress
    generation_in_progress = True

    pb.place(x=75, y=bottom_bar_y + 1, width=win_width - 115, height=24)
    status.place(x=(win_width - 40), y=bottom_bar_y, width=30, height=24)
    status.config(fg='black')
    status_text.set('')
    btn2['text'] = 'Stop'
    disable_all()

    global worker_thread
    set_progress(0)
    worker_thread = Thread(target=do_work, args=(in_filenames, out_file_name))
    worker_thread.start()

    window.after(progress_bar_update_interval, update_progress_fun)


def generate_out_excel():
    global generation_in_progress
    if generation_in_progress:
        btn2.configure(state=DISABLED)
        status.config(fg='black')
        status_text.set('Aborting...')
        pb.place(x=75, y=52, width=0, height=0)
        status.place(x=75, y=bottom_bar_y, width=375, height=24)
        set_exit_flag(True)
        global generation_aborted
        generation_aborted = True
        return

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

    dir_name1 = os.path.dirname(out_file_name)
    if not os.path.exists(dir_name1):
        status.config(fg='red')
        status_text.set('Output file directory does not exist!')
        return

    if not os.access(dir_name1, os.W_OK | os.X_OK):
        status.config(fg='red')
        status_text.set('No permission!')
        return

    set_deleted_cols(listbox.curselection())
    set_average_cols(listbox2.curselection())
    set_sheet_columns(listbox3.curselection())
    set_sheet_highlight_columns(highlight_column_listbox.curselection())
    update_remove_if_cond()
    update_include_if_cond()

    err_msg = ''
    if len(rule_err) > 0:
        err_msg += 'Invalid rules - ' + rule_err + '\n\n'

    if len(rem_inc_if_err) > 0:
        err_msg += 'Invalid ' + ('\'Remove' if len(scroll_txt2.get(0.1, END).strip()) > 0 else '\'Include') + \
                   ' If\' conditions - ' + rem_inc_if_err

    if len(err_msg) > 0:
        messagebox.showerror('Generate', err_msg)
        return

    if do_not_show_summary_var.get():
        trigger_generation()
    else:
        get_summary()


def save_to_default_profile():
    save_profile_to_file(default_profile_file)


def save_profile_to_file(file_name):
    try:
        with open(file_name, 'w') as file:
            cur_profile['in_file_last_dir'] = in_file_last_dir
            cur_profile['out_file_name'] = out_file_name
            cur_profile['avg_sel_cache'] = get_avg_columns()
            cur_profile['rem_sel_cache'] = get_deleted_cols()
            cur_profile['sheet_sel_cache'] = get_sheet_columns()
            cur_profile['remove_if_str'] = get_remove_if_str()
            cur_profile['include_if_str'] = get_include_if_str()
            cur_profile['rules'] = scroll_txt4.get(0.1, END).strip()
            cur_profile['highlight_columns'] = get_sheet_highlight_columns()

            json.dump(cur_profile, file, indent=4)
    except Exception as e:
        messagebox.showerror('Save Profile', str(e))


def update_progress_fun():
    p = get_progress()
    pb['value'] = p
    global generation_aborted
    if not generation_aborted:
        status_text.set(str(int(p)) + '%')

    if p < 100:
        window.after(progress_bar_update_interval, update_progress_fun)
    else:
        global generation_in_progress
        generation_in_progress = False

        pb.stop()
        pb.place(x=75, y=52, width=0, height=0)
        status.place(x=75, y=bottom_bar_y, width=375, height=24)
        err2 = get_last_error()
        if err2 == '':
            status.config(fg='green')
            if generation_aborted:
                btn2.configure(state=NORMAL)
                status_text.set('Aborted successfully!')
                os.remove(out_file_name)
                generation_aborted = False
                set_exit_flag(False)
            else:
                status_text.set('Completed successfully!')
                save_to_default_profile()
                os.startfile(out_file_name)
        else:
            status.config(fg='red')
            status_text.set(err2)

        btn2['text'] = 'Generate'
        enable_all()

        update_df(in_filenames)


def stop_indeterminate_pb():
    pb.stop()
    pb['mode'] = 'determinate'
    pb.place(x=75, y=52, width=0, height=0)
    status.place(x=75, y=bottom_bar_y, width=375, height=24)


def update_columns():
    if get_df_updated():
        global generation_in_progress
        generation_in_progress = False

        err1 = get_last_error()
        if err1 == '':
            index = 1
            cols = get_columns()
            for col in cols:
                listbox.insert(index, col)
                listbox2.insert(index, col)
                listbox3.insert(index, col)
                highlight_column_listbox.insert(index, col)
                index += 1

            stop_indeterminate_pb()

            update_rules()
            load_selection_info_from_default_profile()
            update_remove_if_cond()
            update_include_if_cond()
            update_preview()
        else:
            status.config(fg='red')
            status_text.set(err1)

        btn2.configure(state=NORMAL)
        enable_all()
    else:
        window.after(progress_bar_update_interval, update_columns)


def start_indeterminate_pb():
    pb['mode'] = 'indeterminate'
    pb.place(x=75, y=bottom_bar_y + 1, width=win_width - 85, height=24)
    status.place(x=win_width, y=bottom_bar_y, width=0, height=0)
    status.config(fg='black')
    status_text.set('')
    pb.start()


def trigger_update_columns():
    listbox.delete(0, END)
    listbox2.delete(0, END)
    listbox3.delete(0, END)
    highlight_column_listbox.delete(0, END)

    btn2.configure(state=DISABLED)
    disable_all()

    start_indeterminate_pb()

    global generation_in_progress
    generation_in_progress = True

    global worker_thread2
    worker_thread2 = Thread(target=update_df, args=(in_filenames,))
    worker_thread2.start()

    window.after(progress_bar_update_interval, update_columns)


def load_in_files(filenames):
    if len(filenames) == 0:
        return

    some_files_already_added = False
    no_new_file_added = True
    global in_filenames
    for filename in filenames:
        if filename not in in_filenames:
            no_new_file_added = False
            append_input_file_names(filename + '\n')
            in_filenames.append(filename)
        else:
            some_files_already_added = True

    if no_new_file_added:
        status.configure(fg='orange')
        status_text.set('All file(s) are already added')
        return

    if some_files_already_added:
        status.configure(fg='orange')
        status_text.set('Some file(s) are already added')

    global in_file_last_dir
    in_file_last_dir = os.path.dirname(in_filenames[-1])

    trigger_update_columns()


def browse_in_excel():
    global in_file_last_dir
    filenames = filedialog.askopenfilenames(initialdir=in_file_last_dir,
                                            title="Select a File",
                                            filetypes=(("Excel files",
                                                        "*.xls*"),
                                                       ))

    load_in_files(filenames)


def clear_in_files():
    cur_tab = tabControl.index(tabControl.select())

    if cur_tab == 0:
        in_filenames.clear()
        listbox_infile.delete(0, END)
        listbox.delete(0, END)
        listbox2.delete(0, END)
        listbox3.delete(0, END)
        highlight_column_listbox.delete(0, END)
        clear_preview()
        clear()
        update_rules()
        update_remove_if_cond()
        update_include_if_cond()
    elif cur_tab == 3:
        scroll_txt2.delete('1.0', END)
    elif cur_tab == 4:
        scroll_txt3.delete('1.0', END)
    elif cur_tab == 5:
        scroll_txt4.delete('1.0', END)


def update_avg_sel_view():
    average_cols = get_avg_columns()
    cols = get_columns()
    for col in average_cols:
        if col in cols:
            listbox2.selection_set(cols.index(col))
    on_avg_listbox_selection_changed()


def update_rem_sel_view():
    rem_cols = get_deleted_cols()
    cols = get_columns()
    for col in rem_cols:
        if col in cols:
            listbox.selection_set(cols.index(col))
    on_rem_listbox_selection_changed()


def update_highlight_column_sel_view_from_cur_profile():
    highlight_cols = cur_profile.get('highlight_columns', [])
    cols = get_columns()
    for col in highlight_cols:
        if col in cols:
            highlight_column_listbox.selection_set(cols.index(col))


def update_sheet_sel_view():
    sheet_cols = get_sheet_columns()
    cols = get_columns()
    for col in sheet_cols:
        if col in cols:
            listbox3.selection_set(cols.index(col))


def select_all_numeric_cols_in_list():
    listbox2.selection_clear(0, 'end')
    select_all_numeric_cols()
    average_cols = get_avg_columns()
    cols = get_columns()
    for col in average_cols:
        listbox2.selection_set(cols.index(col))
    on_avg_listbox_selection_changed()


def remove_btn_clicked():
    if_indices = list(listbox_infile.curselection())
    if_indices.sort(reverse=True)
    for index in if_indices:
        in_filenames.pop(index)
        listbox_infile.delete(index)
    btn5.config(state=DISABLED)
    trigger_update_columns()


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
    non_numeric_cols = ''
    non_numeric_col_indexes = []
    count = 0
    for item in sel:
        col_name = get_columns()[item]
        is_numeric = is_column_numeric(col_name)
        if not is_numeric[0]:
            non_numeric_cols += '\n' + col_name + ': ' + str(is_numeric[1])
            non_numeric_col_indexes.append(item)

        if item in sel2:
            clear_items += (', ' if len(clear_items) > 0 else '') + col_name
            count += 1
        listbox.selection_clear(item)

    set_average_cols(listbox2.curselection())

    if len(non_numeric_cols) > 0:
        for item in non_numeric_col_indexes:
            listbox2.selection_clear(item)

        messagebox.showwarning('Average Columns', 'Some selected column' +
                               (' ' if len(non_numeric_col_indexes) == 1 else 's ') +
                               ('is' if len(non_numeric_col_indexes) == 1 else 'are') +
                               ' not numeric so can\'t be selected -' + non_numeric_cols)

    if len(clear_items) != 0:
        status.config(fg='orange')
        status_text.set(str(count) + ' item' + ('s' if (count > 1) else '') +
                        ' removed from Remove Column list: ' + clear_items)

        set_deleted_cols(listbox.curselection())
        update_preview()


def on_infile_listbox_selection_changed():
    cs = listbox_infile.curselection()
    btn5.config(state=NORMAL if len(cs) > 0 else DISABLED)


def on_listbox_selection_changed(evt):
    if evt.widget != listbox_infile:
        return

    on_infile_listbox_selection_changed()


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

    set_deleted_cols(listbox.curselection())

    if len(clear_items) != 0:
        status.config(fg='orange')
        status_text.set(str(count) + ' item' + ('s' if (count > 1) else '') +
                        ' removed from Average Column list: ' + clear_items)

        set_average_cols(listbox2.curselection())

    update_preview()


def on_listbox_selection_changed1(evt):
    if evt.widget != listbox:
        return

    if generation_in_progress:
        revert_selection_change(listbox, get_deleted_cols())
        messagebox.showinfo('Remove Columns', 'A task is currently running, wait for it to finish or stop it.')
        return

    on_rem_listbox_selection_changed()


def revert_selection_change(lb, prev):
    cur = lb.curselection()
    prev = [get_columns().index(i) for i in prev]
    for i in cur:
        if i not in prev:
            lb.selection_clear(i)

    for i in prev:
        if i not in cur:
            lb.selection_set(i)


def on_listbox_selection_changed2(evt):
    if evt.widget != listbox2:
        return

    if generation_in_progress:
        revert_selection_change(listbox2, get_avg_columns())
        messagebox.showinfo('Average Columns', 'A task is currently running, wait for it to finish or stop it.')
        return

    on_avg_listbox_selection_changed()


def on_sheet_listbox_selection_changed():
    set_sheet_columns(listbox3.curselection())


def on_listbox_selection_changed3(evt):
    if evt.widget != listbox3:
        return

    if generation_in_progress:
        revert_selection_change(listbox3, get_sheet_columns())
        messagebox.showinfo('Sheet Columns', 'A task is currently running, wait for it to finish or stop it.')
        return

    on_sheet_listbox_selection_changed()


def update_help():
    if rsf_index & 1:
        tab_name = tabControl.tab(tabControl.select(), 'text')
        scroll_txt5.config(state=NORMAL)
        scroll_txt5.delete('1.0', END)
        scroll_txt5.insert(END, app_help[tab_name])
        scroll_txt5.config(state=DISABLED)


def clear_preview():
    for row in preview.get_children():
        preview.delete(row)


def update_preview():
    if rsf_index & 2:
        clear_preview()
        if not generation_in_progress:
            apply_rules()
            do_remove_if()
            do_include_if()
            delete_columns()
        preview['columns'] = get_df().columns.tolist()
        anchor_pos = 'w'
        preview.column("#0", width=0, stretch=NO)
        preview.heading("#0", text="", anchor=anchor_pos)
        for col in get_df().columns:
            preview.column(col, anchor=anchor_pos, width=60)
            preview.heading(col, text=col, anchor=anchor_pos)

        index = 0
        while index < len(get_df()):
            preview.insert(parent='', index='end', iid=str(index), text='', values=get_df().iloc[index].to_list())
            index += 1

        apply_rules()


def tab_changed(evt):
    if evt.widget != tabControl:
        return

    cur_tab = tabControl.index(tabControl.select())

    global clear_msg_on_tab_change
    if clear_msg_on_tab_change:
        clear_msg_on_tab_change = False
        status.config(fg='black')
        status_text.set('')

    if cur_tab == 3:
        if len(scroll_txt3.get(0.1, END).strip()) == 0:
            if not generation_in_progress:
                scroll_txt2.configure(state=NORMAL)
                if len(rem_inc_if_err) > 0:
                    status.config(fg='red')
                    status_text.set(rem_inc_if_err)
        else:
            scroll_txt2.configure(state=DISABLED)
            status.config(fg='orange')
            status_text.set('Remove all the conditions from "Include If" tab to add condition(s) here!')
            clear_msg_on_tab_change = True
    elif cur_tab == 4:
        if len(scroll_txt2.get(0.1, END).strip()) == 0:
            if not generation_in_progress:
                scroll_txt3.config(state=NORMAL)
                if len(rem_inc_if_err) > 0:
                    status.config(fg='red')
                    status_text.set(rem_inc_if_err)
        else:
            scroll_txt3.configure(state=DISABLED)
            status.config(fg='orange')
            status_text.set('Remove all the conditions from "Remove If" tab to add condition(s) here!')
            clear_msg_on_tab_change = True
    elif cur_tab == 5:
        if len(rule_err) > 0:
            status.config(fg='red')
            status_text.set(rule_err)

    btn3.config(state=(DISABLED if (cur_tab == 1 or cur_tab == 2 or cur_tab == 6
                                    or generation_in_progress) else NORMAL))
    btn4.config(state=(NORMAL if ((cur_tab == 2) and not generation_in_progress) else DISABLED))

    btn5.config(state=(NORMAL if (cur_tab == 0) and (len(listbox_infile.curselection()) > 0)
                and (not generation_in_progress) else DISABLED))

    update_help()


def validate_rem_inc_if_str(rem_str):
    empty = True
    success = True
    err_msg = ''
    if len(rem_str) > 0:
        conditions = rem_str.split('\n')
        for cond in conditions:
            if cond == '':
                continue

            empty = False

            if '==' in cond:
                cond_list = cond.split('==')
                if len(cond_list) < 2:
                    success = False
                    err_msg = 'Right hand value missing'
                else:
                    if cond_list[0] not in get_columns():
                        success = False
                        err_msg = 'Column name \'' + cond_list[0] + '\' not found'
                        break

                    if not is_val_exits_in_column(cond_list[0], cond_list[1]):
                        success = False
                        err_msg = 'Column value \'' + cond_list[1] + '\' not found'
                        break
            else:
                success = False
                err_msg = '== operator is missing'

    if success:
        status.configure(fg='black')
        status_text.set('')
    else:
        status.configure(fg='red')
        status_text.set(err_msg)

    global rem_inc_if_err
    rem_inc_if_err = err_msg

    return (not empty) and success


def update_remove_if_cond():
    rem_str = scroll_txt2.get(0., END)
    if validate_rem_inc_if_str(rem_str):
        set_remove_if_str(rem_str)
    else:
        set_remove_if_str('')


def remove_if_text_changed(evt):
    if evt.widget != scroll_txt2:
        return

    old_rem_str = get_remove_if_str()

    update_remove_if_cond()

    if old_rem_str != get_remove_if_str():
        update_preview()

    scroll_txt2.edit_modified(False)  # reset to detect next change


def update_include_if_cond():
    inc_str = scroll_txt3.get(0., END)
    if validate_rem_inc_if_str(inc_str):
        set_include_if_str(inc_str)
    else:
        set_include_if_str('')


def include_if_text_changed(evt):
    if evt.widget != scroll_txt3:
        return

    old_inc_str = get_include_if_str()

    update_include_if_cond()

    if old_inc_str != get_include_if_str():
        update_preview()

    scroll_txt3.edit_modified(False)  # reset to detect next change


def update_rules():
    success = True
    err_msg = ''
    dict1 = {}
    rules_str = scroll_txt4.get(0., END).strip()
    if len(rules_str) > 0:
        rules_list = rules_str.split('\n')

        for col in get_columns():
            dict1[col] = {}

        for rule in rules_list:
            kv = rule.split('=')
            if len(kv) > 1:
                # v = ''
                cols = []
                has_not = False
                if ':' in kv[1]:
                    val_cols = kv[1].split(':')

                    v = val_cols[0]

                    cols_str = val_cols[1].strip()
                    if len(cols_str) > 0:
                        cols = cols_str.split(',')
                        if '!' in cols_str:
                            has_not = True
                            for col in cols:
                                if col[0] != '!':
                                    err_msg = '! is not applied to all'
                                    success = False
                                    break
                else:
                    v = kv[1]

                if not success:
                    break

                if len(cols) == 0:
                    cols = get_columns()

                if has_not:
                    cols2 = []
                    for col in cols:
                        if is_int(col[1:]):
                            col_num = int(col[1:])
                            if col_num < 1:
                                err_msg = 'Column number must be greater than 0'
                                success = False
                                break

                            col_index = col_num - 1
                            if col_index >= len(get_columns()):
                                err_msg = 'Column number ' + str(col_index) + ' not found'
                                success = False
                                break
                            cols2.append(get_columns()[col_index])
                            continue

                        if col[1:] not in get_columns():
                            err_msg = 'Column name \'' + col[1:] + '\' not found'
                            success = False
                            break
                        else:
                            cols2.append(col[1:])

                    for col in get_columns():
                        if col in cols2:
                            continue
                        dict1[col][kv[0]] = v
                else:
                    for col in cols:
                        if is_int(col):
                            col_num = int(col)
                            if col_num < 1:
                                err_msg = 'Column number must be greater than 0'
                                success = False
                                break

                            col_index = col_num - 1
                            if col_index >= len(get_columns()):
                                err_msg = 'Column number ' + str(col_num) + ' not found'
                                success = False
                                break
                            col = get_columns()[col_index]
                        if col in get_columns():
                            dict1[col][kv[0]] = v
                        else:
                            err_msg = 'Column name \'' + col + '\' not found'
                            success = False
                            break
            else:
                err_msg = '= is missing'
                success = False
                break

    old_rules = get_rules()

    if success:
        set_rules(dict1)
        status.config(fg='black')
        status_text.set('')
    else:
        set_rules({})
        status.config(fg='red')
        status_text.set(err_msg)

    global rule_err
    rule_err = err_msg

    if (old_rules != get_rules()) and get_df_updated():
        apply_rules()
        update_preview()

    scroll_txt4.edit_modified(False)

    return success


def rules_text_changed(evt):
    global scroll_txt4
    if evt.widget != scroll_txt4:
        return

    update_rules()


def settings_btn_clicked():
    if generation_in_progress:
        messagebox.showinfo('Settings', 'A task is currently running, wait for it to finish or stop it.')
        return

    try:
        popup_menu.tk_popup(settings_btn.winfo_rootx(), settings_btn.winfo_rooty() + 28)
    finally:
        popup_menu.grab_release()


def save_profile_to_user_file():
    file_name = filedialog.asksaveasfilename(initialfile='Untitled.json', defaultextension='.json',
                                             filetypes=[('Json File', '*.json')])

    if len(file_name) > 0:
        save_profile_to_file(file_name)


def load_profile_from_user_file():
    file_name = filedialog.askopenfilename(initialdir=profile_dir, title="Select a File",
                                           filetypes=(("Json File", '*.json'),))

    if len(file_name) > 0:
        load_profile_from_file(file_name)


def files_dropped(evt):
    if generation_in_progress:
        messagebox.showinfo('File drop', 'Previous task is still running, stop it to add more files.')
        return

    xl_files = []
    files = window.tk.splitlist(evt.data)
    for file in files:
        if file.lower().endswith(".xls") or file.lower().endswith(".xlsx"):
            xl_files.append(file)

    if len(xl_files) > 0:
        load_in_files(xl_files)

    for file in files:
        if file.lower().endswith(".json"):
            load_profile_from_file(file)


def update_center_view():
    if rsf_index == 1:
        lbl1['text'] = 'Help'
        tc_width = (win_width - 20) * 2.0 / 3
        lbl1.place(x=tc_width + 10, y=40, width=50, height=22)
        tabControl.place(x=10, y=40, width=tc_width, height=win_height - 140)
        scroll_txt5.place(x=tc_width + 10, y=40 + 22, width=(win_width - tc_width - 20), height=win_height - 140 - 22)
        preview_frame.place(width=0, height=0)
    elif rsf_index == 2:
        lbl1['text'] = 'Preview'
        tc_width = (win_width - 20) / 3.0
        lbl1.place(x=tc_width + 10, y=40, width=50, height=22)
        tabControl.place(x=10, y=40, width=tc_width, height=win_height - 140)
        preview_frame.place(x=tc_width + 10, y=40 + 22, width=(win_width - tc_width - 20), height=win_height - 140 - 22)
        scroll_txt5.place(width=0, height=0)
    elif rsf_index == 3:
        lbl1['text'] = 'Help & Preview'
        tc_width = (win_width - 20) / 3.0
        help_height = (win_height - 140 - 22) / 2
        lbl1.place(x=tc_width + 10, y=40, width=100, height=22)
        tabControl.place(x=10, y=40, width=tc_width, height=win_height - 140)
        scroll_txt5.place(x=tc_width + 10, y=40 + 22, width=(win_width - tc_width - 20), height=help_height)
        preview_frame.place(x=tc_width + 10, y=(help_height + 40 + 22),
                            width=(win_width - tc_width - 20), height=help_height)
    else:
        lbl1.place(width=0, height=0)
        tabControl.place(x=10, y=40, width=win_width - 20, height=win_height - 140)
        scroll_txt5.place(width=0, height=0)
        preview_frame.place(width=0, height=0)


def help_button_clicked():
    global rsf_index
    rsf_index = (1 if (help_var.get() == 1) else 0) | (2 if (preview_var.get() == 1) else 0)
    update_center_view()
    update_help()


def preview_button_clicked():
    global rsf_index
    rsf_index = (1 if (help_var.get() == 1) else 0) | (2 if (preview_var.get() == 1) else 0)
    update_center_view()
    update_preview()


def about_app():
    messagebox.showinfo("SPC", '''This is a School Performance Calculator application \
it takes school performance data in the form of Excel file and generates another Excel \
file with well organized data in different sheets and plots.

Contact shipragupta89@gmail.com for more details.''')


def on_closing():
    try:
        settings_dict = {
            'do_not_show_summary_var': do_not_show_summary_var.get(),
        }

        if window.state() != 'zoomed':
            settings_dict['win_width'] = win_width
            settings_dict['win_height'] = win_height

        with open(settings_file, 'w') as file:
            json.dump(settings_dict, file, indent=4)
    except Exception as err1:
        logging.error(err1)

    window.destroy()


def top_window_resized(event):
    global win_width
    global win_height
    if event.widget == window and (win_width != event.width or win_height != event.height):
        win_width = event.width
        win_height = event.height
        update_center_view()
        scroll_txt2.pack(expand=1, fill="both")
        scroll_txt3.pack(expand=1, fill="both")
        scroll_txt4.pack(expand=1, fill="both")
        lbl2.place(x=10, y=event.height - 80, width=70)
        out_file_edit.place(x=90, y=event.height - 80, width=event.width - 100)
        global bottom_bar_y
        bottom_bar_y = event.height - 50
        btn2.place(x=10, y=bottom_bar_y, width=60)
        if generation_in_progress:
            pb.place(x=75, y=bottom_bar_y + 1, width=win_width - 125, height=24)
            status.place(x=(win_width - 40), y=bottom_bar_y, width=30, height=24)
        else:
            status.place(x=75, y=bottom_bar_y, width=(win_width - 95), height=24)
        settings_btn.place(x=event.width - 40, y=10, width=30, height=26)
        btn1.place(x=event.width - 95, y=10, width=50)
        btn3.place(x=event.width - 150, y=10, width=50)
        btn4.place(x=event.width - 265, y=10, width=110)
        btn5.place(x=event.width - 330, y=10, width=60)
        checkbox_help.place(x=10, y=10, width=50)
        checkbox_preview.place(x=70, y=10, width=70)

        load_rules_from_profile()


load_paths_from_default_profile()


treeview_popup_row = ''
treeview_popup_col = ''


def add_remove_if_filter():
    global treeview_popup_row
    global treeview_popup_col
    treeview_popup_item = preview.item(preview.get_children()[int(treeview_popup_row)])['values'][
        int(treeview_popup_col.lstrip('#')) - 1]
    treeview_popup_col_name = preview.column(int(treeview_popup_col.lstrip('#')) - 1)['id']
    rem_if_filter_str = treeview_popup_col_name + '==' + str(treeview_popup_item)
    cur_rem_str = scroll_txt2.get(0., END).strip()
    cur_rem_str = cur_rem_str + ('' if len(cur_rem_str) == 0 else '\n') + rem_if_filter_str
    scroll_txt2.delete('1.0', END)
    scroll_txt2.insert(END, cur_rem_str)


def add_include_if_filter():
    global treeview_popup_row
    global treeview_popup_col
    treeview_popup_item = preview.item(preview.get_children()[int(treeview_popup_row)])['values'][
        int(treeview_popup_col.lstrip('#')) - 1]
    treeview_popup_col_name = preview.column(int(treeview_popup_col.lstrip('#')) - 1)['id']
    rem_if_filter_str = treeview_popup_col_name + '==' + str(treeview_popup_item)
    cur_rem_str = scroll_txt3.get(0., END).strip()
    cur_rem_str = cur_rem_str + ('' if len(cur_rem_str) == 0 else '\n') + rem_if_filter_str
    scroll_txt3.delete('1.0', END)
    scroll_txt3.insert(END, cur_rem_str)


def treeview_popup(event):
    global treeview_popup_row
    global treeview_popup_col
    treeview_popup_row = preview.identify_row(event.y)
    treeview_popup_col = preview.identify_column(event.x)
    if treeview_popup_row and treeview_popup_col:
        context_menu = Menu(window, tearoff=0)
        context_menu.add_command(label="Add Remove Filter", command=add_remove_if_filter)
        context_menu.add_command(label="Add Include Filter", command=add_include_if_filter)
        context_menu.post(event.x_root, event.y_root)


window = TkinterDnD.Tk()
window.drop_target_register(DND_FILES)

do_not_show_summary_var = BooleanVar()
do_not_show_summary_var.set(cur_settings.get('do_not_show_summary_var', False))

window.bind("<Configure>", top_window_resized)
window.dnd_bind('<<Drop>>', files_dropped)
window.protocol("WM_DELETE_WINDOW", on_closing)

settings_photo = PhotoImage(file=r".\icons\settings.png")
settings_btn = Button(window, text="Settings", fg='blue', image=settings_photo, command=settings_btn_clicked)
popup_menu = Menu(window, tearoff=0)
popup_menu.add_command(label="Save Profile", command=save_profile_to_user_file)
popup_menu.add_command(label="Load Profile", command=load_profile_from_user_file)
popup_menu.add_separator()
popup_menu.add_checkbutton(label='Do not show summary window', variable=do_not_show_summary_var)
popup_menu.add_separator()
popup_menu.add_command(label="About", command=about_app)

btn1 = Button(window, text="Add", fg='blue', command=browse_in_excel)
btn3 = Button(window, text="Clear", fg='blue', command=clear_in_files)
btn4 = Button(window, text="Select All Numeric", fg='blue', command=select_all_numeric_cols_in_list)
btn5 = Button(window, text="Remove", fg='blue', command=remove_btn_clicked)

help_var = IntVar()
checkbox_help = Checkbutton(window, text="Help", fg='blue', variable=help_var, command=help_button_clicked)

preview_var = IntVar()
checkbox_preview = Checkbutton(window, text="Preview", fg='blue', variable=preview_var, command=preview_button_clicked)

tabControl = ttk.Notebook(window)
tabControl.bind('<<NotebookTabChanged>>', tab_changed)
tab1 = ttk.Frame(tabControl)
tab2 = ttk.Frame(tabControl)
tab3 = ttk.Frame(tabControl)
tab4 = ttk.Frame(tabControl)
tab5 = ttk.Frame(tabControl)
tab6 = ttk.Frame(tabControl)
tab7 = ttk.Frame(tabControl)
tab8 = ttk.Frame(tabControl)
tabControl.add(tab1, text='Input Excel')
tabControl.add(tab2, text='Remove Columns')
tabControl.add(tab3, text='Average Columns')
tabControl.add(tab4, text='Remove If')
tabControl.add(tab5, text='Include If')
tabControl.add(tab6, text='Rules')
tabControl.add(tab7, text='Sheets Columns')
tabControl.add(tab8, text='Highlight Columns')

scrollbar_infile = Scrollbar(tab1)
scrollbar_infile.pack(side=RIGHT, fill=Y)
listbox_infile = Listbox(tab1, selectmode=MULTIPLE, exportselection=0, yscrollcommand=scrollbar_infile.set)
listbox_infile.pack(side=LEFT, expand=1, fill="both")
listbox_infile.bind('<<ListboxSelect>>', on_listbox_selection_changed)
scrollbar_infile.config(command=listbox_infile.yview)

scrollbar = Scrollbar(tab2)
scrollbar.pack(side=RIGHT, fill=Y)
listbox = Listbox(tab2, selectmode=MULTIPLE, exportselection=0, yscrollcommand=scrollbar.set)
listbox.pack(side=LEFT, expand=1, fill="both")
listbox.bind('<<ListboxSelect>>', on_listbox_selection_changed1)
scrollbar.config(command=listbox.yview)

scrollbar2 = Scrollbar(tab3)
scrollbar2.pack(side=RIGHT, fill=Y)
listbox2 = Listbox(tab3, selectmode=MULTIPLE, exportselection=0, yscrollcommand=scrollbar2.set)
listbox2.pack(side=LEFT, expand=1, fill="both")
listbox2.bind('<<ListboxSelect>>', on_listbox_selection_changed2)
scrollbar2.config(command=listbox2.yview)

scroll_txt2 = ScrolledText(tab4, wrap="none", undo=True, autoseparators=True, maxundo=-1)
scroll_txt2.bind("<<Modified>>", remove_if_text_changed)

scroll_txt3 = ScrolledText(tab5, wrap="none", undo=True, autoseparators=True, maxundo=-1)
scroll_txt3.bind("<<Modified>>", include_if_text_changed)

scroll_txt4 = ScrolledText(tab6, wrap="none", undo=True, autoseparators=True, maxundo=-1)
scroll_txt4.bind("<<Modified>>", rules_text_changed)

lbl1 = Label(window, text="Help", anchor='w')
scroll_txt5 = ScrolledText(window, bg='#f0f0f0', border=0, wrap=WORD)
scroll_txt5.config(state=DISABLED)

preview_frame = Frame(window)
preview_vsb = Scrollbar(preview_frame)
preview_vsb.pack(side=RIGHT, fill=Y)
preview_hsb = Scrollbar(preview_frame, orient='horizontal')
preview_hsb.pack(side=BOTTOM, fill=X)
preview = ttk.Treeview(preview_frame, yscrollcommand=preview_vsb.set, xscrollcommand=preview_hsb.set)
preview.bind("<Button-3>", treeview_popup)
preview.pack(fill=BOTH, expand=True)
preview_vsb.config(command=preview.yview)
preview_hsb.config(command=preview.xview)

scrollbar3 = Scrollbar(tab7)
scrollbar3.pack(side=RIGHT, fill=Y)
listbox3 = Listbox(tab7, selectmode=MULTIPLE, exportselection=0, yscrollcommand=scrollbar2.set)
listbox3.pack(side=LEFT, expand=1, fill="both")
listbox3.bind('<<ListboxSelect>>', on_listbox_selection_changed3)
scrollbar3.config(command=listbox3.yview)

highlight_col_scrollbar = Scrollbar(tab8)
highlight_col_scrollbar.pack(side=RIGHT, fill=Y)
highlight_column_listbox = Listbox(tab8, selectmode=MULTIPLE, exportselection=0, yscrollcommand=highlight_col_scrollbar.set)
highlight_column_listbox.pack(side=LEFT, expand=1, fill="both")
highlight_col_scrollbar.config(command=highlight_column_listbox.yview)

lbl2 = Label(window, text="Output Excel", anchor='w')
out_file_text = StringVar()
if len(out_file_name) > 0:
    if os.path.exists(out_file_name):
        out_file_text.set(out_file_name)
    else:
        dir_name = os.path.dirname(out_file_name)
        if os.path.exists(dir_name):
            out_file_text.set(os.path.join(dir_name, 'out.xlsx'))

if len(out_file_text.get().strip()) == 0:
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    if os.path.isdir(desktop):
        out_file_text.set(os.path.join(desktop, 'out.xlsx'))

out_file_edit = Entry(window, textvariable=out_file_text)

btn2 = Button(window, text="Generate", fg='blue', command=generate_out_excel)
pb = ttk.Progressbar(
    window,
    orient='horizontal',
    mode='determinate',
    length=0
)
status_text = StringVar()
status = Label(window, text="", textvariable=status_text, anchor='w')

window.title('SPC')
window_width = cur_settings.get('win_width', 1000)
window_height = cur_settings.get('win_height', 600)
window.geometry("{}x{}".format(window_width, window_height))
window.mainloop()
set_exit_flag(True)

if worker_thread != 0:
    worker_thread.join()

if worker_thread2 != 0:
    worker_thread2.join()
