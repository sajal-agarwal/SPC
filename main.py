import pandas as pd
from openpyxl import load_workbook
import append_toexcel as axl
import threading
import math
import shutil

pd.options.mode.chained_assignment = None
# pd.set_option('display.max_columns', None)
# pd.set_option('display.max_rows', None)

df = pd.DataFrame()
untouched_df = pd.DataFrame()
summary_df = pd.DataFrame()
columns = []
last_err_str = ''
progress = 0
exit_flag = False
df_updated = False
remove_if_str = ''
include_if_str = ''
rules = {}
progress_lock = threading.Lock()
error_lock = threading.Lock()
exit_flag_lock = threading.Lock()
df_update_lock = threading.Lock()


def set_progress(cur_progress):
    with progress_lock:
        global progress
        progress = cur_progress


def get_progress():
    with progress_lock:
        return progress


def set_last_error(cur_err):
    with error_lock:
        global last_err_str
        last_err_str = cur_err


def get_last_error() -> object:
    with error_lock:
        return last_err_str


def set_exit_flag(flag):
    with exit_flag_lock:
        global exit_flag
        exit_flag = flag


def get_exit_flag():
    with exit_flag_lock:
        return exit_flag


def set_df_updated(updated):
    with df_update_lock:
        global df_updated
        df_updated = updated


def get_df_updated():
    with df_update_lock:
        return df_updated


def get_columns():
    return columns


def get_avg_columns():
    return avg_cols


avg_cols = []


def set_average_cols(cols):
    global avg_cols
    avg_cols = [columns[col] for col in cols]


def set_average_cols_str(cols):
    global avg_cols
    avg_cols = cols


del_cols = []


def set_deleted_cols(cols):
    global del_cols
    del_cols = [columns[col] for col in cols]


def set_deleted_cols_str(cols):
    global del_cols
    del_cols = cols


def get_deleted_cols():
    global del_cols
    return del_cols


def get_remove_if_str():
    global remove_if_str
    return remove_if_str


def set_remove_if_str(rem_str):
    global remove_if_str
    remove_if_str = rem_str.strip()


def get_include_if_str():
    global include_if_str
    return include_if_str


def set_include_if_str(add_str):
    global include_if_str
    include_if_str = add_str.strip()


def set_rules(r):
    global rules
    rules = r


def get_rules():
    global rules
    return rules


def clear():
    global df
    global summary_df
    global columns
    global last_err_str
    global progress
    global exit_flag
    global df_updated
    global avg_cols
    global del_cols
    # global remove_if_str
    # global include_if_str

    df = pd.DataFrame()
    summary_df = pd.DataFrame()
    set_last_error('')
    set_progress(0)
    set_exit_flag(False)
    set_df_updated(False)
    columns = []
    avg_cols = []
    del_cols = []
    # remove_if_str = ''
    # include_if_str = ''

# [
#     'Approach of teachers during the PTM.',
#     'Satisfaction levels on responses/replies received from teachers.',
#     'Approach of PROs.',
#     'Responsiveness & approach of the Admin Team',
#     'Overall happiness of the child in School',
#     'a) Academic subjects transaction',
#     'b) Activity classes transaction',
#     'c) Class Teacher’s approach\n',
#     'd) Subject Teacher’s Approach',
#     'e) Written work/ Assignments'
# ]


def is_int(s):
    try:
        int(s)
        return True
    except ValueError:
        return False


def is_float(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


tmp_col_rules = {}


def get_val_from_rule(str1):
    val = tmp_col_rules.get(str1, '')
    if len(val) > 0:
        if is_int(val):
            return int(val)
        elif is_float(val):
            return float(val)
        else:
            return val
    else:
        return str1


def apply_rules_on_column(df1, col_name):
    if len(rules) == 0:
        return

    global tmp_col_rules
    tmp_col_rules = rules[col_name]
    df1[col_name] = df[col_name].apply(get_val_from_rule)


def update_average(cl_df, col_name):
    cl_df.at['Average', col_name] = round(cl_df[col_name].mean(), 2)


def select_all_numeric_cols():
    if get_df_updated():
        global avg_cols
        avg_cols.clear()
        for col in columns:
            is_numeric = True
            count_nan = 0
            for item in df[col]:
                if not isinstance(item, (int, float)):
                    is_numeric = False
                    break
                elif math.isnan(item):
                    count_nan += 1

            if is_numeric and (count_nan < df[col].size):
                avg_cols.append(col)


def apply_rules():
    if get_df_updated() and (len(rules) > 0):
        global df
        df.reset_index(inplace=True, drop=True)
        for col in columns:
            apply_rules_on_column(df, col)


def update_df(in_paths):
    success = True
    set_last_error('')
    set_df_updated(False)
    global df
    global untouched_df
    try:
        untouched_df = pd.DataFrame()
        for in_path in in_paths:
            if get_exit_flag():
                break

            untouched_df = untouched_df.append(pd.read_excel(in_path))

        global columns
        columns = untouched_df.columns.to_list()
    except Exception as e:
        set_last_error(e)
        success = False

    set_df_updated(True)

    df = untouched_df.copy()

    apply_rules()

    return success


def reset_df():
    global df
    global untouched_df
    df = untouched_df.copy()


def delete_columns():
    global del_cols
    for col in del_cols:
        df.drop(col, inplace=True, axis=1)


def create_output_file_from_template(out_file):
    shutil.copyfile('./data/OutputTemplate.xlsx', out_file)


def get_count_df():
    global df
    global avg_cols
    if get_df_updated():
        count_df = pd.DataFrame(columns=avg_cols, index=range(10, 0, -1))
        for col in avg_cols:
            vc = df[col].value_counts()
            count_df[col] = [(vc[i] if i in vc else 0) for i in range(10, 0, -1)]
        return count_df
    return pd.DataFrame()


def do_remove_if():
    global df
    global remove_if_str

    remove_if_str = remove_if_str.strip()

    if len(remove_if_str) == 0:
        return

    conditions = remove_if_str.split('\n')
    try:
        for cond in conditions:
            if '==' in cond:
                cond_list = cond.split('==')
                if len(cond_list) == 0:
                    raise Exception("")
                df = df[[(str(i) != cond_list[1]) for i in df[cond_list[0]]]]
            else:
                raise Exception('')
            # elif '!=' in cond:
            #     cond_list = cond.split('!=')
            #     if len(cond_list) == 0:
            #         continue
            #     new_df.append(df[[(str(i) == cond_list[1]) for i in df[cond_list[0]]]])
            # elif '<' in cond:
            #     cond_list = cond.split('<')
            #     if len(cond_list) == 0:
            #         continue
            #     df = df[df[cond_list[0]] < cond_list[1]]
            # elif '<=' in cond:
            #     cond_list = cond.split('<=')
            #     if len(cond_list) == 0:
            #         continue
            #     df = df[df[cond_list[0]] <= cond_list[1]]
            # elif '>' in cond:
            #     cond_list = cond.split('>')
            #     if len(cond_list) == 0:
            #         continue
            #     df = df[df[cond_list[0]] > cond_list[1]]
            # elif '>=' in cond:
            #     cond_list = cond.split('>=')
            #     if len(cond_list) == 0:
            #         continue
            #     df = df[df[cond_list[0]] >= cond_list[1]]
    except Exception as e:
        raise Exception("Invalid condition - " + ("only == supported with no extra spaces"
                                                  if (len(str(e)) == 0) else str(e)))


def do_include_if():
    global df
    global include_if_str

    include_if_str = include_if_str.strip()

    if len(include_if_str) == 0:
        return

    conditions = include_if_str.split('\n')
    new_df = pd.DataFrame()
    try:
        for cond in conditions:
            if '==' in cond:
                cond_list = cond.split('==')
                if len(cond_list) == 0:
                    raise Exception("")

                new_df = new_df.append(df[[(str(i) == cond_list[1]) for i in df[cond_list[0]]]])
            else:
                raise Exception('')
        df = new_df
    except Exception as e:
        raise Exception("Invalid condition - " + ("only == supported with no extra spaces"
                                                  if (len(str(e)) == 0) else str(e)))


def do_work(in_paths, out_path):
    global df
    global summary_df
    success = True
    try:
        set_last_error('')

        if not get_df_updated():
            update_df(in_paths)

        delete_columns()

        do_remove_if()

        do_include_if()

        create_output_file_from_template(out_path)

        axl.append_df_to_excel(out_path, df, sheet_name='MasterSheet')

        classes = df['Class'] + '-' + df['Section']
        classes.drop_duplicates(inplace=True)
        classes.sort_values(inplace=True)
        classes.reset_index(inplace=True, drop=True)

        is_summary_needed = (len(avg_cols) > 0)

        if is_summary_needed:
            summary_df = pd.DataFrame(columns=avg_cols, index=[str1.replace("-", "") for str1 in classes])

        counter = 0
        for c in classes:
            if get_exit_flag():
                break

            cl_sec = c.split('-')
            cl_df = df[(df['Class'] == cl_sec[0]) & (df['Section'] == cl_sec[1])]
            for col in avg_cols:
                update_average(cl_df, col)

            row_name = (cl_sec[0] + cl_sec[1])
            if is_summary_needed:
                av_ser = cl_df.iloc[-1]
                av_ser = av_ser[~av_ser.isnull()]
                summary_df.loc[row_name] = av_ser

            axl.append_df_to_excel(out_path, cl_df, sheet_name=row_name)

            counter += 1
            set_progress(counter*100/(classes.size + 1))

        count_df = get_count_df()
        if is_summary_needed:
            avg_of_avg = summary_df.mean().round(decimals=2)
            count_df.loc['Weighted Average'] = avg_of_avg
            summary_df.loc['Average'] = avg_of_avg
            axl.append_df_to_excel(out_path, summary_df, sheet_name='Summary')

        axl.append_df_to_excel(out_path, count_df, sheet_name='Counts', startrow=0, startcol=0)
        # book = load_workbook(out_path)
        # writer = pd.ExcelWriter(out_path, engine='openpyxl')
        # writer.book = book
        # writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        # count_df.to_excel(writer, "Counts", startrow=0, startcol=0)
        # writer.save()

        print('Done!!!')
    except Exception as e:
        set_last_error(e)
        success = False

    set_progress(100)
    set_df_updated(False)
    return success
