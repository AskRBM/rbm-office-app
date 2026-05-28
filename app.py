import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime, date
from io import BytesIO

st.set_page_config(page_title="RBM AI Office", page_icon="🏢", layout="wide")

DATA_FOLDER = Path(r"C:\Users\Admin\OneDrive\Apps\RBM_DATA_App")
DATA_FOLDER.mkdir(parents=True, exist_ok=True)

FILES = {
    "settings": DATA_FOLDER / "settings.csv",
    "users": DATA_FOLDER / "users.csv",
    "employees": DATA_FOLDER / "employees.csv",
    "attendance": DATA_FOLDER / "attendance.csv",
    "inout": DATA_FOLDER / "inout.csv",
    "visitors": DATA_FOLDER / "visitors.csv",
    "tasks": DATA_FOLDER / "tasks.csv",
}

COLUMNS = {
    "settings": ["Setting", "Value"],
    "users": ["Username", "Password", "Role", "Full Name"],
    "employees": ["Employee ID", "Employee Name", "Mobile", "Email", "Department", "Designation", "Status"],
    "attendance": ["Date", "Employee Name", "Status", "In Time", "Out Time", "Working Hours", "Remarks", "Created By"],
    "inout": ["Date", "Person Name", "Purpose", "In Time", "Out Time", "Remarks", "Created By"],
    "visitors": ["Date", "Visitor Name", "Mobile", "Company", "Meeting With", "Purpose", "In Time", "Out Time", "Remarks", "Created By"],
    "tasks": ["Date", "Task", "Assigned To", "Priority", "Due Date", "Status", "Remarks", "Created By"],
}

st.markdown("""
<style>
#MainMenu {visibility:hidden;}
footer {visibility:hidden;}
header {visibility:hidden;}
.block-container {padding-top:1.5rem;padding-bottom:2rem;}
.rbm-header {
    background:linear-gradient(135deg,#001f54,#003566);
    padding:24px 28px;border-radius:18px;margin-bottom:25px;
    box-shadow:0px 8px 25px rgba(0,31,84,0.25);
    display:flex;align-items:center;gap:16px;flex-wrap:wrap;
}
.rbm-title {color:white;font-size:38px;font-weight:900;margin:0;line-height:1;}
.rbm-divider {color:white;font-size:34px;font-weight:300;}
.rbm-subtitle {color:white;font-size:17px;font-weight:500;}
.rbm-client {background:#002855;color:#38bdf8;font-size:15px;font-weight:600;padding:8px 12px;border-radius:4px;}
.metric-card {background:white;padding:22px;border-radius:18px;box-shadow:0px 6px 18px rgba(0,0,0,0.08);border:1px solid #e5e7eb;text-align:center;}
.metric-number {font-size:34px;font-weight:800;color:#001f54;}
.metric-label {color:#64748b;font-size:15px;}
.stButton button, .stDownloadButton button {border-radius:12px;font-weight:700;}
@media only screen and (max-width:768px){
.rbm-title{font-size:30px}.rbm-divider{font-size:26px}.rbm-subtitle{font-size:15px}.rbm-client{font-size:13px}
}
</style>
""", unsafe_allow_html=True)


def load_csv(key):
    file_path = FILES[key]
    columns = COLUMNS[key]
    if file_path.exists():
        df = pd.read_csv(file_path)
        for col in columns:
            if col not in df.columns:
                df[col] = ""
        return df[columns]
    df = pd.DataFrame(columns=columns)
    df.to_csv(file_path, index=False)
    return df


def save_csv(key, df):
    df.to_csv(FILES[key], index=False)


def init_users():
    df = load_csv("users")
    if df.empty:
        df = pd.DataFrame([
            {"Username": "admin", "Password": "rbm123", "Role": "Admin", "Full Name": "RBM Admin"},
            {"Username": "user", "Password": "user123", "Role": "User", "Full Name": "RBM User"},
        ])
        save_csv("users", df)
    return df


def init_settings():
    df = load_csv("settings")
    if df.empty:
        df = pd.DataFrame([
            {"Setting": "APP_NAME", "Value": "RBM AI Office Management App"},
            {"Setting": "CLIENT_NAME", "Value": "RBM Client"},
        ])
        save_csv("settings", df)
    return df


def get_setting(name, default):
    df = init_settings()
    row = df[df["Setting"] == name]
    if row.empty:
        return default
    return str(row.iloc[0]["Value"])


def update_setting(name, value):
    df = init_settings()
    if name in df["Setting"].astype(str).tolist():
        df.loc[df["Setting"] == name, "Value"] = value
    else:
        df = pd.concat([df, pd.DataFrame([{"Setting": name, "Value": value}])], ignore_index=True)
    save_csv("settings", df)


APP_NAME = get_setting("APP_NAME", "RBM AI Office Management App")
CLIENT_NAME = get_setting("CLIENT_NAME", "RBM Client")


def rbm_header():
    st.markdown(f"""
    <div class="rbm-header">
        <div class="rbm-title">RBM AI</div>
        <div class="rbm-divider">|</div>
        <div class="rbm-subtitle">Robotic Business Management</div>
        <div class="rbm-client">{APP_NAME} | {CLIENT_NAME}</div>
    </div>
    """, unsafe_allow_html=True)


def show_metric_card(label, value):
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-number">{value}</div>
        <div class="metric-label">{label}</div>
    </div>
    """, unsafe_allow_html=True)


def filter_dataframe(df, keyword):
    if keyword.strip() == "":
        return df
    keyword = keyword.lower()
    mask = df.astype(str).apply(lambda row: row.str.lower().str.contains(keyword, na=False).any(), axis=1)
    return df[mask]


def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Report")
    return output.getvalue()


def calculate_hours(in_time, out_time):
    try:
        t1 = datetime.strptime(str(in_time), "%H:%M:%S")
        t2 = datetime.strptime(str(out_time), "%H:%M:%S")
        return round((t2 - t1).seconds / 3600, 2)
    except Exception:
        return 0


def next_employee_id(df):
    if df.empty:
        return "EMP001"
    nums = []
    for x in df["Employee ID"].dropna().astype(str):
        if x.upper().startswith("EMP"):
            n = x.upper().replace("EMP", "")
            if n.isdigit():
                nums.append(int(n))
    return f"EMP{max(nums) + 1:03d}" if nums else "EMP001"


def show_table_with_edit_delete(key, df, title):
    st.subheader(title)

    search = st.text_input(f"Search {title}", key=f"search_{key}")
    filtered = filter_dataframe(df, search)

    st.dataframe(filtered, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "Download Excel",
            data=to_excel_bytes(filtered),
            file_name=f"{key}_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"xlsx_{key}"
        )
    with c2:
        st.download_button(
            "Download CSV",
            data=filtered.to_csv(index=False).encode("utf-8"),
            file_name=f"{key}_report.csv",
            mime="text/csv",
            use_container_width=True,
            key=f"csv_{key}"
        )

    if st.session_state.get("role") == "Admin" and not df.empty:
        st.divider()
        st.subheader("Admin Edit / Delete")

        row_no = st.number_input(
            "Enter row number to edit/delete",
            min_value=0,
            max_value=max(len(df) - 1, 0),
            value=0,
            step=1,
            key=f"row_{key}"
        )

        st.caption("Row number is the left-side index shown in the table.")

        with st.expander("Edit Selected Row"):
            edited_values = {}
            for col in df.columns:
                edited_values[col] = st.text_input(
                    col,
                    value=str(df.loc[row_no, col]) if row_no < len(df) else "",
                    key=f"edit_{key}_{col}"
                )

            if st.button("Update Selected Row", use_container_width=True, key=f"update_{key}"):
                for col in df.columns:
                    df.loc[row_no, col] = edited_values[col]
                save_csv(key, df)
                st.success("Record updated successfully")
                st.rerun()

        with st.expander("Delete Selected Row"):
            st.warning("This will permanently delete selected row from OneDrive CSV.")
            if st.button("Delete Selected Row", use_container_width=True, key=f"delete_{key}"):
                df = df.drop(index=row_no).reset_index(drop=True)
                save_csv(key, df)
                st.success("Record deleted successfully")
                st.rerun()


def login_page():
    rbm_header()
    st.subheader("Secure Login")
    users = init_users()

    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")

        if st.button("Login", use_container_width=True):
            match = users[
                (users["Username"].astype(str) == username) &
                (users["Password"].astype(str) == password)
            ]
            if not match.empty:
                st.session_state["logged_in"] = True
                st.session_state["username"] = username
                st.session_state["role"] = match.iloc[0]["Role"]
                st.session_state["full_name"] = match.iloc[0]["Full Name"]
                st.rerun()
            else:
                st.error("Wrong username or password")

        st.info("Default Admin: admin / rbm123")


def dashboard():
    st.header("Admin Dashboard")

    emp = load_csv("employees")
    att = load_csv("attendance")
    visitors = load_csv("visitors")
    tasks = load_csv("tasks")

    pending_tasks = 0
    if not tasks.empty:
        pending_tasks = len(tasks[tasks["Status"].astype(str) != "Completed"])

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        show_metric_card("Employees", len(emp))
    with c2:
        show_metric_card("Attendance", len(att))
    with c3:
        show_metric_card("Visitors", len(visitors))
    with c4:
        show_metric_card("Pending Tasks", pending_tasks)

    st.divider()
    st.subheader("Today Summary")

    today_text = str(date.today())
    today_att = att[att["Date"].astype(str) == today_text] if not att.empty else att
    today_visitors = visitors[visitors["Date"].astype(str) == today_text] if not visitors.empty else visitors

    c5, c6 = st.columns(2)
    with c5:
        st.write("Today Attendance")
        st.dataframe(today_att, use_container_width=True)
    with c6:
        st.write("Today Visitors")
        st.dataframe(today_visitors, use_container_width=True)

    st.subheader("Pending Tasks")
    if not tasks.empty:
        st.dataframe(tasks[tasks["Status"].astype(str) != "Completed"], use_container_width=True)
    else:
        st.info("No task data found.")


def employee_master():
    st.header("Employee Master")

    df = load_csv("employees")
    auto_id = next_employee_id(df)

    with st.form("employee_form"):
        c1, c2 = st.columns(2)
        emp_id = c1.text_input("Employee ID", value=auto_id)
        emp_name = c2.text_input("Employee Name")
        mobile = c1.text_input("Mobile")
        email = c2.text_input("Email")
        department = c1.text_input("Department")
        designation = c2.text_input("Designation")
        status = c1.selectbox("Status", ["Active", "Inactive"])

        if st.form_submit_button("Save Employee", use_container_width=True):
            if emp_name.strip() == "":
                st.error("Employee Name is required")
            else:
                new_row = {
                    "Employee ID": emp_id,
                    "Employee Name": emp_name,
                    "Mobile": mobile,
                    "Email": email,
                    "Department": department,
                    "Designation": designation,
                    "Status": status,
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                save_csv("employees", df)
                st.success("Employee saved successfully")
                st.rerun()

    show_table_with_edit_delete("employees", df, "Employee List")


def attendance():
    st.header("Attendance Management")

    df = load_csv("attendance")
    emp = load_csv("employees")

    emp_list = emp["Employee Name"].dropna().tolist()
    if not emp_list:
        emp_list = ["No Employee Found"]

    with st.form("attendance_form"):
        c1, c2 = st.columns(2)
        entry_date = c1.date_input("Date", value=date.today())
        employee = c2.selectbox("Employee Name", emp_list)
        status = c1.selectbox("Status", ["Present", "Absent", "Half Day", "Leave"])
        in_time = c2.time_input("In Time")
        out_time = c1.time_input("Out Time")
        remarks = c2.text_input("Remarks")

        if st.form_submit_button("Save Attendance", use_container_width=True):
            if employee == "No Employee Found":
                st.error("Please create employee first")
            else:
                new_row = {
                    "Date": entry_date,
                    "Employee Name": employee,
                    "Status": status,
                    "In Time": in_time,
                    "Out Time": out_time,
                    "Working Hours": calculate_hours(in_time, out_time),
                    "Remarks": remarks,
                    "Created By": st.session_state["username"],
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                save_csv("attendance", df)
                st.success("Attendance saved successfully")
                st.rerun()

    show_table_with_edit_delete("attendance", df, "Attendance Records")


def inout_register():
    st.header("IN / OUT Register")

    df = load_csv("inout")

    with st.form("inout_form"):
        c1, c2 = st.columns(2)
        entry_date = c1.date_input("Date", value=date.today())
        person = c2.text_input("Person Name")
        purpose = c1.text_input("Purpose")
        in_time = c2.time_input("In Time")
        out_time = c1.time_input("Out Time")
        remarks = c2.text_input("Remarks")

        if st.form_submit_button("Save IN / OUT Entry", use_container_width=True):
            if person.strip() == "":
                st.error("Person Name is required")
            else:
                new_row = {
                    "Date": entry_date,
                    "Person Name": person,
                    "Purpose": purpose,
                    "In Time": in_time,
                    "Out Time": out_time,
                    "Remarks": remarks,
                    "Created By": st.session_state["username"],
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                save_csv("inout", df)
                st.success("IN / OUT entry saved")
                st.rerun()

    show_table_with_edit_delete("inout", df, "IN / OUT Records")


def visitor_register():
    st.header("Visitor Register")

    df = load_csv("visitors")

    with st.form("visitor_form"):
        c1, c2 = st.columns(2)
        entry_date = c1.date_input("Date", value=date.today())
        visitor = c2.text_input("Visitor Name")
        mobile = c1.text_input("Mobile")
        company = c2.text_input("Company")
        meeting_with = c1.text_input("Meeting With")
        purpose = c2.text_input("Purpose")
        in_time = c1.time_input("In Time")
        out_time = c2.time_input("Out Time")
        remarks = c1.text_input("Remarks")

        if st.form_submit_button("Save Visitor", use_container_width=True):
            if visitor.strip() == "":
                st.error("Visitor Name is required")
            else:
                new_row = {
                    "Date": entry_date,
                    "Visitor Name": visitor,
                    "Mobile": mobile,
                    "Company": company,
                    "Meeting With": meeting_with,
                    "Purpose": purpose,
                    "In Time": in_time,
                    "Out Time": out_time,
                    "Remarks": remarks,
                    "Created By": st.session_state["username"],
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                save_csv("visitors", df)
                st.success("Visitor saved successfully")
                st.rerun()

    show_table_with_edit_delete("visitors", df, "Visitor Records")


def task_delegation():
    st.header("Task Delegation")

    df = load_csv("tasks")
    emp = load_csv("employees")
    emp_list = emp["Employee Name"].dropna().tolist()
    if not emp_list:
        emp_list = ["Manual Entry"]

    with st.form("task_form"):
        c1, c2 = st.columns(2)
        task_date = c1.date_input("Task Date", value=date.today())
        task = c2.text_area("Task")
        assign_mode = c1.selectbox("Assign Type", ["Select Employee", "Manual Entry"])

        if assign_mode == "Select Employee" and emp_list != ["Manual Entry"]:
            assigned_to = c2.selectbox("Assigned To", emp_list)
        else:
            assigned_to = c2.text_input("Assigned To")

        priority = c1.selectbox("Priority", ["Low", "Medium", "High", "Urgent"])
        due_date = c2.date_input("Due Date")
        status = c1.selectbox("Status", ["Pending", "In Progress", "Completed"])
        remarks = c2.text_input("Remarks")

        if st.form_submit_button("Save Task", use_container_width=True):
            if task.strip() == "":
                st.error("Task is required")
            else:
                new_row = {
                    "Date": task_date,
                    "Task": task,
                    "Assigned To": assigned_to,
                    "Priority": priority,
                    "Due Date": due_date,
                    "Status": status,
                    "Remarks": remarks,
                    "Created By": st.session_state["username"],
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                save_csv("tasks", df)
                st.success("Task saved successfully")
                st.rerun()

    show_table_with_edit_delete("tasks", df, "Task Records")


def export_reports():
    st.header("Excel / CSV Export Reports")

    report_options = ["users", "employees", "attendance", "inout", "visitors", "tasks"]
    report = st.selectbox("Select Report", report_options)

    if st.session_state["role"] != "Admin" and report == "users":
        st.warning("Only Admin can download user report.")
        return

    df = load_csv(report)
    search = st.text_input("Search Report")
    filtered = filter_dataframe(df, search)

    st.dataframe(filtered, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "Download CSV",
            data=filtered.to_csv(index=False).encode("utf-8"),
            file_name=f"{report}_report.csv",
            mime="text/csv",
            use_container_width=True
        )
    with c2:
        st.download_button(
            "Download Excel",
            data=to_excel_bytes(filtered),
            file_name=f"{report}_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )


def user_management():
    st.header("User Management")

    if st.session_state["role"] != "Admin":
        st.warning("Only Admin can access this page.")
        return

    df = init_users()

    with st.form("user_form"):
        c1, c2 = st.columns(2)
        username = c1.text_input("New Username")
        password = c2.text_input("Password")
        role = c1.selectbox("Role", ["Admin", "User"])
        full_name = c2.text_input("Full Name")

        if st.form_submit_button("Create User", use_container_width=True):
            if username.strip() == "" or password.strip() == "":
                st.error("Username and password are required")
            elif username in df["Username"].astype(str).tolist():
                st.error("Username already exists")
            else:
                new_row = {
                    "Username": username,
                    "Password": password,
                    "Role": role,
                    "Full Name": full_name,
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                save_csv("users", df)
                st.success("User created successfully")
                st.rerun()

    show_table_with_edit_delete("users", df, "User List")


def app_settings():
    st.header("App / Client Settings")

    if st.session_state["role"] != "Admin":
        st.warning("Only Admin can access this page.")
        return

    current_app = get_setting("APP_NAME", "RBM AI Office Management App")
    current_client = get_setting("CLIENT_NAME", "RBM Client")

    with st.form("settings_form"):
        app_name = st.text_input("App Name", value=current_app)
        client_name = st.text_input("Client Name", value=current_client)

        if st.form_submit_button("Save Settings", use_container_width=True):
            update_setting("APP_NAME", app_name)
            update_setting("CLIENT_NAME", client_name)
            st.success("Settings saved. Please refresh app.")
            st.rerun()

    st.info("For different clients, copy the project and change Client Name + DATA_FOLDER path.")


def main_app():
    rbm_header()

    st.sidebar.title("RBM AI")
    st.sidebar.write(f"Client: {CLIENT_NAME}")
    st.sidebar.write(f"User: {st.session_state['full_name']}")
    st.sidebar.write(f"Role: {st.session_state['role']}")

    if st.sidebar.button("Logout", use_container_width=True):
        st.session_state.clear()
        st.rerun()

    if st.session_state["role"] == "Admin":
        menu = [
            "Admin Dashboard",
            "Employee Master",
            "Attendance Management",
            "IN / OUT Register",
            "Visitor Register",
            "Task Delegation",
            "Excel Export Reports",
            "User Management",
            "App Settings",
        ]
    else:
        menu = [
            "Attendance Management",
            "IN / OUT Register",
            "Visitor Register",
            "Task Delegation",
            "Excel Export Reports",
        ]

    choice = st.sidebar.radio("Select Module", menu)

    if choice == "Admin Dashboard":
        dashboard()
    elif choice == "Employee Master":
        employee_master()
    elif choice == "Attendance Management":
        attendance()
    elif choice == "IN / OUT Register":
        inout_register()
    elif choice == "Visitor Register":
        visitor_register()
    elif choice == "Task Delegation":
        task_delegation()
    elif choice == "Excel Export Reports":
        export_reports()
    elif choice == "User Management":
        user_management()
    elif choice == "App Settings":
        app_settings()


if "logged_in" not in st.session_state:
    login_page()
else:
    main_app()