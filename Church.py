# main.py
import sys
import os
import threading
import socket
import ssl
import smtplib
import csv
from datetime import datetime, date, timedelta

from email.message import EmailMessage

# PyQt5 UI
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QComboBox, QMessageBox,
    QTableWidget, QTableWidgetItem, QDialog, QFileDialog,
    QStackedWidget, QFrame, QSpacerItem, QSizePolicy, QTextEdit,
    QFormLayout, QSpinBox
)
from PyQt5.QtCore import Qt, pyqtSignal, QTimer
from PyQt5.QtGui import QFont

# Flask for phone registration
from flask import Flask, request, render_template_string, send_from_directory

# QR / camera / images
import qrcode
from pyzbar.pyzbar import decode
import cv2

# plotting and exports
import matplotlib.pyplot as plt
import pandas as pd
import sqlite3

# Optional nicer theme
try:
    import qdarkstyle
except Exception:
    qdarkstyle = None

# ---------------------------
# CONFIG
# ---------------------------
DB_FILE = "attendance.db"
SMTP_EMAIL = "nccmultimedia2022@gmail.com"
# Put your Gmail app password here (spaces allowed, code will strip them before use)
SMTP_APP_PASSWORD_RAW = "yuzg ffpq unuv zgmn"

FLASK_PORT = 5000
POSTER_QR_FILENAME = "poster_qr.png"

# ---------------------------
# Helper functions
# ---------------------------
def local_ip():
    """Find local IP for poster link (best-effort)."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def now_iso():
    return datetime.now().isoformat(sep=" ", timespec="seconds")

# ---------------------------
# DATABASE (SQLite)
# ---------------------------
class Database:
    def __init__(self, path=DB_FILE):
        self.conn = sqlite3.connect(path, check_same_thread=False)
        self.conn.execute("PRAGMA foreign_keys = ON")
        self._create_tables()

    def _create_tables(self):
        c = self.conn.cursor()
        # members: include optional contact & facebook
        c.execute("""
        CREATE TABLE IF NOT EXISTS members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT NOT NULL,
            contact TEXT,
            facebook TEXT,
            role TEXT NOT NULL,
            status TEXT DEFAULT 'active',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        )
        """)
        # attendance: one row per (member,date)
        c.execute("""
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            member_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            login_time TEXT,
            logout_time TEXT,
            event TEXT,
            FOREIGN KEY(member_id) REFERENCES members(id) ON DELETE CASCADE,
            UNIQUE(member_id, date)
        )
        """)
        # admins
        c.execute("""
        CREATE TABLE IF NOT EXISTS admins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE,
            password TEXT,
            role TEXT DEFAULT 'admin'
        )
        """)
        # seed default admin
        c.execute("SELECT 1 FROM admins WHERE username='admin'")
        if not c.fetchone():
            c.execute("INSERT INTO admins (username,password,role) VALUES (?,?,?)", ("admin", "1234", "super"))
        self.conn.commit()

    # admin
    def verify_admin(self, username, password):
        c = self.conn.cursor()
        c.execute("SELECT id, username, role FROM admins WHERE username=? AND password=?", (username, password))
        return c.fetchone()

    # members
    def add_member(self, name, email, contact, facebook, role):
        c = self.conn.cursor()
        c.execute("INSERT INTO members (name,email,contact,facebook,role) VALUES (?,?,?,?,?)",
                  (name, email, contact, facebook, role))
        self.conn.commit()
        return c.lastrowid

    def update_member(self, member_id, name, email, contact, facebook, role):
        c = self.conn.cursor()
        c.execute("""
            UPDATE members SET name=?, email=?, contact=?, facebook=?, role=?
            WHERE id=?
        """, (name, email, contact, facebook, role, member_id))
        self.conn.commit()

    def get_member(self, member_id):
        c = self.conn.cursor()
        c.execute("SELECT id,name,email,contact,facebook,role,status FROM members WHERE id=?", (member_id,))
        return c.fetchone()

    def find_member_by_email(self, email):
        c = self.conn.cursor()
        c.execute("SELECT id,name,email,contact,facebook,role,status FROM members WHERE email=?", (email,))
        return c.fetchone()

    def list_members(self, active_only=True):
        c = self.conn.cursor()
        if active_only:
            c.execute("SELECT id,name,email,contact,facebook,role,status FROM members WHERE status='active' ORDER BY name")
        else:
            c.execute("SELECT id,name,email,contact,facebook,role,status FROM members ORDER BY name")
        return c.fetchall()

    def set_member_inactive(self, member_id):
        c = self.conn.cursor()
        c.execute("UPDATE members SET status='inactive' WHERE id=?", (member_id,))
        self.conn.commit()

    # attendance operations
    def _attendance_row(self, member_id, for_date=None):
        if for_date is None:
            for_date = date.today().isoformat()
        c = self.conn.cursor()
        c.execute("SELECT id, login_time, logout_time FROM attendance WHERE member_id=? AND date=?", (member_id, for_date))
        return c.fetchone()

    def ensure_attendance_row(self, member_id, event="General", for_date=None):
        if for_date is None:
            for_date = date.today().isoformat()
        c = self.conn.cursor()
        row = self._attendance_row(member_id, for_date)
        if row:
            return row
        c.execute("INSERT INTO attendance (member_id, date, event) VALUES (?,?,?)", (member_id, for_date, event))
        self.conn.commit()
        return self._attendance_row(member_id, for_date)

    def set_login(self, member_id):
        row = self.ensure_attendance_row(member_id)
        if row[1] is None:
            c = self.conn.cursor()
            c.execute("UPDATE attendance SET login_time=? WHERE id=?", (now_iso(), row[0]))
            self.conn.commit()
            return True, "Login recorded"
        return False, "Already logged in today"

    def set_logout(self, member_id):
        row = self.ensure_attendance_row(member_id)
        if row[1] is None:
            return False, "Cannot logout before login"
        if row[2] is None:
            c = self.conn.cursor()
            c.execute("UPDATE attendance SET logout_time=? WHERE id=?", (now_iso(), row[0]))
            self.conn.commit()
            return True, "Logout recorded"
        return False, "Already logged out today"

    def todays_attendance(self):
        d = date.today().isoformat()
        c = self.conn.cursor()
        c.execute("""
            SELECT a.id, m.id, m.name, m.role, a.login_time, a.logout_time, a.event
            FROM attendance a JOIN members m ON a.member_id=m.id
            WHERE a.date=?
            ORDER BY a.login_time
        """, (d,))
        return c.fetchall()

    def attendance_history(self, member_id=None, start=None, end=None):
        q = """
            SELECT a.id, m.id, m.name, m.role, a.date, a.login_time, a.logout_time, a.event
            FROM attendance a JOIN members m ON a.member_id=m.id
        """
        cond = []
        params = []
        if member_id:
            cond.append("m.id=?"); params.append(member_id)
        if start:
            cond.append("a.date>=?"); params.append(start)
        if end:
            cond.append("a.date<=?"); params.append(end)
        if cond:
            q += " WHERE " + " AND ".join(cond)
        q += " ORDER BY a.date DESC, a.login_time DESC"
        c = self.conn.cursor()
        c.execute(q, tuple(params))
        return c.fetchall()

    def attendance_counts(self):
        c = self.conn.cursor()
        c.execute("""
            SELECT m.name, COUNT(a.id) as cnt
            FROM members m LEFT JOIN attendance a ON m.id=a.member_id
            WHERE m.status='active'
            GROUP BY m.id
            ORDER BY cnt DESC, m.name
        """)
        return c.fetchall()

    def absent_for_weeks(self, weeks=3):
        cutoff = (date.today() - timedelta(weeks=weeks)).isoformat()
        c = self.conn.cursor()
        c.execute("""
            SELECT m.id, m.name, m.email
            FROM members m
            WHERE m.status='active' AND m.id NOT IN (
                SELECT member_id FROM attendance WHERE date >= ?
            )
            ORDER BY m.name
        """, (cutoff,))
        return c.fetchall()

# single DB instance used by Flask and GUI
DB = Database()

# ---------------------------
# QR generation & emailing
# ---------------------------
def generate_qr_image(data, path):
    qr = qrcode.QRCode(version=2, box_size=8, border=3)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(path)

def send_qr_email(to_email, member_id, qr_path):
    # build email
    msg = EmailMessage()
    msg["Subject"] = "Your NCC Attendance QR Code"
    msg["From"] = SMTP_EMAIL
    msg["To"] = to_email
    msg.set_content(f"Hello,\n\nAttached is your NCC Attendance QR code. Member ID: {member_id}\n\nBlessings.")
    try:
        with open(qr_path, "rb") as f:
            data = f.read()
            msg.add_attachment(data, maintype="image", subtype="png", filename=os.path.basename(qr_path))
    except Exception as e:
        return False, f"QR file error: {e}"
    app_pw = SMTP_APP_PASSWORD_RAW.replace(" ", "")
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
            smtp.login(SMTP_EMAIL, app_pw)
            smtp.send_message(msg)
        return True, "Email sent"
    except Exception as e:
        return False, f"Email failed: {e}"

# ---------------------------
# Flask server for phone registration
# ---------------------------
flask_app = Flask(__name__)

FLASK_FORM = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>NCC Registration</title>
  <style>
    body { font-family: Arial, sans-serif; background:#f4f6f8; padding:20px; }
    .card { background:white; padding:20px; border-radius:8px; max-width:480px; margin:auto; box-shadow:0 6px 18px rgba(0,0,0,0.08); }
    input, select { width:100%; padding:10px; margin:8px 0; border-radius:6px; border:1px solid #ccc; }
    button { padding:10px 16px; background:#2d89ef; color:white; border:none; border-radius:6px; }
  </style>
</head>
<body>
  <div class="card">
    <h2>NCC Attendance - Register</h2>
    <p>Register to receive your personal QR code for attendance.</p>
    <form method="post">
      <label>Name (required)</label><input name="name" required>
      <label>Email (required)</label><input name="email" type="email" required>
      <label>Contact (optional)</label><input name="contact" placeholder="09XXXXXXXXX">
      <label>Facebook (optional)</label><input name="facebook" placeholder="facebook.com/username">
      <label>Role</label>
      <select name="role">
        <option>Youth</option><option>Young Pro</option><option>Tanders</option>
      </select>
      <div style="margin-top:12px;"><button type="submit">Register</button></div>
    </form>
  </div>
</body>
</html>
"""

@flask_app.route("/register", methods=["GET", "POST"])
def register_route():
    if request.method == "POST":
        name = request.form.get("name","").strip()
        email = request.form.get("email","").strip()
        contact = request.form.get("contact","").strip()
        facebook = request.form.get("facebook","").strip()
        role = request.form.get("role","Youth")
        if not name or not email:
            return "<p>Name and email required. <a href=''>Back</a></p>", 400
        # create member
        member_id = DB.add_member(name, email, contact, facebook, role)
        qr_path = f"qr_{member_id}.png"
        generate_qr_image(str(member_id), qr_path)
        ok, msg = send_qr_email(email, member_id, qr_path)
        # show simple confirmation + link to QR image
        return f"""
           <p>Registered! Member ID: {member_id}. QR generated.</p>
           <p>{msg}</p>
           <p><img src="/qr/{os.path.basename(qr_path)}" alt="QR"></p>
        """
    return FLASK_FORM

@flask_app.route("/qr/<path:filename>")
def route_qr(filename):
    return send_from_directory(os.getcwd(), filename)

def start_flask_server_background():
    # build & save poster QR linking to /register
    ip = local_ip()
    url = f"http://{ip}:{FLASK_PORT}/register"
    generate_qr_image(url, POSTER_QR_FILENAME)
    print("Poster registration URL:", url)
    print("Poster QR file generated:", POSTER_QR_FILENAME, "(print this and place at the entrance)")
    thread = threading.Thread(target=lambda: flask_app.run(host="0.0.0.0", port=FLASK_PORT, threaded=True), daemon=True)
    thread.start()
    return thread

# ---------------------------
# PyQt GUI (single-window app with pages)
# ---------------------------
class LoginDialog(QDialog):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.setWindowTitle("NCC Attendance - Admin Login")
        self.setFixedSize(360, 220)
        layout = QVBoxLayout(self)
        title = QLabel("NCC Attendance — Admin")
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        self.user = QLineEdit(); self.user.setPlaceholderText("Username")
        self.pwd = QLineEdit(); self.pwd.setPlaceholderText("Password"); self.pwd.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.user); layout.addWidget(self.pwd)
        btn = QPushButton("Sign In")
        btn.clicked.connect(self.attempt_login)
        layout.addWidget(btn)
        layout.addStretch()
        self.result = None

    def attempt_login(self):
        u = self.user.text().strip(); p = self.pwd.text().strip()
        if DB.verify_admin(u, p):
            self.accept()
        else:
            QMessageBox.warning(self, "Login failed", "Invalid username or password.")

# MainWindow has signals so background threads can safely notify GUI
class MainWindow(QMainWindow):
    scanner_message = pyqtSignal(str)    # to show short messages
    refresh_signal = pyqtSignal()        # to trigger table refreshes

    def __init__(self):
        super().__init__()
        self.setWindowTitle("NCC Attendance - Admin")
        self.resize(1100, 700)

        self._scanner_running = False
        self._scanner_thread = None

        root_widget = QWidget()
        root = QHBoxLayout(root_widget)

        # sidebar
        sidebar = QFrame(); sidebar.setFixedWidth(220)
        sl = QVBoxLayout(sidebar)
        title = QLabel("NCC Attendance")
        title.setFont(QFont("Segoe UI", 14, QFont.Bold))
        sl.addWidget(title)
        self.btn_dash = QPushButton("Dashboard"); self.btn_att = QPushButton("Attendance")
        self.btn_mem = QPushButton("Members"); self.btn_rep = QPushButton("Reports")
        for b in (self.btn_dash, self.btn_att, self.btn_mem, self.btn_rep):
            b.setFixedHeight(42); sl.addWidget(b)
        sl.addItem(QSpacerItem(20,20, QSizePolicy.Minimum, QSizePolicy.Expanding))
        root.addWidget(sidebar)

        # pages stack
        self.pages = QStackedWidget()
        self.page_dashboard = DashboardPage(self)
        self.page_attendance = AttendancePage(self)
        self.page_members = MembersPage(self)
        self.page_reports = ReportsPage(self)
        for p in (self.page_dashboard, self.page_attendance, self.page_members, self.page_reports):
            self.pages.addWidget(p)
        root.addWidget(self.pages)
        self.setCentralWidget(root_widget)

        # wire buttons
        self.btn_dash.clicked.connect(lambda: self.pages.setCurrentWidget(self.page_dashboard))
        self.btn_att.clicked.connect(lambda: self.pages.setCurrentWidget(self.page_attendance))
        self.btn_mem.clicked.connect(lambda: self.pages.setCurrentWidget(self.page_members))
        self.btn_rep.clicked.connect(lambda: self.pages.setCurrentWidget(self.page_reports))

        # connect signals
        self.scanner_message.connect(self.show_scanner_message_box)
        self.refresh_signal.connect(self.refresh_all_pages)

        # start an auto-refresh timer for dashboard (every 8s)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.page_dashboard.refresh)
        self.timer.start(8000)

    def show_scanner_message_box(self, text):
        # non-blocking info (use information)
        QMessageBox.information(self, "Scanner", text)

    def refresh_all_pages(self):
        try:
            self.page_dashboard.refresh()
            self.page_attendance.reload_members_combo()
            self.page_members.reload_table()
            self.page_reports.refresh_preview()
        except Exception:
            pass

    # scanning control
    def start_scanner(self):
        if self._scanner_running:
            QMessageBox.information(self, "Scanner", "Scanner already running.")
            return
        self._scanner_running = True
        self._scanner_thread = threading.Thread(target=self._scanner_loop, daemon=True)
        self._scanner_thread.start()
        QMessageBox.information(self, "Scanner", "Scanner started. Webcam window will appear. Press Q in the webcam window to stop manually, or use Stop Scanner button.")

    def stop_scanner(self):
        if not self._scanner_running:
            QMessageBox.information(self, "Scanner", "Scanner is not running.")
            return
        self._scanner_running = False
        QMessageBox.information(self, "Scanner", "Stopping scanner... (webcam window will close shortly)")

    def _scanner_loop(self):
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            self.scanner_message.emit("Cannot open webcam.")
            self._scanner_running = False
            return
        last_seen = {}  # member_id -> timestamp to avoid immediate duplicates
        while self._scanner_running:
            ret, frame = cap.read()
            if not ret:
                break
            decoded = decode(frame)
            if decoded:
                for d in decoded:
                    raw = d.data.decode("utf-8")
                    # expected content: member id integer
                    try:
                        member_id = int(raw)
                    except:
                        member_id = None
                    if member_id:
                        now_ts = datetime.now().timestamp()
                        if member_id in last_seen and now_ts - last_seen[member_id] < 2:
                            # skip if same id scanned within 2 seconds
                            continue
                        last_seen[member_id] = now_ts

                        member = DB.get_member(member_id)
                        if member:
                            # attempt login or logout
                            # ensure attendance row exists
                            row = DB.ensure_attendance_row(member_id)
                            # row: (id, login_time, logout_time)
                            if row[1] is None:
                                ok, msg = DB.set_login(member_id)
                                self.scanner_message.emit(f"{member[1]}: {msg}")
                            elif row[2] is None:
                                ok, msg = DB.set_logout(member_id)
                                self.scanner_message.emit(f"{member[1]}: {msg}")
                            else:
                                self.scanner_message.emit(f"{member[1]}: already logged in/out today")
                            # refresh GUI
                            self.refresh_signal.emit()
                        else:
                            # unknown QR - instruct registration
                            ip = local_ip()
                            url = f"http://{ip}:{FLASK_PORT}/register"
                            self.scanner_message.emit(f"Unknown QR. Ask person to register via {url} (poster QR printed).")
                    else:
                        # not an integer - maybe it's a URL (poster) - ignore
                        pass
            # show webcam window so operator can see scans
            cv2.imshow("NCC Scanner (press Q to stop)", frame)
            if cv2.waitKey(1) & 0xFF == ord('q'):
                # allow operator to stop scanner with Q in the window
                self._scanner_running = False
                break
        cap.release()
        cv2.destroyAllWindows()
        self._scanner_running = False

# ---------------------------
# GUI PAGES
# ---------------------------
class DashboardPage(QWidget):
    def __init__(self, main_win):
        super().__init__()
        self.main = main_win
        self.db = DB
        layout = QVBoxLayout(self)
        title = QLabel("Dashboard")
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        layout.addWidget(title)

        # role filter and controls
        ctrl = QHBoxLayout()
        ctrl.addWidget(QLabel("Filter role:"))
        self.role_filter = QComboBox(); self.role_filter.addItems(["All","Youth","Young Pro","Tanders"])
        self.role_filter.currentIndexChanged.connect(self.refresh)
        ctrl.addWidget(self.role_filter)
        btn_refresh = QPushButton("Refresh"); btn_refresh.clicked.connect(self.refresh)
        ctrl.addWidget(btn_refresh)
        layout.addLayout(ctrl)

        # today's table
        layout.addWidget(QLabel("Today's Attendance"))
        self.tbl_today = QTableWidget(); self.tbl_today.setColumnCount(6)
        self.tbl_today.setHorizontalHeaderLabels(["ID","Name","Role","Login","Logout","Event"])
        layout.addWidget(self.tbl_today)

        # quick history preview
        layout.addWidget(QLabel("Recent Attendance (preview)"))
        self.txt_history = QTextEdit(); self.txt_history.setReadOnly(True)
        layout.addWidget(self.txt_history, 1)

        # bottom summary
        bottom = QHBoxLayout()
        self.lbl_top = QLabel("Top 5: -"); self.lbl_absent = QLabel("Absent >3w: -")
        bottom.addWidget(self.lbl_top); bottom.addStretch(); bottom.addWidget(self.lbl_absent)
        layout.addLayout(bottom)

        self.refresh()

    def refresh(self):
        self.tbl_today.setRowCount(0)
        role_sel = self.role_filter.currentText()
        rows = self.db.todays_attendance()
        for r in rows:
            _, mid, name, role, login, logout, event = r
            if role_sel != "All" and role != role_sel:
                continue
            row = self.tbl_today.rowCount(); self.tbl_today.insertRow(row)
            for col, val in enumerate([mid, name, role, login or "", logout or "", event or ""]):
                self.tbl_today.setItem(row, col, QTableWidgetItem(str(val)))
            # highlight absent >3w is not applicable here; it's today's list
        # history preview
        hist = self.db.attendance_history()
        lines = []
        for rec in hist[:30]:
            lines.append(f"{rec[4]} | {rec[2]} | in:{rec[5] or '-'} out:{rec[6] or '-'}")
        self.txt_history.setText("\n".join(lines))
        # top5 and absent
        top = self.db.attendance_counts()[:5]
        self.lbl_top.setText("Top 5: " + (", ".join([f"{t[0]}({t[1]})" for t in top]) if top else "-"))
        absent = self.db.absent_for_weeks(3)
        self.lbl_absent.setText("Absent>3w: " + (", ".join([a[1] for a in absent[:5]]) if absent else "-"))

class AttendancePage(QWidget):
    def __init__(self, main_win):
        super().__init__()
        self.main = main_win
        self.db = DB
        layout = QVBoxLayout(self)
        title = QLabel("Attendance")
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        layout.addWidget(title)

        # Manual controls
        form = QHBoxLayout()
        form.addWidget(QLabel("Member:"))
        self.combo_members = QComboBox()
        form.addWidget(self.combo_members, 2)
        form.addWidget(QLabel("Event:"))
        self.event_box = QComboBox(); self.event_box.addItems(["General","Raged Youth","Sunday Service","Prayer Meeting"])
        form.addWidget(self.event_box, 1)
        btn_manual = QPushButton("Toggle Check-in/Check-out (Manual)")
        btn_manual.clicked.connect(self.toggle_manual)
        form.addWidget(btn_manual)
        layout.addLayout(form)

        # scanner controls
        srow = QHBoxLayout()
        self.btn_start = QPushButton("Start Scanner (continuous)"); self.btn_stop = QPushButton("Stop Scanner")
        self.btn_start.clicked.connect(self.main.start_scanner); self.btn_stop.clicked.connect(self.main.stop_scanner)
        srow.addWidget(self.btn_start); srow.addWidget(self.btn_stop)
        layout.addLayout(srow)

        # today's list
        layout.addWidget(QLabel("Today's Attendance"))
        self.tbl_today = QTableWidget(); self.tbl_today.setColumnCount(6)
        self.tbl_today.setHorizontalHeaderLabels(["ID","Name","Role","Login","Logout","Event"])
        layout.addWidget(self.tbl_today)
        btn_refresh = QPushButton("Refresh Today's List"); btn_refresh.clicked.connect(self.refresh_today)
        layout.addWidget(btn_refresh)

        self.reload_members_combo()
        self.refresh_today()

    def reload_members_combo(self):
        self.combo_members.clear()
        for m in self.db.list_members(active_only=True):
            self.combo_members.addItem(f"{m[1]} ({m[5]})", m[0])

    def toggle_manual(self):
        member_id = self.combo_members.currentData()
        if not member_id:
            QMessageBox.warning(self, "Select", "Choose a member")
            return
        row = self.db.ensure_attendance_row(member_id)
        if row[1] is None:
            ok, msg = self.db.set_login(member_id)
            QMessageBox.information(self, "Manual", msg)
        elif row[2] is None:
            ok, msg = self.db.set_logout(member_id)
            QMessageBox.information(self, "Manual", msg)
        else:
            QMessageBox.information(self, "Manual", "Already logged in and out today.")
        # refresh both dashboard & today's list
        self.main.refresh_signal.emit()

    def refresh_today(self):
        self.tbl_today.setRowCount(0)
        rows = self.db.todays_attendance()
        for r in rows:
            _, mid, name, role, login, logout, event = r
            row = self.tbl_today.rowCount(); self.tbl_today.insertRow(row)
            for col, val in enumerate([mid, name, role, login or "", logout or "", event or ""]):
                self.tbl_today.setItem(row, col, QTableWidgetItem(str(val)))

class MembersPage(QWidget):
    def __init__(self, main_win):
        super().__init__()
        self.main = main_win
        self.db = DB
        layout = QVBoxLayout(self)
        title = QLabel("Members")
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        layout.addWidget(title)

        form = QFormLayout()
        self.name = QLineEdit(); self.email = QLineEdit(); self.contact = QLineEdit(); self.facebook = QLineEdit()
        self.role = QComboBox(); self.role.addItems(["Youth","Young Pro","Tanders"])
        form.addRow("Full name:", self.name)
        form.addRow("Email:", self.email)
        form.addRow("Contact (optional):", self.contact)
        form.addRow("Facebook (optional):", self.facebook)
        form.addRow("Role:", self.role)
        layout.addLayout(form)

        btn_reg = QPushButton("Register & Email QR"); btn_reg.clicked.connect(self.register_member)
        layout.addWidget(btn_reg)

        # members table
        self.table = QTableWidget(); self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["ID","Name","Email","Contact","Facebook","Role"])
        layout.addWidget(self.table)
        btns = QHBoxLayout()
        btn_reload = QPushButton("Reload"); btn_reload.clicked.connect(self.reload_table)
        btn_deact = QPushButton("Deactivate Selected"); btn_deact.clicked.connect(self.deactivate_selected)
        btns.addWidget(btn_reload); btns.addWidget(btn_deact)
        layout.addLayout(btns)

        self.reload_table()

    def register_member(self):
        name = self.name.text().strip(); email = self.email.text().strip()
        contact = self.contact.text().strip(); facebook = self.facebook.text().strip()
        role = self.role.currentText()
        if not name or not email:
            QMessageBox.warning(self, "Missing", "Name and email required.")
            return
        member_id = self.db.add_member(name, email, contact, facebook, role)
        qr_path = f"qr_{member_id}.png"
        generate_qr_image(str(member_id), qr_path)
        ok, msg = send_qr_email(email, member_id, qr_path)
        if ok:
            QMessageBox.information(self, "Registered", f"Saved and emailed QR to {email}.")
        else:
            QMessageBox.information(self, "Registered", f"Saved. Email issue: {msg} (QR saved at {qr_path})")
        self.name.clear(); self.email.clear(); self.contact.clear(); self.facebook.clear()
        self.reload_table()
        self.main.refresh_signal.emit()

    def reload_table(self):
        self.table.setRowCount(0)
        for m in self.db.list_members(active_only=False):
            r = self.table.rowCount(); self.table.insertRow(r)
            for i, val in enumerate(m[:6]):
                self.table.setItem(r, i, QTableWidgetItem(str(val)))

    def deactivate_selected(self):
        r = self.table.currentRow()
        if r < 0:
            QMessageBox.warning(self, "Select", "Select a member first")
            return
        mid = int(self.table.item(r,0).text())
        self.db.set_member_inactive(mid)
        QMessageBox.information(self, "Deactivated", "Member set to inactive.")
        self.reload_table()
        self.main.refresh_signal.emit()

class ReportsPage(QWidget):
    def __init__(self, main_win=None):
        super().__init__()
        self.db = DB
        layout = QVBoxLayout(self)
        title = QLabel("Reports")
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        layout.addWidget(title)

        btn_chart = QPushButton("Show Attendance Chart")
        btn_chart.clicked.connect(self.show_chart)
        btn_absent = QPushButton("Show Absent >3 weeks")
        btn_absent.clicked.connect(self.show_absent)
        btn_export = QPushButton("Export Attendance CSV")
        btn_export.clicked.connect(self.export_csv)
        btn_export_xlsx = QPushButton("Export Attendance XLSX")
        btn_export_xlsx.clicked.connect(self.export_xlsx)
        layout.addWidget(btn_chart); layout.addWidget(btn_absent); layout.addWidget(btn_export); layout.addWidget(btn_export_xlsx)

        self.info = QTextEdit(); self.info.setReadOnly(True)
        layout.addWidget(self.info)
        self.refresh_preview()

    def refresh_preview(self):
        counts = self.db.attendance_counts()
        preview = "\n".join([f"{n}: {t}" for n, t in counts[:10]])
        self.info.setText("Top attendees:\n" + (preview if preview else "No data"))

    def show_chart(self):
        data = self.db.attendance_counts()
        if not data:
            QMessageBox.information(self, "Chart", "No attendance data yet.")
            return
        names = [d[0] for d in data]; totals = [d[1] for d in data]
        plt.figure(figsize=(8,4)); plt.bar(names, totals); plt.xticks(rotation=45); plt.tight_layout(); plt.show()

    def show_absent(self):
        rows = self.db.absent_for_weeks(3)
        if not rows:
            QMessageBox.information(self, "Absent", "No one absent >3 weeks.")
            return
        txt = "\n".join([f"{r[1]} <{r[2]}>" for r in rows])
        QMessageBox.information(self, "Absent >3 weeks", txt)

    def export_csv(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV", "attendance_export.csv", "CSV Files (*.csv)")
        if not path:
            return
        rows = self.db.attendance_history()
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["AttendanceID","MemberID","Name","Role","Date","Login","Logout","Event"])
            for r in rows:
                w.writerow(r)
        QMessageBox.information(self, "Export", f"Saved CSV to {path}")

    def export_xlsx(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save XLSX", "attendance_export.xlsx", "Excel Files (*.xlsx)")
        if not path:
            return
        rows = self.db.attendance_history()
        df = pd.DataFrame(rows, columns=["AttendanceID","MemberID","Name","Role","Date","Login","Logout","Event"])
        df.to_excel(path, index=False)
        QMessageBox.information(self, "Export", f"Saved Excel to {path}")

# ---------------------------
# Application entry
# ---------------------------
def main():
    # start flask server in background and create poster qr
    flask_thread = start_flask_server_background()

    app = QApplication(sys.argv)
    if qdarkstyle:
        app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())

    login = LoginDialog(DB)
    if login.exec_() != QDialog.Accepted:
        print("Login cancelled; exiting.")
        return

    mainwin = MainWindow()
    mainwin.show()

    # print poster URL for admin convenience
    ip = local_ip()
    print(f"Poster registration URL: http://{ip}:{FLASK_PORT}/register")
    print(f"Poster QR file generated: {POSTER_QR_FILENAME} (print this and place at the entrance)")

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
