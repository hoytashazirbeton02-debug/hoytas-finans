"""Central sign-in and approval dialogs for the desktop application."""
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from login_memory import LoginFields


def run_request(window, controls, work, success):
    for control in controls:
        control.state(["disabled"])
    worker = ThreadPoolExecutor(max_workers=1)
    task = worker.submit(work)
    worker.shutdown(wait=False)

    def poll():
        if not task.done():
            window.after(100, poll)
            return
        for control in controls:
            control.state(["!disabled"])
        try:
            result = task.result()
        except (OSError, ValueError, RuntimeError, KeyError) as error:
            messagebox.showerror("HOYTAŞ", str(error), parent=window)
        else:
            success(result)

    window.after(100, poll)


def sign_in(parent, client, *, setup_admin=False, prepare=None):
    window = tk.Toplevel(parent)
    window.title("HOYTAŞ Hazır Beton - Merkezi Giriş")
    window.transient(parent)
    window.grab_set()
    window.resizable(False, False)
    frame = ttk.Frame(window, padding=24)
    frame.pack()
    ttk.Label(frame, text="HOYTAŞ Hazır Beton", font=("Segoe UI", 20, "bold")).grid(
        row=0, column=0, columnspan=2, pady=(0, 15))
    ttk.Label(frame, text="Kullanıcı adı").grid(row=1, column=0, sticky="w", pady=5)
    ttk.Label(frame, text="Şifre").grid(row=2, column=0, sticky="w", pady=5)
    fields = LoginFields(frame, 1, "merkezi")
    username, password = fields.username, fields.password
    account = None

    def credentials():
        name, secret = username.get().strip(), password.get()
        password.configure(show="*")
        return name, secret

    def login():
        if login_button.instate(["disabled"]):
            return
        name, secret = credentials()
        remember = fields.remember.get()

        def signed_in(result):
            nonlocal account
            fields.save(name, secret, remember)
            account = result
            window.destroy()

        def authenticate():
            result = client.login(name, secret)
            if prepare is not None:
                prepare(result)
            return result

        run_request(window, controls, authenticate, signed_in)

    def request():
        invitation = simpledialog.askstring("Yeni üyelik", "Yöneticiden aldığınız davet kodu",
                                            show="*", parent=window)
        if invitation:
            name, secret = credentials()
            run_request(window, controls, lambda: client.request(name, secret, invitation.strip()),
                        lambda _: messagebox.showinfo("HOYTAŞ", "Başvuru yönetici onayına gönderildi.",
                                                      parent=window))

    def setup():
        code = simpledialog.askstring("İlk yönetici", "Render yönetici kurulum kodu",
                                      show="*", parent=window)
        if code:
            name, secret = credentials()
            run_request(window, controls, lambda: client.setup_admin(name, secret, code.strip()),
                        lambda _: messagebox.showinfo(
                            "HOYTAŞ", "Yönetici hesabı oluşturuldu. Şifrenizi tekrar yazıp giriş yapın.",
                            parent=window))

    login_button = ttk.Button(frame, text="Giriş Yap", command=login)
    request_button = ttk.Button(frame, text="Yeni Üye Ol", command=request)
    controls = [login_button, request_button]
    login_button.grid(row=4, column=0, pady=(14, 0))
    request_button.grid(row=4, column=1, pady=(14, 0))
    if setup_admin:
        setup_button = ttk.Button(frame, text="İlk Yönetici Hesabını Oluştur", command=setup)
        setup_button.grid(row=5, column=0, columnspan=2, pady=(8, 0))
        controls.append(setup_button)
    password.bind("<Return>", lambda _: login())
    username.focus_set()
    parent.wait_window(window)
    return account


def manage_pending(parent, client, session, publish=None):
    window = tk.Toplevel(parent)
    window.title("HOYTAŞ - Üyelik Onayları")
    window.transient(parent)
    window.grab_set()
    frame = ttk.Frame(window, padding=18)
    frame.pack(fill="both", expand=True)
    ttk.Label(frame, text="Onay bekleyen başvurular").pack(anchor="w")
    pending = tk.Listbox(frame, width=42, height=8, exportselection=False)
    pending.pack(fill="both", expand=True, pady=8)

    def refresh():
        def display(names):
            pending.delete(0, "end")
            for name in names:
                pending.insert("end", name)

        run_request(window, controls, lambda: client.pending(session), display)

    def approve():
        selection = pending.curselection()
        if not selection:
            return
        name = pending.get(selection[0])
        if messagebox.askyesno("Üyelik onayı", f"{name} hesabını onaylıyor musunuz?", parent=window):
            run_request(window, controls, lambda: client.approve(session, name), lambda _: refresh())

    controls = [ttk.Button(frame, text="Yenile", command=refresh),
                ttk.Button(frame, text="Onayla", command=approve)]
    if publish is not None:
        def publish_baseline():
            run_request(window, controls, publish, lambda _: messagebox.showinfo(
                "HOYTAŞ", "Başlangıç kayıtları onaylı bilgisayarlar için hazır.", parent=window))

        controls.append(ttk.Button(frame, text="Başlangıç Kayıtlarını Paylaş", command=publish_baseline))
    for control in controls:
        control.pack(side="left", padx=(0, 8))
    refresh()
