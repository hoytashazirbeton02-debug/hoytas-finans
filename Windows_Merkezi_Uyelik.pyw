"""Windows client for the isolated central membership flow."""
from concurrent.futures import ThreadPoolExecutor
import os
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from hybrid_membership_client import MembershipClient


class MembershipWindow(tk.Tk):
    def __init__(self, client):
        super().__init__()
        self.client = client
        self.session = None
        self.worker = ThreadPoolExecutor(max_workers=1)
        self.title("HOYTAŞ | Merkezi Üyelik")
        self.protocol("WM_DELETE_WINDOW", self.close)
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Kullanıcı adı").grid(row=0, column=0, sticky="w", pady=5)
        self.username = ttk.Entry(frame, width=30)
        self.username.grid(row=0, column=1, pady=5)
        ttk.Label(frame, text="Şifre").grid(row=1, column=0, sticky="w", pady=5)
        self.password = ttk.Entry(frame, width=30, show="*")
        self.password.grid(row=1, column=1, pady=5)
        self.buttons = [
            ttk.Button(frame, text="Giriş Yap", command=self.login),
            ttk.Button(frame, text="Yeni Üyelik", command=self.request),
            ttk.Button(frame, text="Onayla", command=self.approve),
        ]
        self.buttons[0].grid(row=2, column=0, pady=8)
        self.buttons[1].grid(row=2, column=1, pady=8)
        self.buttons[2].grid(row=4, column=1, sticky="e", pady=5)
        self.buttons[2].grid_remove()
        self.pending = tk.Listbox(frame, width=43, height=7, exportselection=False)
        self.pending.grid(row=3, column=0, columnspan=2, sticky="ew", pady=5)
        self.pending.grid_remove()
        self.status = ttk.Label(frame, text="")
        self.status.grid(row=5, column=0, columnspan=2, sticky="w")
        self.username.focus()

    def close(self):
        self.worker.shutdown(wait=False, cancel_futures=True)
        self.destroy()

    def run(self, work, success):
        for button in self.buttons:
            button.state(["disabled"])
        self.status.config(text="Bağlanıyor…")
        task = self.worker.submit(work)

        def poll():
            if not task.done():
                self.after(100, poll)
                return
            self.status.config(text="")
            for button in self.buttons:
                button.state(["!disabled"])
            try:
                success(task.result())
            except (OSError, ValueError, RuntimeError, KeyError) as error:
                messagebox.showerror("HOYTAŞ", str(error), parent=self)

        self.after(100, poll)

    def login(self):
        username, password = self.username.get(), self.password.get()
        self.password.delete(0, "end")
        self.session = None
        self.pending.grid_remove()
        self.buttons[2].grid_remove()

        def success(result):
            if result["role"] == "admin":
                self.session = result["session"]
                self.refresh()
            else:
                self.status.config(text="Merkezi kullanıcı girişi doğrulandı.")

        self.run(lambda: self.client.login(username, password), success)

    def request(self):
        invitation = simpledialog.askstring("Yeni üyelik", "Yöneticiden aldığınız davet kodu", show="*", parent=self)
        if not invitation:
            return
        username, password = self.username.get(), self.password.get()
        self.password.delete(0, "end")
        self.run(lambda: self.client.request(username, password, invitation.strip()),
                 lambda _: self.status.config(text="Başvuru yönetici onayına gönderildi."))

    def refresh(self):
        token = self.session

        def update(users):
            self.pending.delete(0, "end")
            for username in users:
                self.pending.insert("end", username)
            self.pending.grid()
            self.buttons[2].grid()
            self.status.config(text="Onay bekleyen başvurular")

        self.run(lambda: self.client.pending(token), update)

    def approve(self):
        selected = self.pending.curselection()
        if not selected or not self.session:
            return
        username, token = self.pending.get(selected[0]), self.session
        if messagebox.askyesno("Üyelik onayı", f"{username} hesabını onaylıyor musunuz?", parent=self):
            self.run(lambda: self.client.approve(token, username), lambda _: self.refresh())


if __name__ == "__main__":
    endpoint = os.environ.get("HOYTAS_MEMBERSHIP_URL", "").strip()
    if endpoint:
        MembershipWindow(MembershipClient(endpoint)).mainloop()
    else:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("HOYTAŞ", "Merkezi üyelik sunucusu henüz yapılandırılmamış.", parent=root)
        root.destroy()
