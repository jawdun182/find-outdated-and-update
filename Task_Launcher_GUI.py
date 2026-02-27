import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import winreg
import os

class TaskLauncherGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("App Updater")
        self.root.geometry("500x500")
        self.root.resizable(True, True)

        # Configure Style
        style = ttk.Style()
        style.theme_use('clam')  # Use a clean theme
        
        # Define styles
        style.configure('Header.TLabel', font=('Segoe UI', 16, 'bold'), foreground='#333333')
        style.configure('TButton', font=('Segoe UI', 11), padding=8)
        style.configure('Status.TLabel', font=('Segoe UI', 9), foreground='#666666')

        # Header
        header_frame = ttk.Frame(root)
        header_frame.pack(pady=25)
        ttk.Label(header_frame, text="Find Outdated and Update", style='Header.TLabel').pack()

        # Button Container
        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=10, fill='x', padx=60)

        # --- Task Buttons ---

        # 1. list apps
        self.btn_updates = ttk.Button(btn_frame, text="List of installed apps", command=self.open_apps)
        self.btn_updates.pack(fill='x', pady=8)

        # 2. check versions of apps
        self.btn_sysinfo = ttk.Button(btn_frame, text="Check outdated apps", command=self.scan_for_outdated_apps)
        self.btn_sysinfo.pack(fill='x', pady=8)

        # 3. update the outdated apps
        self.btn_taskmgr = ttk.Button(btn_frame, text="Update outdated apps", command=self.run_update_script)
        self.btn_taskmgr.pack(fill='x', pady=8)


        # Separator
        ttk.Separator(root, orient='horizontal').pack(fill='x', padx=30, pady=20)

        # Close Button
        self.btn_close = ttk.Button(root, text="Quit", command=root.quit)
        self.btn_close.pack(pady=5, ipadx=20)

    def open_apps(self):
        # Create a new window to list apps
        top = tk.Toplevel(self.root)
        top.title("Installed Applications")
        top.geometry("800x600")

        # Treeview
        columns = ("Name", "Version", "Publisher", "Install Date")
        tree = ttk.Treeview(top, columns=columns, show="headings")
        
        for col in columns:
            tree.heading(col, text=col, command=lambda c=col: self.sort_treeview(tree, c, False))
            tree.column(col, width=150)
        tree.column("Name", width=300)

        scrollbar = ttk.Scrollbar(top, orient="vertical", command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True)

        # Context Menu
        context_menu = tk.Menu(top, tearoff=0)
        context_menu.add_command(label="Open Folder Location", command=lambda: self.open_location(tree))
        tree.bind("<Button-3>", lambda event: self.show_context_menu(event, tree, context_menu))

        # Populate data
        tree.app_locations = {}
        for app in self.get_installed_apps():
            item_id = tree.insert("", "end", values=app[:4])
            tree.app_locations[item_id] = app[4]

    def show_context_menu(self, event, tree, menu):
        item = tree.identify_row(event.y)
        if item:
            tree.selection_set(item)
            menu.post(event.x_root, event.y_root)

    def open_location(self, tree):
        selected = tree.selection()
        if selected:
            path = tree.app_locations.get(selected[0])
            if path and os.path.exists(path):
                os.startfile(path)
            else:
                messagebox.showinfo("Info", "Location not available for this app.")

    def sort_treeview(self, tree, col, reverse):
        """Sort treeview items by column in ascending or descending order."""
        sorted_items = [(tree.set(k, col), k) for k in tree.get_children('')]
        sorted_items.sort(reverse=reverse)

        # Rearrange items in sorted positions
        for index, (val, k) in enumerate(sorted_items):
            tree.move(k, '', index)

        # Reverse sort next time
        tree.heading(col, command=lambda: self.sort_treeview(tree, col, not reverse))

    def get_installed_apps(self):
        apps = {}
        roots = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
        ]
        for hkey, path in roots:
            try:
                with winreg.OpenKey(hkey, path) as key:
                    for i in range(winreg.QueryInfoKey(key)[0]):
                        try:
                            sub_key_name = winreg.EnumKey(key, i)
                            with winreg.OpenKey(key, sub_key_name) as sub_key:
                                name = self.get_reg_value(sub_key, "DisplayName")
                                if not name:
                                    continue

                                # Filter out system components (drivers, updates, etc.)
                                if self.get_reg_value(sub_key, "SystemComponent") == 1:
                                    continue

                                version = self.get_reg_value(sub_key, "DisplayVersion")
                                publisher = self.get_reg_value(sub_key, "Publisher")
                                date = str(self.get_reg_value(sub_key, "InstallDate"))
                                location = self.get_reg_value(sub_key, "InstallLocation")
                                if len(date) == 8 and date.isdigit():
                                    date = f"{date[6:]}/{date[4:6]}/{date[:4]}"
                                
                                # Deduplication: prioritize entries with dates
                                if name in apps:
                                    if not apps[name][3] and date:
                                        apps[name] = (name, version, publisher, date, location)
                                else:
                                    apps[name] = (name, version, publisher, date, location)
                        except Exception:
                            continue
            except OSError:
                continue
        return sorted(apps.values(), key=lambda x: x[0].lower())

    def get_reg_value(self, key, name):
        try:
            return winreg.QueryValueEx(key, name)[0]
        except FileNotFoundError:
            return ""

    def scan_for_outdated_apps(self):
        # Placeholder for scanning logic
        messagebox.showinfo("Scan", "Scanning for outdated apps... (This is a placeholder)")

    def run_update_script(self):
        # Placeholder for update script logic
        messagebox.showinfo("Update", "Running update script... (This is a placeholder)")

if __name__ == "__main__":
    root = tk.Tk()
    app = TaskLauncherGUI(root)
    root.mainloop()
