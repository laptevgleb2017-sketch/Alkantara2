import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import os


class AssetManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Ведомость остатков ОС, НМА, НПА")
        self.root.geometry("1500x800")
        self.root.configure(bg='#f0f0f0')

        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        self.data_dir = os.path.join(desktop, 'Учёт имущества')
        os.makedirs(self.data_dir, exist_ok=True)
        self.excel_path = os.path.join(self.data_dir, 'assets.xlsx')

        self.assets = []
        self.tree_items = {}      # item_id -> asset dict
        self.filter_combos = {}   # имя фильтра -> Combobox

        if not os.path.exists(self.excel_path):
            self.create_excel_file()
        self.load_assets()
        self.setup_ui()
        self.refresh()

    # ---------------- Excel ----------------
    def create_excel_file(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Ведомость"
        headers = ['Счет', 'Ответственный', 'Место хранения', '№ п/п',
                   'Основное средство', 'Инвентарный номер', 'Дата принятия',
                   'Балансовая стоимость', 'Количество', 'Сумма амортизации']
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='2196F3', end_color='2196F3', fill_type='solid')
        widths = [28, 28, 28, 8, 40, 18, 15, 18, 12, 18]
        for col, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w
        wb.save(self.excel_path)

    def load_assets(self):
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb.active
            self.assets = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and row[4]:
                    asset = {
                        'account':      str(row[0] or ''),
                        'responsible':  str(row[1] or ''),
                        'location':     str(row[2] or ''),
                        'num':          row[3] if isinstance(row[3], int) else len(self.assets) + 1,
                        'name':         str(row[4] or ''),
                        'inventory':    str(row[5] or ''),
                        'date':         str(row[6] or ''),
                        'cost':         float(row[7]) if isinstance(row[7], (int, float)) else 0.0,
                        'quantity':     int(row[8]) if isinstance(row[8], (int, float)) else 1,
                        'depreciation': float(row[9]) if isinstance(row[9], (int, float)) else 0.0,
                    }
                    asset['residual'] = asset['cost'] - asset['depreciation']
                    self.assets.append(asset)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {e}")

    def save_to_excel(self):
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Ведомость"
            headers = ['Счет', 'Ответственный', 'Место хранения', '№ п/п',
                       'Основное средство', 'Инвентарный номер', 'Дата принятия',
                       'Балансовая стоимость', 'Количество', 'Сумма амортизации']
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=h)
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill(start_color='2196F3', end_color='2196F3', fill_type='solid')
                cell.alignment = Alignment(horizontal='center')
            for r, a in enumerate(self.assets, 2):
                ws.cell(row=r, column=1, value=a['account'])
                ws.cell(row=r, column=2, value=a['responsible'])
                ws.cell(row=r, column=3, value=a['location'])
                ws.cell(row=r, column=4, value=a['num'])
                ws.cell(row=r, column=5, value=a['name'])
                ws.cell(row=r, column=6, value=a['inventory'])
                ws.cell(row=r, column=7, value=a['date'])
                ws.cell(row=r, column=8, value=a['cost'])
                ws.cell(row=r, column=9, value=a['quantity'])
                ws.cell(row=r, column=10, value=a['depreciation'])
            widths = [28, 28, 28, 8, 40, 18, 15, 18, 12, 18]
            for col, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(col)].width = w
            ws.freeze_panes = 'A2'
            wb.save(self.excel_path)
            return True
        except PermissionError:
            messagebox.showwarning("Внимание",
                "Файл занят (возможно, открыт в Excel).\nЗакройте Excel и повторите.")
            return False
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить: {e}")
            return False

    # ---------------- UI ----------------
    def setup_ui(self):
        top = tk.Frame(self.root, bg='#f0f0f0')
        top.pack(fill=tk.X, padx=15, pady=(15, 5))
        tk.Label(top, text="📊 Ведомость остатков ОС, НМА, НПА",
                 font=('Segoe UI', 18, 'bold'), bg='#f0f0f0').pack(side=tk.LEFT)

        btns = tk.Frame(top, bg='#f0f0f0')
        btns.pack(side=tk.RIGHT)
        for text, cmd, color in [
            ("➕ Добавить", self.add_asset, '#4CAF50'),
            ("✏️ Изменить", self.edit_asset, '#FF9800'),
            ("🗑 Удалить", self.delete_asset, '#F44336'),
            ("📥 Импорт", self.import_from_excel, '#9C27B0'),
            ("📤 Экспорт", self.export_to_excel, '#2196F3'),
        ]:
            tk.Button(btns, text=text, command=cmd, bg=color, fg='white',
                      relief=tk.FLAT, padx=12, pady=6, cursor='hand2').pack(side=tk.LEFT, padx=3)

        # Фильтры
        flt = tk.Frame(self.root, bg='#ffffff', relief=tk.RAISED, bd=1)
        flt.pack(fill=tk.X, padx=15, pady=5)
        inner = tk.Frame(flt, bg='#ffffff')
        inner.pack(fill=tk.X, padx=10, pady=8)

        self.f_account = tk.StringVar()
        self.f_resp = tk.StringVar()
        self.f_loc = tk.StringVar()
        self.f_search = tk.StringVar()

        tk.Label(inner, text="🔍 Поиск:", bg='#ffffff').pack(side=tk.LEFT)
        e = tk.Entry(inner, textvariable=self.f_search, width=25)
        e.pack(side=tk.LEFT, padx=5)
        e.bind('<KeyRelease>', lambda ev: self.refresh())

        # Явно привязываем комбобоксы к именам фильтров
        filter_defs = [
            ("Счет:",           self.f_account, 'account'),
            ("Ответственный:",  self.f_resp,    'responsible'),
            ("Место хранения:", self.f_loc,     'location'),
        ]
        for label, var, key in filter_defs:
            tk.Label(inner, text=label, bg='#ffffff').pack(side=tk.LEFT, padx=(15, 3))
            cb = ttk.Combobox(inner, textvariable=var, width=22, state='readonly')
            cb.pack(side=tk.LEFT)
            cb.bind('<<ComboboxSelected>>', lambda ev: self.refresh())
            self.filter_combos[key] = cb

        tk.Button(inner, text="✖ Сброс", command=self.reset_filters,
                  bg='#9E9E9E', fg='white', relief=tk.FLAT,
                  padx=8, pady=3).pack(side=tk.LEFT, padx=10)

        # Таблица
        table = tk.Frame(self.root, bg='#ffffff', relief=tk.SOLID, bd=1)
        table.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        cols = ('num', 'name', 'inventory', 'date',
                'cost', 'qty', 'depreciation', 'residual')
        heads = ('№ п/п', 'Основное средство', 'Инв. номер', 'Дата принятия',
                 'Балансовая ст-ть', 'Кол-во', 'Амортизация', 'Остаточная ст-ть')
        widths = (60, 420, 150, 110, 150, 70, 150, 150)

        self.tree = ttk.Treeview(table, columns=cols, show='tree headings')
        self.tree.heading('#0', text='Счёт / Ответственный / Место')
        self.tree.column('#0', width=380, anchor='w')
        for c, h, w in zip(cols, heads, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor='center' if c != 'name' else 'w')

        vs = ttk.Scrollbar(table, orient='vertical', command=self.tree.yview)
        hs = ttk.Scrollbar(table, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vs.grid(row=0, column=1, sticky='ns')
        hs.grid(row=1, column=0, sticky='ew')
        table.grid_rowconfigure(0, weight=1)
        table.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use('default')
        style.configure('Treeview', rowheight=28, font=('Segoe UI', 10))
        style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'))
        self.tree.tag_configure('group1', background='#E3F2FD', font=('Segoe UI', 10, 'bold'))
        self.tree.tag_configure('group2', background='#F1F8E9', font=('Segoe UI', 10, 'italic'))
        self.tree.tag_configure('group3', background='#FFF8E1', font=('Segoe UI', 10))

        self.tree.bind('<Double-Button-1>', self.on_double_click)

        self.totals = tk.Label(self.root, text='', font=('Segoe UI', 11, 'bold'),
                                bg='#f0f0f0', anchor='e', fg='#1a237e')
        self.totals.pack(fill=tk.X, padx=15, pady=(0, 10))

    # ---------------- Логика отображения ----------------
    def reset_filters(self):
        self.f_account.set('')
        self.f_resp.set('')
        self.f_loc.set('')
        self.f_search.set('')
        self.refresh()

    def filtered(self):
        acc = self.f_account.get()
        resp = self.f_resp.get()
        loc = self.f_loc.get()
        search = self.f_search.get().strip().lower()
        out = []
        for a in self.assets:
            if acc and a['account'] != acc:
                continue
            if resp and a['responsible'] != resp:
                continue
            if loc and a['location'] != loc:
                continue
            if search:
                hay = ' '.join([a['name'], a['inventory'], a['responsible'],
                                a['location'], a['account']]).lower()
                if search not in hay:
                    continue
            out.append(a)
        return out

    def update_filter_values(self):
        for key, field in [('account', 'account'),
                           ('responsible', 'responsible'),
                           ('location', 'location')]:
            values = sorted({a[field] for a in self.assets if a[field]})
            cb = self.filter_combos.get(key)
            if cb is not None:
                current = cb.get()
                cb['values'] = [''] + values
                if current not in values:
                    cb.set('')

    def refresh(self):
        self.update_filter_values()
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.tree_items.clear()

        data = self.filtered()

        groups = {}
        for a in data:
            groups.setdefault(a['account'], {}) \
                  .setdefault(a['responsible'], {}) \
                  .setdefault(a['location'], []).append(a)

        for acc in sorted(groups.keys()):
            acc_items = [x for r in groups[acc].values()
                         for loc in r.values() for x in loc]
            acc_sum = self._sum(acc_items)
            acc_id = self.tree.insert('', 'end',
                text=f"📁 {acc or '— без счёта'}",
                values=('', '', '', '', *acc_sum[:1], acc_sum[1], acc_sum[2], acc_sum[3]),
                tags=('group1',), open=True)

            for resp in sorted(groups[acc].keys()):
                resp_items = [x for loc in groups[acc][resp].values() for x in loc]
                resp_sum = self._sum(resp_items)
                resp_id = self.tree.insert(acc_id, 'end',
                    text=f"👤 {resp or '— без ответственного'}",
                    values=('', '', '', '', *resp_sum[:1], resp_sum[1], resp_sum[2], resp_sum[3]),
                    tags=('group2',), open=False)

                for loc in sorted(groups[acc][resp].keys()):
                    loc_items = groups[acc][resp][loc]
                    loc_sum = self._sum(loc_items)
                    loc_id = self.tree.insert(resp_id, 'end',
                        text=f"📍 {loc or '— без места'}",
                        values=('', '', '', '', *loc_sum[:1], loc_sum[1], loc_sum[2], loc_sum[3]),
                        tags=('group3',), open=False)

                    for a in sorted(loc_items, key=lambda x: x['num']):
                        iid = self.tree.insert(loc_id, 'end',
                            text='', values=(
                                a['num'], a['name'], a['inventory'], a['date'],
                                f"{a['cost']:,.2f}", a['quantity'],
                                f"{a['depreciation']:,.2f}", f"{a['residual']:,.2f}"))
                        self.tree_items[iid] = a

        self.totals.config(text=self._totals_text(data))

    @staticmethod
    def _sum(items):
        return (sum(x['cost'] for x in items),
                sum(x['quantity'] for x in items),
                sum(x['depreciation'] for x in items),
                sum(x['residual'] for x in items))

    def _totals_text(self, data):
        c, q, d, r = self._sum(data)
        return (f"ИТОГО:  Балансовая стоимость: {c:,.2f} ₽   |   "
                f"Количество: {q}   |   Амортизация: {d:,.2f} ₽   |   "
                f"Остаточная: {r:,.2f} ₽")

    # ---------------- CRUD ----------------
    def on_double_click(self, event):
        iid = self.tree.identify_row(event.y)
        if iid in self.tree_items:
            self.edit_asset(self.tree_items[iid])

    def add_asset(self):
        dlg = AssetDialog(self.root, "Добавить основное средство")
        self.root.wait_window(dlg.dialog)
        if dlg.result:
            dlg.result['num'] = len(self.assets) + 1
            dlg.result['residual'] = dlg.result['cost'] - dlg.result['depreciation']
            self.assets.append(dlg.result)
            self.save_to_excel()
            self.refresh()

    def edit_asset(self, asset=None):
        if asset is None:
            iid = self.tree.focus()
            asset = self.tree_items.get(iid)
        if not asset:
            messagebox.showwarning("Внимание",
                "Выберите актив (двойной клик по строке актива)")
            return
        dlg = AssetDialog(self.root, "Редактировать", asset)
        self.root.wait_window(dlg.dialog)
        if dlg.result:
            dlg.result['num'] = asset['num']
            dlg.result['residual'] = dlg.result['cost'] - dlg.result['depreciation']
            idx = self.assets.index(asset)
            self.assets[idx] = dlg.result
            self.save_to_excel()
            self.refresh()

    def delete_asset(self):
        iid = self.tree.focus()
        asset = self.tree_items.get(iid)
        if not asset:
            messagebox.showwarning("Внимание",
                "Выберите актив (двойной клик по строке актива)")
            return
        if messagebox.askyesno("Подтверждение", f"Удалить «{asset['name']}»?"):
            self.assets.remove(asset)
            self.save_to_excel()
            self.refresh()

    # ---------------- Импорт / Экспорт ----------------
    def import_from_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            count = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and any(row):
                    if len(row) >= 10 and row[4]:
                        asset = {
                            'account':      str(row[0] or ''),
                            'responsible':  str(row[1] or ''),
                            'location':     str(row[2] or ''),
                            'num':          len(self.assets) + 1,
                            'name':         str(row[4] or ''),
                            'inventory':    str(row[5] or ''),
                            'date':         str(row[6] or ''),
                            'cost':         float(row[7]) if isinstance(row[7], (int, float)) else 0.0,
                            'quantity':     int(row[8]) if isinstance(row[8], (int, float)) else 1,
                            'depreciation': float(row[9]) if isinstance(row[9], (int, float)) else 0.0,
                        }
                        asset['residual'] = asset['cost'] - asset['depreciation']
                        self.assets.append(asset)
                        count += 1
            self.save_to_excel()
            self.refresh()
            messagebox.showinfo("Импорт", f"Загружено записей: {count}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось импортировать: {e}")

    def export_to_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             initialfile="ведомость.xlsx")
        if not path:
            return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Ведомость"
            headers = ['Счет', 'Ответственный', 'Место хранения', '№ п/п',
                       'Основное средство', 'Инвентарный номер', 'Дата принятия',
                       'Балансовая стоимость', 'Количество', 'Сумма амортизации',
                       'Остаточная стоимость']
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=h)
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill(start_color='2196F3', end_color='2196F3', fill_type='solid')

            data = self.filtered()
            r = 1
            for r, a in enumerate(data, 2):
                ws.cell(row=r, column=1, value=a['account'])
                ws.cell(row=r, column=2, value=a['responsible'])
                ws.cell(row=r, column=3, value=a['location'])
                ws.cell(row=r, column=4, value=a['num'])
                ws.cell(row=r, column=5, value=a['name'])
                ws.cell(row=r, column=6, value=a['inventory'])
                ws.cell(row=r, column=7, value=a['date'])
                ws.cell(row=r, column=8, value=a['cost'])
                ws.cell(row=r, column=9, value=a['quantity'])
                ws.cell(row=r, column=10, value=a['depreciation'])
                ws.cell(row=r, column=11, value=a['residual'])

            total_row = r + 1 if data else 2
            ws.cell(row=total_row, column=5, value='ИТОГО').font = Font(bold=True)
            ws.cell(row=total_row, column=8,
                    value=sum(a['cost'] for a in data)).font = Font(bold=True)
            ws.cell(row=total_row, column=9,
                    value=sum(a['quantity'] for a in data)).font = Font(bold=True)
            ws.cell(row=total_row, column=10,
                    value=sum(a['depreciation'] for a in data)).font = Font(bold=True)
            ws.cell(row=total_row, column=11,
                    value=sum(a['residual'] for a in data)).font = Font(bold=True)

            widths = [28, 28, 28, 8, 40, 18, 15, 18, 12, 18, 18]
            for col, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(col)].width = w
            wb.save(path)
            messagebox.showinfo("Успех", f"Экспортировано в {path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать: {e}")


class AssetDialog:
    def __init__(self, parent, title, asset=None):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("520x560")
        self.dialog.configure(bg='#f0f0f0')
        self.dialog.resizable(False, False)
        self.dialog.grab_set()

        self.result = None
        self.asset = asset or {}

        tk.Label(self.dialog, text=title, font=('Segoe UI', 14, 'bold'),
                 bg='#f0f0f0').pack(pady=15)

        form = tk.Frame(self.dialog, bg='#f0f0f0')
        form.pack(padx=25, fill=tk.X)

        fields = [
            ('Счет:',                         'account',      True),
            ('Ответственный:',                'responsible',  True),
            ('Место хранения:',               'location',     True),
            ('Основное средство:',            'name',         True),
            ('Инвентарный номер:',            'inventory',    False),
            ('Дата принятия (ДД.ММ.ГГГГ):',   'date',         False),
            ('Балансовая стоимость:',         'cost',         False),
            ('Количество:',                   'quantity',     False),
            ('Сумма амортизации:',            'depreciation', False),
        ]
        self.entries = {}
        for i, (label, key, req) in enumerate(fields):
            tk.Label(form, text=label + (' *' if req else ''),
                     bg='#f0f0f0', font=('Segoe UI', 10)).grid(row=i, column=0, sticky='w', pady=4)
            default = self.asset.get(key, '')
            if key == 'quantity' and not default:
                default = 1
            var = tk.StringVar(value=str(default) if default != '' else '')
            tk.Entry(form, textvariable=var, width=32,
                     font=('Segoe UI', 10)).grid(row=i, column=1, sticky='w', padx=8)
            self.entries[key] = var

        btns = tk.Frame(self.dialog, bg='#f0f0f0')
        btns.pack(pady=20)
        tk.Button(btns, text="💾 Сохранить", command=self.save,
                  bg='#4CAF50', fg='white', relief=tk.FLAT,
                  padx=20, pady=8, font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=8)
        tk.Button(btns, text="Отмена", command=self.dialog.destroy,
                  bg='#9E9E9E', fg='white', relief=tk.FLAT,
                  padx=20, pady=8, font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=8)

    def save(self):
        for key in ('account', 'responsible', 'location', 'name'):
            if not self.entries[key].get().strip():
                messagebox.showwarning("Внимание", "Заполните обязательные поля (*)")
                return
        try:
            cost_str = self.entries['cost'].get().replace(' ', '').replace(',', '.') or '0'
            dep_str = self.entries['depreciation'].get().replace(' ', '').replace(',', '.') or '0'
            cost = float(cost_str)
            dep = float(dep_str)
            qty = int(self.entries['quantity'].get().strip() or 1)
        except ValueError:
            messagebox.showerror("Ошибка",
                "Стоимость, амортизация и количество должны быть числами")
            return

        self.result = {
            'account':      self.entries['account'].get().strip(),
            'responsible':  self.entries['responsible'].get().strip(),
            'location':     self.entries['location'].get().strip(),
            'name':         self.entries['name'].get().strip(),
            'inventory':    self.entries['inventory'].get().strip(),
            'date':         self.entries['date'].get().strip(),
            'cost':         cost,
            'quantity':     qty,
            'depreciation': dep,
        }
        self.dialog.destroy()


def main():
    root = tk.Tk()
    AssetManager(root)
    root.mainloop()


if __name__ == "__main__":
    main()
