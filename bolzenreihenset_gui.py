from __future__ import annotations

import threading
import tkinter as tk
from tkinter import messagebox, ttk

from bolzenreihenset_pruefung import DISPLAY_COLUMNS, clear_cache, get_display_rows


EXAMPLE_ARTICLE = "HK_12/29_ze_L2_W_78_D_2.4_RHOM"


class BolzenreihensetApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Bolzenreihenset pruefen")
        self.geometry("1500x820")
        self.minsize(1100, 620)

        self.article_var = tk.StringVar(value=EXAMPLE_ARTICLE)
        self.status_var = tk.StringVar(value="Bereit. Es wird nichts in Excel geschrieben.")

        self._configure_style()
        self._build_layout()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.configure("Treeview", rowheight=26)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Status.TLabel", foreground="#444")

    def _build_layout(self) -> None:
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        control_frame = ttk.Frame(self, padding=(12, 12, 10, 12))
        control_frame.grid(row=0, column=0, sticky="ns")
        control_frame.columnconfigure(0, weight=1)

        ttk.Label(control_frame, text="Bolzenreihenset").grid(row=0, column=0, sticky="w")
        self.article_entry = ttk.Entry(control_frame, textvariable=self.article_var, width=44)
        self.article_entry.grid(row=1, column=0, sticky="ew", pady=(4, 8))
        self.article_entry.bind("<Return>", lambda _event: self.run_search())

        self.search_button = ttk.Button(control_frame, text="Pruefen", command=self.run_search)
        self.search_button.grid(row=2, column=0, sticky="ew")

        self.reload_button = ttk.Button(control_frame, text="Daten neu laden", command=self.reload_data)
        self.reload_button.grid(row=3, column=0, sticky="ew", pady=(8, 0))

        ttk.Separator(control_frame).grid(row=4, column=0, sticky="ew", pady=14)

        guidance = (
            "Die Tabelle zeigt vorhandene und wahrscheinliche Treffer. "
            "Alles unter 100 % ist ein Pruefvorschlag."
        )
        ttk.Label(control_frame, text=guidance, wraplength=310, justify="left").grid(row=5, column=0, sticky="w")

        result_frame = ttk.Frame(self, padding=(0, 12, 12, 12))
        result_frame.grid(row=0, column=1, sticky="nsew")
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(1, weight=1)
        result_frame.rowconfigure(3, weight=0)

        self.status_label = ttk.Label(result_frame, textvariable=self.status_var, style="Status.TLabel")
        self.status_label.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        self.tree = ttk.Treeview(result_frame, columns=DISPLAY_COLUMNS, show="headings", height=22)
        self.tree.grid(row=1, column=0, sticky="nsew")

        vertical_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        vertical_scrollbar.grid(row=1, column=1, sticky="ns")
        horizontal_scrollbar = ttk.Scrollbar(result_frame, orient="horizontal", command=self.tree.xview)
        horizontal_scrollbar.grid(row=2, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=vertical_scrollbar.set, xscrollcommand=horizontal_scrollbar.set)

        self._configure_columns()

        details_label = ttk.Label(result_frame, text="Details / Hinweise")
        details_label.grid(row=3, column=0, sticky="w", pady=(12, 4))

        self.details_text = tk.Text(result_frame, height=8, wrap="word", font=("Segoe UI", 9))
        self.details_text.grid(row=4, column=0, columnspan=2, sticky="nsew")
        self.details_text.configure(state="disabled")

    def _configure_columns(self) -> None:
        widths = {
            "Kandidat": 260,
            "Wahrscheinlichkeit": 116,
            "Status": 140,
            "Code": 70,
            "Hinweis": 560,
        }
        for column in DISPLAY_COLUMNS:
            self.tree.heading(column, text=column)
            anchor = "center" if column in {"Wahrscheinlichkeit", "Status", "Code"} else "w"
            width = widths.get(column, 132)
            stretch = column == "Hinweis"
            self.tree.column(column, width=width, minwidth=70, anchor=anchor, stretch=stretch)

    def set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.search_button.configure(state=state)
        self.reload_button.configure(state=state)
        if busy:
            self.status_var.set("Pruefung laeuft. Excel und Word-Daten werden nur gelesen...")

    def reload_data(self) -> None:
        clear_cache()
        self.status_var.set("Cache geleert. Die naechste Pruefung liest Excel/Word neu ein.")
        self.clear_results()

    def run_search(self) -> None:
        article_text = self.article_var.get().strip()
        self.set_busy(True)
        self.clear_results()

        thread = threading.Thread(target=self._search_worker, args=(article_text,), daemon=True)
        thread.start()

    def _search_worker(self, article_text: str) -> None:
        try:
            result = get_display_rows(article_text)
        except Exception as exc:  # GUI boundary: show error instead of crashing Tk.
            self.after(0, self._show_error, exc)
            return

        self.after(0, self._show_result, article_text, result)

    def _show_error(self, exc: Exception) -> None:
        self.set_busy(False)
        self.status_var.set("Fehler bei der Pruefung.")
        messagebox.showerror("Fehler", str(exc))

    def _show_result(self, article_text: str, result: dict) -> None:
        self.set_busy(False)
        self.clear_results()

        rows = result.get("rows", [])
        for row in rows:
            values = [row.get(column, "") for column in DISPLAY_COLUMNS]
            self.tree.insert("", "end", values=values)

        messages = result.get("messages", [])
        details = result.get("details", [])
        summary = f"{article_text}: {len(rows)} Zeilen angezeigt."
        if messages:
            summary = f"{summary} {messages[0]}"
        self.status_var.set(summary)

        detail_lines = []
        detail_lines.extend(messages)
        if messages and details:
            detail_lines.append("")
        detail_lines.extend(details)
        if rows:
            detail_lines.append("")
            detail_lines.append("Hinweis: 100 % bedeutet bestehender Excel-Eintrag. Alle niedrigeren Werte sind Pruefvorschlaege.")
        self.set_details("\n".join(detail_lines))

    def clear_results(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.set_details("")

    def set_details(self, text: str) -> None:
        self.details_text.configure(state="normal")
        self.details_text.delete("1.0", "end")
        if text:
            self.details_text.insert("1.0", text)
        self.details_text.configure(state="disabled")


def main() -> None:
    app = BolzenreihensetApp()
    app.mainloop()


if __name__ == "__main__":
    main()