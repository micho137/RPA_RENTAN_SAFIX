# src/ui/overlay_status.py
from __future__ import annotations

import threading
import queue
import time
import tkinter as tk
from dataclasses import dataclass
from typing import Optional


@dataclass
class StatusPayload:
    document_id: str = ""
    placa: str = ""
    current: int = 0
    total: int = 0
    elapsed_global_s: float = 0.0
    footer: str = ""  # opcional por si quieres mensajes cortos


def _fmt_hhmmss(seconds: float) -> str:
    s = max(0, int(seconds))
    hh = s // 3600
    mm = (s % 3600) // 60
    ss = s % 60
    return f"{hh:02d}:{mm:02d}:{ss:02d}"


class StatusOverlay:
    """
    Overlay always-on-top minimal:
    - document_id
    - placa
    - progreso i/N
    - tiempo global desde que aparece el overlay
    """

    def __init__(
        self,
        title: str = "RENTAN - Procesando",
        topmost: bool = True,
        alpha: float = 0.92,
        x: int = 20,
        y: int = 20,
        width: int = 560,
        height: int = 120,
        poll_ms: int = 120,
    ) -> None:
        self._title = title
        self._topmost = topmost
        self._alpha = alpha
        self._x = x
        self._y = y
        self._width = width
        self._height = height
        self._poll_ms = poll_ms

        self._q: "queue.Queue[Optional[StatusPayload]]" = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._started = threading.Event()

        self._lbl_doc: Optional[tk.Label] = None
        self._lbl_plate: Optional[tk.Label] = None
        self._lbl_meta: Optional[tk.Label] = None
        self._lbl_footer: Optional[tk.Label] = None

        self._t0 = time.perf_counter()  # tiempo global inicia al crear el objeto

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        # tiempo global inicia al mostrar overlay
        self._t0 = time.perf_counter()
        self._thread = threading.Thread(target=self._run_tk, name="StatusOverlayThread", daemon=True)
        self._thread.start()
        self._started.wait(timeout=2.0)

    def stop(self) -> None:
        self._q.put(None)

    def update(
        self,
        document_id: str = "",
        placa: str = "",
        current: int = 0,
        total: int = 0,
        footer: str = "",
    ) -> None:
        elapsed = time.perf_counter() - self._t0
        self._q.put(
            StatusPayload(
                document_id=document_id,
                placa=placa,
                current=current,
                total=total,
                elapsed_global_s=elapsed,
                footer=footer,
            )
        )

    # ---------------- internals ----------------
    def _run_tk(self) -> None:
        root = tk.Tk()
        root.title(self._title)
        root.overrideredirect(True)
        root.attributes("-topmost", self._topmost)
        try:
            root.attributes("-alpha", self._alpha)
        except tk.TclError:
            pass

        root.geometry(f"{self._width}x{self._height}+{self._x}+{self._y}")

        frame = tk.Frame(root, bg="#111111", highlightthickness=2, highlightbackground="#444444")
        frame.pack(fill="both", expand=True)

        font_line = ("Segoe UI", 12, "bold")
        font_mid = ("Segoe UI", 12)
        font_small = ("Segoe UI", 10)

        self._lbl_doc = tk.Label(frame, text="document_id: -", fg="white", bg="#111111", font=font_line, anchor="w")
        self._lbl_doc.pack(fill="x", padx=14, pady=(10, 2))

        self._lbl_plate = tk.Label(frame, text="placa: -", fg="#DADADA", bg="#111111", font=font_mid, anchor="w")
        self._lbl_plate.pack(fill="x", padx=14, pady=(0, 2))

        self._lbl_meta = tk.Label(frame, text="progreso: - | tiempo: 00:00:00", fg="#BDBDBD", bg="#111111", font=font_small, anchor="w")
        self._lbl_meta.pack(fill="x", padx=14, pady=(6, 0))

        self._lbl_footer = tk.Label(frame, text="", fg="#AAAAAA", bg="#111111", font=font_small, anchor="w")
        self._lbl_footer.pack(fill="x", padx=14, pady=(4, 10))

        # mover overlay con drag
        def _start_move(event):
            root._drag_start_x = event.x
            root._drag_start_y = event.y

        def _do_move(event):
            nx = root.winfo_x() + (event.x - root._drag_start_x)
            ny = root.winfo_y() + (event.y - root._drag_start_y)
            root.geometry(f"+{nx}+{ny}")

        frame.bind("<Button-1>", _start_move)
        frame.bind("<B1-Motion>", _do_move)

        root.bind("<Escape>", lambda e: root.destroy())

        self._started.set()

        def poll_queue():
            try:
                while True:
                    item = self._q.get_nowait()
                    if item is None:
                        root.destroy()
                        return
                    self._apply(item)
            except queue.Empty:
                pass
            root.after(self._poll_ms, poll_queue)

        root.after(self._poll_ms, poll_queue)
        root.mainloop()

    def _apply(self, p: StatusPayload) -> None:
        self._lbl_doc.config(text=f"ID Documento: {p.document_id or '-'}")
        self._lbl_plate.config(text=f"Placa: {p.placa or '-'}")

        prog = f"{p.current}/{p.total}" if p.total else "-"
        t = _fmt_hhmmss(p.elapsed_global_s)
        self._lbl_meta.config(text=f"Progreso: {prog} | tiempo: {t}")

        self._lbl_footer.config(text=p.footer or "")
