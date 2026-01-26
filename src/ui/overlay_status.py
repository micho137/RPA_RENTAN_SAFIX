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
    etapa: str = ""
    extra: str = ""


class StatusOverlay:
    """
    Overlay always-on-top para mostrar estado sin bloquear el bot.
    Corre Tkinter en un thread dedicado y recibe updates por queue.
    """

    def __init__(
        self,
        title: str = "RENTAN - Procesando",
        topmost: bool = True,
        alpha: float = 0.92,
        x: int = 20,
        y: int = 20,
        width: int = 560,
        height: int = 130,
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

        self._root: Optional[tk.Tk] = None
        self._lbl_stage: Optional[tk.Label] = None
        self._lbl_doc: Optional[tk.Label] = None
        self._lbl_plate: Optional[tk.Label] = None
        self._lbl_extra: Optional[tk.Label] = None

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._run_tk, name="StatusOverlayThread", daemon=True)
        self._thread.start()
        self._started.wait(timeout=2.0)

    def stop(self) -> None:
        self._q.put(None)

    def update(self, document_id: str = "", placa: str = "", etapa: str = "", extra: str = "") -> None:
        self._q.put(StatusPayload(document_id=document_id, placa=placa, etapa=etapa, extra=extra))

    # ---------------- internals ----------------
    def _run_tk(self) -> None:
        root = tk.Tk()
        self._root = root

        root.title(self._title)
        root.overrideredirect(True)  # sin bordes
        root.attributes("-topmost", self._topmost)
        try:
            root.attributes("-alpha", self._alpha)
        except tk.TclError:
            pass

        root.geometry(f"{self._width}x{self._height}+{self._x}+{self._y}")

        frame = tk.Frame(root, bg="#111111", highlightthickness=2, highlightbackground="#444444")
        frame.pack(fill="both", expand=True)

        font_stage = ("Segoe UI", 14, "bold")
        font_line = ("Segoe UI", 12)
        font_small = ("Segoe UI", 10)

        self._lbl_stage = tk.Label(frame, text="Iniciando…", fg="white", bg="#111111", font=font_stage, anchor="w")
        self._lbl_stage.pack(fill="x", padx=14, pady=(10, 6))

        self._lbl_doc = tk.Label(frame, text="document_id: -", fg="#DADADA", bg="#111111", font=font_line, anchor="w")
        self._lbl_doc.pack(fill="x", padx=14)

        self._lbl_plate = tk.Label(frame, text="placa: -", fg="#DADADA", bg="#111111", font=font_line, anchor="w")
        self._lbl_plate.pack(fill="x", padx=14)

        self._lbl_extra = tk.Label(frame, text="", fg="#AAAAAA", bg="#111111", font=font_small, anchor="w")
        self._lbl_extra.pack(fill="x", padx=14, pady=(6, 10))

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

        # cerrar con ESC
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
        if not self._lbl_stage:
            return
        stage = p.etapa.strip() or "Procesando…"
        self._lbl_stage.config(text=stage)
        self._lbl_doc.config(text=f"document_id: {p.document_id or '-'}")
        self._lbl_plate.config(text=f"placa: {p.placa or '-'}")
        self._lbl_extra.config(text=p.extra or "")
