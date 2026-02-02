from __future__ import annotations

import threading
import queue
import time
import tkinter as tk
from dataclasses import dataclass
from typing import Optional


@dataclass
class StatusPayload:
    etapa: str = ""
    document_id: str = ""
    placa: str = ""
    current: int = 0
    total: int = 0
    extra: str = ""


def _fmt_hhmmss(seconds: float) -> str:
    s = max(0, int(seconds))
    hh = s // 3600
    mm = (s % 3600) // 60
    ss = s % 60
    return f"{hh:02d}:{mm:02d}:{ss:02d}"


class StatusOverlay:
    """
    Overlay always-on-top:
    - etapa (línea 1)
    - document_id
    - placa
    - progreso i/N + tiempo global (corre desde start(), aunque no haya updates)
    - extra (línea final opcional)
    """

    def __init__(
        self,
        title: str = "RENTAN - Procesando",
        topmost: bool = True,
        alpha: float = 0.92,
        x: int = 20,
        y: int = 20,
        width: int = 340,
        height: int = 165,
        poll_ms: int = 120,
        tick_ms: int = 250,
    ) -> None:
        self._title = title
        self._topmost = topmost
        self._alpha = alpha
        self._x = x
        self._y = y
        self._width = width
        self._height = height
        self._poll_ms = poll_ms
        self._tick_ms = tick_ms

        self._q: "queue.Queue[Optional[StatusPayload]]" = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._started = threading.Event()

        self._lbl_stage: Optional[tk.Label] = None
        self._lbl_doc: Optional[tk.Label] = None
        self._lbl_plate: Optional[tk.Label] = None
        self._lbl_meta: Optional[tk.Label] = None
        self._lbl_extra: Optional[tk.Label] = None

        self._t0 = 0.0
        self._last = StatusPayload(etapa="Procesando…", document_id="", placa="", current=0, total=0, extra="")

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._t0 = time.perf_counter()
        self._thread = threading.Thread(target=self._run_tk, name="StatusOverlayThread", daemon=True)
        self._thread.start()
        self._started.wait(timeout=2.0)

    def stop(self) -> None:
        self._q.put(None)

    def update(
        self,
        *,
        document_id: str = "",
        placa: str = "",
        current: int = 0,
        total: int = 0,
        etapa: str = "",
        extra: str = "",
        footer: str = "",
    ) -> None:
        if not extra and footer:
            extra = footer

        self._q.put(
            StatusPayload(
                etapa=etapa,
                document_id=document_id,
                placa=placa,
                current=current,
                total=total,
                extra=extra,
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

        font_stage = ("Segoe UI", 13, "bold")
        font_line = ("Segoe UI", 11)
        font_small = ("Segoe UI", 10)

        self._lbl_stage = tk.Label(frame, text=self._last.etapa, fg="white", bg="#111111", font=font_stage, anchor="w")
        self._lbl_stage.pack(fill="x", padx=12, pady=(10, 6))

        self._lbl_doc = tk.Label(frame, text="document_id: -", fg="#DADADA", bg="#111111", font=font_line, anchor="w")
        self._lbl_doc.pack(fill="x", padx=12, pady=(0, 2))

        self._lbl_plate = tk.Label(frame, text="placa: -", fg="#DADADA", bg="#111111", font=font_line, anchor="w")
        self._lbl_plate.pack(fill="x", padx=12, pady=(0, 2))

        self._lbl_meta = tk.Label(frame, text="progreso: - | tiempo: 00:00:00", fg="#BDBDBD", bg="#111111", font=font_small, anchor="w")
        self._lbl_meta.pack(fill="x", padx=12, pady=(6, 0))

        self._lbl_extra = tk.Label(frame, text="", fg="#AAAAAA", bg="#111111", font=font_small, anchor="w", wraplength=self._width - 24, justify="left")
        self._lbl_extra.pack(fill="x", padx=12, pady=(4, 10))

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
                    self._last = item
                    self._apply_static(self._last)
            except queue.Empty:
                pass
            root.after(self._poll_ms, poll_queue)

        def tick_timer():
            self._apply_meta(self._last)
            root.after(self._tick_ms, tick_timer)

        root.after(self._poll_ms, poll_queue)
        root.after(self._tick_ms, tick_timer)
        root.mainloop()

    def _apply_static(self, p: StatusPayload) -> None:
        if self._lbl_stage:
            self._lbl_stage.config(text=p.etapa or "Procesando…")
        if self._lbl_doc:
            self._lbl_doc.config(text=f"Factura: {p.document_id or '-'}")
        if self._lbl_plate:
            self._lbl_plate.config(text=f"Placa: {p.placa or '-'}")
        if self._lbl_extra:
            self._lbl_extra.config(text=p.extra or "")
        self._apply_meta(p)

    def _apply_meta(self, p: StatusPayload) -> None:
        if not self._lbl_meta:
            return
        prog = f"{p.current}/{p.total}" if p.total else "-"
        elapsed = (time.perf_counter() - self._t0) if self._t0 else 0.0
        t = _fmt_hhmmss(elapsed)
        self._lbl_meta.config(text=f"Progreso: {prog} | Tiempo: {t}")
