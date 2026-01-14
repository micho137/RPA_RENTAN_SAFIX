import os
import threading
from datetime import datetime, timedelta
from pathlib import Path

import flet as ft

from src.workflows.invoice_pipeline import run_pipeline

DATE_FMT = "%d/%m/%Y"


def main(page: ft.Page):
    page.title = "Query Parameters"
    page.padding = 20
    page.scroll = ft.ScrollMode.AUTO

    # ---- Window sizing (cross-version best-effort) ----
    target_w, target_h = 720, 570

    # Old API (some versions)
    try:
        page.window.resizable = False
        page.window_width = target_w
        page.window_height = target_h
    except Exception:
        pass

    # New API (some versions)
    if hasattr(page, "window") and page.window is not None:
        try:
            page.window.width = target_w
            page.window.height = target_h
            page.window.resizable = False
        except Exception:
            pass

    # ===============================
    # FILE PICKER (Excel)
    # ===============================
    selected_file_txt = ft.Text(value="", selectable=True)

    def on_file_picked(e: ft.FilePickerResultEvent):
        if e.files:
            selected_file_txt.value = e.files[0].path
        else:
            selected_file_txt.value = ""
        refresh()
        page.update()

    file_picker = ft.FilePicker(on_result=on_file_picked)
    page.overlay.append(file_picker)

    def open_file_picker(_):
        file_picker.pick_files(
            allow_multiple=False,
            allowed_extensions=["xlsx", "xls"],
        )

    # --- Helpers ---
    def parse_int(value: str, default: int = 0) -> int:
        try:
            v = int((value or "").strip())
            return max(0, v)
        except Exception:
            return default

    def compute_range(days_back: int) -> tuple[str, str]:
        today = datetime.now().date()
        start = today - timedelta(days=days_back)
        return start.strftime(DATE_FMT), today.strftime(DATE_FMT)

    def render_env(only_unread: bool, days_back: int, mark_as_read: bool, move_to_processed: bool) -> str:
        return (
            f"ONLY_UNREAD={'true' if only_unread else 'false'}\n"
            f"DAYS_BACK={days_back}\n"
            f"MARK_AS_READ={'true' if mark_as_read else 'false'}\n"
            f"MOVE_TO_PROCESSED={'true' if move_to_processed else 'false'}"
        )

    # --- Controls (tus parámetros originales) ---
    days_input = ft.TextField(
        label="DAYS_BACK",
        value="0",
        width=260,
        keyboard_type=ft.KeyboardType.NUMBER,
    )

    only_unread_cb = ft.Checkbox(label="ONLY_UNREAD", value=False)
    mark_as_read_cb = ft.Checkbox(label="MARK_AS_READ", value=False)
    move_to_processed_cb = ft.Checkbox(label="MOVE_TO_PROCESSED", value=False)

    start_date_txt = ft.Text(value="-", selectable=True)
    end_date_txt = ft.Text(value="-", selectable=True)

    env_preview = ft.TextField(
        label="Preview",
        value="",
        multiline=True,
        min_lines=5,
        max_lines=6,
        read_only=True,
        expand=True,
    )

    status = ft.Text(value="", color=ft.Colors.RED_700)

    # ---- Botones (compatibles con versiones viejas: texto posicional) ----
    load_excel_btn = ft.ElevatedButton(
        "Load Excel file",
        icon=ft.Icons.UPLOAD_FILE,
        on_click=open_file_picker,
    )

    run_btn = ft.ElevatedButton(
        "Run Pipeline + SAFIX",
        icon=ft.Icons.PLAY_ARROW,
    )
    run_btn.disabled = True

    def set_status(msg: str, is_error: bool = False):
        status.value = msg
        status.color = ft.Colors.RED_700 if is_error else ft.Colors.GREEN_700
        page.update()

    def refresh(_=None):
        days_back = parse_int(days_input.value, default=0)
        start_s, end_s = compute_range(days_back)

        start_date_txt.value = start_s
        end_date_txt.value = end_s

        env_preview.value = render_env(
            only_unread=bool(only_unread_cb.value),
            days_back=days_back,
            mark_as_read=bool(mark_as_read_cb.value),
            move_to_processed=bool(move_to_processed_cb.value),
        )

        # habilitar RUN si hay Excel y no está vacío
        run_btn.disabled = not bool((selected_file_txt.value or "").strip())

        page.update()

    def validate_days(_e):
        _ = parse_int(days_input.value, default=0)
        refresh()

    days_input.on_change = validate_days
    only_unread_cb.on_change = refresh
    mark_as_read_cb.on_change = refresh
    move_to_processed_cb.on_change = refresh

    # ===============================
    # EJECUCIÓN PIPELINE + SAFIX
    # ===============================
    def apply_env_from_ui():
        """
        Setea variables de entorno para que tu workflow actual (run_download/settings)
        las lea sin tener que refactorizar módulos.
        """
        days_back = parse_int(days_input.value, default=0)

        os.environ["ONLY_UNREAD"] = "true" if bool(only_unread_cb.value) else "false"
        os.environ["DAYS_BACK"] = str(days_back)
        os.environ["MARK_AS_READ"] = "true" if bool(mark_as_read_cb.value) else "false"
        os.environ["MOVE_TO_PROCESSED"] = "true" if bool(move_to_processed_cb.value) else "false"

    def run_job(excel_path_str: str):
        try:
            set_status("Aplicando parámetros y ejecutando Pipeline (download/extract/aggregate)...", is_error=False)

            apply_env_from_ui()

            # Nota: run_pipeline ya ejecuta SAFIX al final con excel_path (según el ajuste que hicimos)
            run_pipeline(
                output_dir=Path("./output"),
                lang="spa",
                dpi=300,
                aggregate_by_id=True,
                excel_path=Path(excel_path_str),
                run_safix=True,
            )

            set_status("Finalizado OK. Pipeline + SAFIX ejecutados.", is_error=False)

        except Exception as ex:
            set_status(f"Error: {ex}", is_error=True)
        finally:
            run_btn.disabled = False
            page.update()

    def on_run_click(_):
        excel_path_str = (selected_file_txt.value or "").strip()
        if not excel_path_str:
            set_status("Selecciona un Excel primero.", is_error=True)
            return

        p = Path(excel_path_str)
        if not p.exists():
            set_status("El archivo seleccionado no existe.", is_error=True)
            return

        run_btn.disabled = True
        page.update()

        t = threading.Thread(target=run_job, args=(excel_path_str,), daemon=True)
        t.start()

    run_btn.on_click = on_run_click

    # --- Layout (tu layout original + botones al final) ---
    page.add(
        ft.Column(
            spacing=14,
            controls=[
                ft.Text("Query configuration", size=20, weight=ft.FontWeight.BOLD),
                ft.Row(
                    spacing=16,
                    vertical_alignment=ft.CrossAxisAlignment.START,
                    controls=[
                        ft.Container(
                            content=ft.Column(
                                spacing=10,
                                controls=[
                                    days_input,
                                    ft.Divider(),
                                    only_unread_cb,
                                    mark_as_read_cb,
                                    move_to_processed_cb,
                                ],
                            ),
                            padding=16,
                            border_radius=10,
                            border=ft.border.all(1, ft.Colors.GREY_300),
                            width=320,
                        ),
                        ft.Container(
                            content=ft.Column(
                                spacing=10,
                                controls=[
                                    ft.Text("Date range", weight=ft.FontWeight.W_600),
                                    ft.Row(
                                        controls=[
                                            ft.Text("From:", weight=ft.FontWeight.W_600),
                                            start_date_txt,
                                        ]
                                    ),
                                    ft.Row(
                                        controls=[
                                            ft.Text("To:", weight=ft.FontWeight.W_600),
                                            end_date_txt,
                                        ]
                                    ),
                                    ft.Divider(),
                                    env_preview,
                                ],
                            ),
                            padding=16,
                            border_radius=10,
                            border=ft.border.all(1, ft.Colors.GREY_300),
                            expand=True,
                        ),
                    ],
                ),

                # --- Excel loader + Run at bottom ---
                ft.Divider(),
                load_excel_btn,
                selected_file_txt,
                run_btn,
                status,
            ],
        )
    )

    refresh()


if __name__ == "__main__":
    ft.app(target=main, view=ft.AppView.FLET_APP)
