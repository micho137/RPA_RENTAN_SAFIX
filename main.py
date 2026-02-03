def main():
    import flet as ft
    from src.config import settings
    from src.core.logging_config import init_logging, redirect_std_streams
    from src.ui.ui_flet import main as ui_main

    logger = init_logging(settings.log_dir)
    redirect_std_streams(logger)

    ft.app(target=ui_main, view=ft.AppView.FLET_APP)


if __name__ == "__main__":
    main()
