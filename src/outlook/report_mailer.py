from __future__ import annotations

from pathlib import Path
from typing import Iterable

import win32com.client

from src.config import settings
from src.core.com_init import com_initialized


def send_report_email(*, attachments: Iterable[Path], logger, subject_suffix: str = "") -> bool:
    """
    Envia correo de reporte con Outlook Desktop via COM.
    Retorna True si se envio, False si no aplica o falla.
    """
    if not settings.report_email_enabled:
        return False

    to_value = str(settings.report_email_to or "").strip()
    if not to_value:
        logger.warning("[MAIL] REPORT_EMAIL_ENABLED=true pero REPORT_EMAIL_TO vacio. Se omite envio.")
        return False

    files = [Path(p).resolve() for p in attachments if p and Path(p).exists()]
    if not files:
        logger.warning("[MAIL] No hay adjuntos disponibles para enviar reporte.")
        return False

    subject = settings.report_email_subject
    if subject_suffix:
        subject = f"{subject} | {subject_suffix}"

    with com_initialized():
        try:
            app = win32com.client.Dispatch("Outlook.Application")
            mail = app.CreateItem(0)
            mail.To = to_value
            if settings.report_email_cc:
                mail.CC = settings.report_email_cc
            mail.Subject = subject
            mail.Body = settings.report_email_body
            for f in files:
                mail.Attachments.Add(str(f))
            mail.Send()
            logger.info("[MAIL] Reporte enviado a %s con %s adjuntos.", to_value, len(files))
            return True
        except Exception as ex:
            logger.error("[MAIL] Error enviando reporte: %s", ex)
            return False
