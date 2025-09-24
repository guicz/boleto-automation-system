"""Notification helper for boleto automation."""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Dict, Optional

import requests


class WebhookNotifier:
    """Sends boleto notifications through a configured webhook (n8n)."""

    def __init__(
        self,
        webhook_url: str,
        method: str,
        headers: Optional[Dict[str, str]],
        message_template: str,
        logger: logging.Logger,
        timeout: int = 30,
    ) -> None:
        self.webhook_url = webhook_url
        self.method = method.upper()
        self.headers = headers or {"Content-Type": "application/json"}
        self.message_template = message_template
        self.logger = logger
        self.timeout = timeout

    def send_notification(
        self,
        phone_number: str,
        nome: str,
        grupo: str,
        cota: str,
        pdf_path: Path,
        file_url: Optional[str],
        drive_file_id: Optional[str] = None,
    ) -> bool:
        message = self.message_template.format(nome=nome, grupo=grupo, cota=cota)
        payload = {
            "phone": phone_number,
            "nome": nome,
            "grupo": grupo,
            "cota": cota,
            "message": message,
            "file_name": pdf_path.name,
            "file_url": file_url,
            "drive_file_id": drive_file_id,
        }

        if not file_url:
            self.logger.error("Notification skipped: no file URL available for %s", pdf_path.name)
            return False

        try:
            if self.headers.get("Content-Type", "").startswith("application/json"):
                data = json.dumps(payload)
                response = requests.request(
                    self.method,
                    self.webhook_url,
                    data=data,
                    headers=self.headers,
                    timeout=self.timeout,
                )
            else:
                response = requests.request(
                    self.method,
                    self.webhook_url,
                    data=payload,
                    headers=self.headers,
                    timeout=self.timeout,
                )

            if response.status_code >= 200 and response.status_code < 300:
                self.logger.info(
                    "Notification sent successfully for %s (status %s)",
                    phone_number,
                    response.status_code,
                )
                return True

            self.logger.error(
                "Notification webhook failed (%s): %s", response.status_code, response.text
            )
        except Exception as error:
            self.logger.error("Notification webhook error: %s", error)

        return False
