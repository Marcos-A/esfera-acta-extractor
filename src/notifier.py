"""
Failure notification helpers.
"""

from __future__ import annotations

import json
import os
import smtplib
import ssl
import urllib.parse
import urllib.request
from email.message import EmailMessage
from typing import Callable


def notify_failure(subject: str, body: str) -> None:
    """
    Best-effort notification.
    Tries Telegram, SMTP, and generic webhook delivery when configured.
    """
    telegram_sent = _safe_send("Telegram", _send_telegram_message, subject, body)
    smtp_sent = _safe_send("SMTP", _send_smtp_message, subject, body)
    webhook_sent = _safe_send("webhook", _send_webhook_message, subject, body)
    if not telegram_sent and not smtp_sent and not webhook_sent:
        print("Notification skipped: no Telegram, SMTP, or webhook configuration found.")


def _safe_send(
    channel_name: str,
    sender: Callable[[str, str], bool],
    subject: str,
    body: str,
) -> bool:
    try:
        return sender(subject, body)
    except Exception as exc:
        print(f"{channel_name} notification failed: {exc}")
        return False


def _send_telegram_message(subject: str, body: str) -> bool:
    """Send a plain-text Telegram alert when bot credentials are configured."""
    bot_token = os.getenv("TELEGRAM_BOT_TOKEN")
    chat_id = os.getenv("TELEGRAM_CHAT_ID")
    if not bot_token or not chat_id:
        return False

    message_text = f"{subject}\n\n{body}"
    payload = urllib.parse.urlencode(
        {
            "chat_id": chat_id,
            "text": message_text,
            "disable_web_page_preview": "true",
        }
    ).encode("utf-8")
    request = urllib.request.Request(
        f"https://api.telegram.org/bot{bot_token}/sendMessage",
        data=payload,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=10):
        return True


def _send_smtp_message(subject: str, body: str) -> bool:
    """Send an email alert using either implicit SSL or STARTTLS, depending on settings."""
    host = os.getenv("SMTP_HOST")
    port = os.getenv("SMTP_PORT")
    username = os.getenv("SMTP_USERNAME")
    password = os.getenv("SMTP_PASSWORD")
    sender = os.getenv("ALERT_FROM_EMAIL")
    recipient = os.getenv("ALERT_TO_EMAIL")

    if not all([host, port, sender, recipient]):
        return False

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = sender
    message["To"] = recipient
    message.set_content(body)

    port_number = int(port)
    if os.getenv("SMTP_USE_SSL", "0") == "1":
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(host, port_number, context=context) as server:
            if username and password:
                server.login(username, password)
            server.send_message(message)
        return True

    with smtplib.SMTP(host, port_number) as server:
        if os.getenv("SMTP_USE_TLS", "1") == "1":
            server.starttls(context=ssl.create_default_context())
        if username and password:
            server.login(username, password)
        server.send_message(message)
    return True


def _send_webhook_message(subject: str, body: str) -> bool:
    """Send a simple JSON payload to any generic webhook endpoint."""
    webhook_url = os.getenv("ALERT_WEBHOOK_URL")
    if not webhook_url:
        return False

    payload = json.dumps({"text": f"{subject}\n\n{body}"}).encode("utf-8")
    request = urllib.request.Request(
        webhook_url,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=10):
        return True
