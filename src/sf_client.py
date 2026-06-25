"""Salesforce connection helper - shared by discover.py, import_lenders.py, the app."""

import configparser
import os
import sys

from simple_salesforce import Salesforce

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app_paths import data_file  # noqa: E402


def _config_path():
    """config.ini location: project folder in dev, per-user folder in the packaged app."""
    return data_file("config.ini")


def config_exists():
    return os.path.exists(_config_path())


def save_config(username, password, security_token, domain, api_name=""):
    """Write/replace config.ini (used by the app's Settings screen)."""
    cfg = configparser.ConfigParser()
    cfg["salesforce"] = {
        "username": username, "password": password,
        "security_token": security_token, "domain": domain or "login",
    }
    cfg["object"] = {"api_name": api_name}
    with open(_config_path(), "w") as fh:
        cfg.write(fh)


def load_config():
    path = _config_path()
    if not os.path.exists(path):
        sys.exit(
            "ERROR: config.ini not found.\n"
            "  Copy config.example.ini to config.ini and fill in your "
            "Salesforce login details first."
        )
    cfg = configparser.ConfigParser()
    cfg.read(path)
    return cfg


def connect(cfg=None):
    """Return a logged-in Salesforce client using config.ini credentials."""
    if cfg is None:
        cfg = load_config()
    sf_cfg = cfg["salesforce"]

    username = sf_cfg.get("username", "").strip()
    password = sf_cfg.get("password", "").strip()
    token = sf_cfg.get("security_token", "").strip()
    domain = sf_cfg.get("domain", "login").strip() or "login"

    if not username or not password:
        sys.exit("ERROR: username and password must be set in config.ini")

    kwargs = dict(username=username, password=password, security_token=token)

    # "login" / "test" are simple_salesforce shortcuts; a custom My Domain
    # is passed via instance_url-style domain.
    if domain in ("login", "test"):
        kwargs["domain"] = domain
    else:
        # Normalize a My Domain like "https://acme.my.salesforce.com/" -> "acme.my"
        # (simple_salesforce appends ".salesforce.com" itself).
        d = domain.replace("https://", "").replace("http://", "").rstrip("/")
        if d.endswith(".salesforce.com"):
            d = d[: -len(".salesforce.com")]
        kwargs["domain"] = d

    try:
        return Salesforce(**kwargs)
    except Exception as exc:  # noqa: BLE001 - surface a friendly message
        sys.exit(
            "ERROR: could not log in to Salesforce.\n"
            f"  {exc}\n"
            "  Check username, password, security_token and domain in config.ini."
        )


def get_object_api_name(cfg=None):
    if cfg is None:
        cfg = load_config()
    name = cfg["object"].get("api_name", "").strip()
    return name or None
