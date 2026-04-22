"""Module auto-discovery for the RCW Processing Suite.

Walks ``app.modules`` at startup and calls each subpackage's ``register(app)``
hook. Modules can optionally expose ``MODULE_META`` (dict with id/name/
description) which is collected and returned so the UI can render tabs
without the core app knowing anything about individual modules.
"""
from __future__ import annotations

import importlib
import logging
import pkgutil
from typing import Any

logger = logging.getLogger(__name__)


def load_modules(app) -> list[dict[str, Any]]:
    """Discover every subpackage under ``app.modules`` and register it.

    Returns a list of MODULE_META dicts for modules that provide one.
    """
    modules_pkg = importlib.import_module("app.modules")
    registered: list[dict[str, Any]] = []

    for info in pkgutil.iter_modules(modules_pkg.__path__):
        if not info.ispkg:
            continue
        mod_name = f"app.modules.{info.name}"
        try:
            module = importlib.import_module(mod_name)
        except Exception as exc:
            logger.exception("Failed to import module %s: %s", mod_name, exc)
            continue

        register = getattr(module, "register", None)
        if register is None:
            logger.warning("Module %s has no register(app) function; skipping", mod_name)
            continue

        try:
            register(app)
        except Exception as exc:
            logger.exception("Module %s failed to register: %s", mod_name, exc)
            continue

        meta = getattr(module, "MODULE_META", None)
        if meta:
            registered.append(meta)
            logger.info("Registered module: %s", meta.get("id", info.name))
        else:
            logger.info("Registered module: %s (no meta)", info.name)

    return registered
