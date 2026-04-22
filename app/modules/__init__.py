"""Feature modules for the RCW Processing Suite.

Each subpackage is a self-contained module. To add a new module:

1. Create a new package under ``app/modules/<name>/``.
2. In its ``__init__.py``, expose ``register(app)`` that includes the module's
   router (and optionally mounts static/templates).
3. Set ``MODULE_META`` with id/name/description for the UI.

The loader in ``app.core.registry`` auto-discovers every subpackage and calls
its ``register`` function — no changes to ``main.py`` are required.
"""
