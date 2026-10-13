#!/usr/bin/env python
"""Django's command-line utility for administrative tasks."""
import os
import sys


def auto_migrate():
    """
    Automatically apply migrations on first run so users never see
    'no such table' errors after a fresh unzip.
    """
    try:
        from django.db import connection
        from django.core.management import call_command
        tables = connection.introspection.table_names()
        if 'barcode_app_hawbstate' not in tables:
            print("\n[HAWB Pro] First run — applying database migrations...")
            call_command('migrate', verbosity=1, interactive=False)
            print("[HAWB Pro] Done. Server starting...\n")
    except Exception as exc:
        # Non-fatal — user can still run `python manage.py migrate` manually
        print(f"[HAWB Pro] Auto-migrate skipped: {exc}")


def main():
    """Run administrative tasks."""
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'hawb_project.settings')
    try:
        from django.core.management import execute_from_command_line
    except ImportError as exc:
        raise ImportError(
            "Couldn't import Django.\n"
            "Install dependencies first:  pip install -r requirements.txt"
        ) from exc

    # Only auto-migrate when launching the development server
    cmd = sys.argv[1] if len(sys.argv) >= 2 else ''
    if cmd == 'runserver':
        import django
        django.setup()
        auto_migrate()

    execute_from_command_line(sys.argv)


if __name__ == '__main__':
    main()
