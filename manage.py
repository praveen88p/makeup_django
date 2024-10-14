import os
import sys

def main():
    """Run administrative tasks."""
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'seating_chart_project.settings')
    try:
        from django.core.management import execute_from_command_line
    except ImportError as exc:
        raise ImportError(
            "Couldn't import Django. Are you sure it's installed and "
            "available on your PYTHONPATH environment variable? Did you "
            "forget to activate a virtual environment?"
        ) from exc

    # Set the default port if none is provided
    port = os.environ.get('PORT', '8000')

    # Update the command to use the port if 'runserver' is the command
    if len(sys.argv) == 1 or sys.argv[1] == 'runserver':
        sys.argv += ['0.0.0.0', port]

    execute_from_command_line(sys.argv)

if __name__ == '__main__':
    main()
