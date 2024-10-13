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

    # Check if the PORT environment variable is set, and use it if available
    port = os.environ.get('PORT', '8000')  # Default to port 8000 if PORT isn't set
    sys.argv = sys.argv[:1] + ['runserver', '0.0.0.0:' + port] + sys.argv[1:]

    execute_from_command_line(sys.argv)

if __name__ == '__main__':
    main()
