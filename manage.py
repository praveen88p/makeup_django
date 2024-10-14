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

    # Get the port from the environment, default to 8000
    port = os.environ.get('PORT', '8000')

    # Automatically bind to 0.0.0.0:<PORT> when running the server
    if len(sys.argv) == 1 or sys.argv[1] == 'runserver':
        sys.argv = [sys.argv[0], '0.0.0.0:'+port]

    execute_from_command_line(sys.argv)

if __name__ == '__main__':
    main()
