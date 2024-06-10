import os
import django

def create_superuser():
    # Set the DJANGO_SETTINGS_MODULE environment variable
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "core.settings")  # Use your actual settings module name

    # Initialize Django
    django.setup()

    # Import and run createsuperuser command
    from django.core.management import call_command
    call_command("createsuperuser")

if __name__ == "__main__":
    create_superuser()
