from django.apps import AppConfig  # type: ignore[import-not-found]

class CoreConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'core'