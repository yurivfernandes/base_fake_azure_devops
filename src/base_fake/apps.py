from django.apps import AppConfig


class BaseFakeConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'base_fake'
