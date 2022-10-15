from django.apps import AppConfig


class BookConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'book'

    # 修改后台数据库名字，INSTALLED_APPS里面的app格式要为：book.apps.BookConfig
    verbose_name = 'tushu'
