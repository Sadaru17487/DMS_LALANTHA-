from django.conf import settings

def cache_version(request):
    return {
        'cache_version': settings.CACHE_VERSION
    }