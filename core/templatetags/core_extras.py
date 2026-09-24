from django import template

register = template.Library()



# core/templatetags/core_extras.py
@register.filter
def get_item(dictionary, key):
    if hasattr(dictionary, 'get'):
        return dictionary.get(key, False)
    return False

