from django import template

register = template.Library()


@register.filter
def get_item(dictionary, key):
    
    """Get an item from a dictionary by key."""
    return dictionary.get(key, 0)

@register.filter
def has_perm(profile, permission):
    """Check if a user profile has a specific permission."""
    try:
        return profile.has_permission(permission)
    except Exception:
        return False

