from django.shortcuts import redirect  # type: ignore[reportMissingModuleSource]
from django.contrib import messages  # type: ignore[reportMissingModuleSource]
from .models import UserProfile

def check_permission(required_permission):
    def decorator(view_func):
        def wrapper(request, *args, **kwargs):
            if not request.user.is_authenticated:
                messages.error(request, 'Please login first!')
                return redirect('/login/')
            
            try:
                profile = request.user.profile
                if profile.has_permission(required_permission):
                    return view_func(request, *args, **kwargs)
                else:
                    messages.error(request, f'Permission denied. Required: {required_permission}')
                    return redirect('/reports/')
            except UserProfile.DoesNotExist:
                messages.error(request, 'User profile not found.')
                return redirect('/logout/')
        
        return wrapper
    return decorator