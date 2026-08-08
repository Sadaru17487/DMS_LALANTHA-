from django.shortcuts import redirect
from django.contrib import messages
from .models import UserProfile

def permission_required(permission):
    def decorator(view_func):
        def wrapper(request, *args, **kwargs):
            if not request.user.is_authenticated:
                messages.error(request, 'Please login first.')
                return redirect('/login/')  # ✅ Always returns
            
            try:
                profile = request.user.profile
                if profile.has_permission(permission):
                    return view_func(request, *args, **kwargs)  # ✅ Returns view response
                else:
                    messages.error(request, f'You do not have permission to access this page. Required: {permission}')
                    return redirect('/reports/')  # ✅ Always returns
            except UserProfile.DoesNotExist:
                messages.error(request, 'User profile not found. Please contact admin.')
                return redirect('/logout/')  # ✅ Always returns
        
        return wrapper
    return decorator