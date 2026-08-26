from django import forms
from .models import Vehicle, SalesBill, Employee


class VehicleLoadForm(forms.Form):
    vehicle = forms.ModelChoiceField(
        queryset=Vehicle.objects.filter(is_active=True), 
        label="Select Vehicle / Driver",
        widget=forms.Select(attrs={'class': 'form-control'})
    )    


class SalesBillForm(forms.ModelForm):
    class Meta:
        model = SalesBill
        fields = ['vehicle', 'date', 'invoice_no', ...]
        widgets = {
            'date': forms.DateInput(attrs={'type': 'date'}),
        }
    
    def clean_date(self):
        date = self.cleaned_data.get('date')
        # Allow any date (past, present, future)
        # Remove any validation that restricts past dates
        return date


class EmployeeForm(forms.ModelForm):
    class Meta:
        model = Employee
        fields = ['name', 'position', 'phone', 'email', 'address', 'hire_date', 'salary', 'notes', 'is_active', 'rep_code']
        widgets = {
            'rep_code': forms.TextInput(attrs={'placeholder': 'e.g., 008SIN'}),
        }
    
    def clean_rep_code(self):
        position = self.cleaned_data.get('position')
        rep_code = self.cleaned_data.get('rep_code')
        
        if position == 'Rep':
            if not rep_code:
                raise forms.ValidationError('Rep Code is required for Sales Representatives.')
            # Check uniqueness (case-insensitive)
            if Employee.objects.filter(rep_code__iexact=rep_code).exclude(pk=self.instance.pk).exists():
                raise forms.ValidationError('This Rep Code is already in use.')
        else:
            # If not Rep, rep_code should be empty
            if rep_code:
                raise forms.ValidationError('Rep Code is only allowed for Sales Representatives.')
        return rep_code        

    