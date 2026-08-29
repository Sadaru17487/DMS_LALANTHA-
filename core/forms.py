from django import forms
from .models import Product, Vehicle, SalesBill, Employee


class VehicleLoadForm(forms.Form):
    vehicle = forms.ModelChoiceField(
        queryset=Vehicle.objects.filter(is_active=True), 
        label="Select Vehicle / Driver",
        widget=forms.Select(attrs={'class': 'form-control'})
    )    


class SalesBillForm(forms.ModelForm):
    # ✅ Override the field to make it optional
    discount_total = forms.DecimalField(required=False, initial=0, max_digits=12, decimal_places=2)

    class Meta:
        model = SalesBill
        fields = ['vehicle', 'rep', 'shop_name', 'shop_code', 'invoice_no', 'discount_total']
        widgets = {
            'vehicle': forms.Select(attrs={'class': 'form-control'}),
            'shop_name': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g., ABC Supermarket'}),
            'shop_code': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Optional code'}),
            'invoice_no': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g., INV-001'}),
            # ✅ Remove discount_total from widgets since we override it
        }


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


class ProductForm(forms.ModelForm):
    class Meta:
        model = Product
        fields = [
            'name', 
            'code', 
            'category', 
            'unit', 
            'cost_price', 
            'selling_price',
            'description',  # ✅ ADD THIS
            'notes',        # ✅ ADD THIS
            'is_active'
        ]
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Enter product name'}),
            'code': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Enter product code'}),
            'category': forms.Select(attrs={'class': 'form-control'}),
            'unit': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g., Pcs, Kg, Ltr'}),
            'cost_price': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.01', 'placeholder': '0.00'}),
            'selling_price': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.01', 'placeholder': '0.00'}),
            'description': forms.Textarea(attrs={'rows': 3, 'class': 'form-control', 'placeholder': 'Product description...'}),
            'notes': forms.Textarea(attrs={'rows': 3, 'class': 'form-control', 'placeholder': 'Additional notes...'}),
            'is_active': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        }
        labels = {
            'name': 'Product Name',
            'code': 'Product Code',
            'category': 'Category',
            'unit': 'Unit',
            'cost_price': 'Cost Price (Rs)',
            'selling_price': 'Selling Price (Rs)',
            'description': 'Description',
            'notes': 'Notes',
            'is_active': 'Active',
        }