from datetime import date, datetime

from django.db import models  
from django.contrib.auth.models import User  
from django.db.models import Sum  
from decimal import Decimal



class Category(models.Model):
    name = models.CharField(max_length=100, unique=True)
    description = models.TextField(blank=True, null=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    
    def __str__(self):
        return self.name
    
    class Meta:
        ordering = ['name']
        verbose_name_plural = "Categories"


class Product(models.Model):
    UNIT_CHOICES = [
        ('Pcs', 'Pieces'),
        ('Kg', 'Kilogram'),
        ('Ltr', 'Liter'),
        ('Ml', 'Milliliter'),
        ('Box', 'Box'),
        ('Carton', 'Carton'),
        ('Pack', 'Pack'),
        ('Bag', 'Bag'),
        ('Bottle', 'Bottle'),
        ('Can', 'Can'),
        ('Other', 'Other'),
    ]
    
    
    name = models.CharField(max_length=200)
    code = models.CharField(max_length=50, unique=True)
    category = models.ForeignKey(Category, on_delete=models.SET_NULL, null=True, blank=True, related_name='products')
    unit = models.CharField(max_length=20, choices=UNIT_CHOICES, default='Pcs')
    selling_price = models.DecimalField(max_digits=10, decimal_places=2)
    cost_price = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)
    notes = models.TextField(blank=True, null=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    def get_warehouse_stock(self):
        """Get current warehouse stock as Decimal"""
        try:
            return self.stock.quantity
        except WarehouseStock.DoesNotExist:
            return Decimal('0')

    def get_vehicle_stock(self):
        """Get total stock across all vehicles as Decimal"""
        result = VehicleStock.objects.filter(product=self).aggregate(total=Sum('quantity'))
        total = result.get('total')
        return total if total is not None else Decimal('0')

    def get_total_stock(self):
        """Get total stock (warehouse + all vehicles) as Decimal"""
        return self.get_warehouse_stock() + self.get_vehicle_stock()

    def get_profit_margin(self):
        """Calculate profit margin percentage"""
        if self.cost_price and self.cost_price > 0:
            return ((self.selling_price - self.cost_price) / self.cost_price) * 100
        return None


class WarehouseStock(models.Model):
    """Tracks physical stock in the warehouse"""
    product = models.OneToOneField(Product, on_delete=models.CASCADE, related_name='stock')
    quantity = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    last_updated = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.product.name}: {self.quantity} {self.product.unit}"


class Vehicle(models.Model):
    VEHICLE_TYPES = [
        ('Lorry', 'Lorry'),
        ('Van', 'Van'),
        ('Truck', 'Truck'),
        ('Pickup', 'Pickup'),
        ('Tractor', 'Tractor'),
        ('Other', 'Other'),
    ]
    
    FUEL_TYPES = [
        ('Petrol', 'Petrol'),
        ('Diesel', 'Diesel'),
        ('Electric', 'Electric'),
        ('Hybrid', 'Hybrid'),
    ]
    
    STATUS_CHOICES = [
        ('ACTIVE', 'Active'),
        ('INACTIVE', 'Inactive'),
        ('MAINTENANCE', 'Under Maintenance'),
    ]
    
    # Basic Info
    vehicle_number = models.CharField(max_length=50, unique=True, help_text="e.g., LC-3634")
    vehicle_type = models.CharField(max_length=20, choices=VEHICLE_TYPES, default='Lorry')
    driver_name = models.CharField(max_length=100, help_text="Current driver name")
    is_active = models.BooleanField(default=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='ACTIVE')
    
    # Registration
    registration_number = models.CharField(max_length=50, blank=True, null=True)
    registration_expiry = models.DateField(null=True, blank=True)
    
    # Insurance
    insurance_expiry = models.DateField(null=True, blank=True)
    
    # Maintenance
    last_service_date = models.DateField(null=True, blank=True)
    mileage = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True, help_text="Current odometer reading")
    
    # Details
    fuel_type = models.CharField(max_length=20, choices=FUEL_TYPES, default='Diesel')
    capacity = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True, help_text="Weight capacity in tons or volume in liters")
    notes = models.TextField(blank=True, null=True)
    
    # Purchase Info
    purchase_date = models.DateField(null=True, blank=True)
    purchase_price = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    current_value = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True, help_text="Optional depreciated value")
    
    # Timestamps
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    def __str__(self):
        return f"{self.vehicle_number} - {self.driver_name}"
    
    def get_total_expenses(self):
        """Get total expenses for this vehicle"""
        total = Expense.objects.filter(vehicle=self).aggregate(total=Sum('amount'))['total']
        return total or 0
    
    def get_total_sales(self):
        """Get total sales from this vehicle"""
        total = SalesBill.objects.filter(vehicle=self).aggregate(total=Sum('net_total'))['total']
        return total or 0
    
    def get_net_performance(self):
        """Get net performance (sales - expenses)"""
        return self.get_total_sales() - self.get_total_expenses()
    
    def is_insurance_expiring_soon(self, days=30):
        """Check if insurance expires within the specified days"""
        if self.insurance_expiry:
            from datetime import date, timedelta
            return self.insurance_expiry - date.today() <= timedelta(days=days) and self.insurance_expiry >= date.today()
        return False
    
    def is_registration_expiring_soon(self, days=30):
        """Check if registration expires within the specified days"""
        if self.registration_expiry:
            from datetime import date, timedelta
            return self.registration_expiry - date.today() <= timedelta(days=days) and self.registration_expiry >= date.today()
        return False
    
    def get_status_badge(self):
        """Get status badge class"""
        if self.status == 'ACTIVE':
            return 'complete'
        elif self.status == 'MAINTENANCE':
            return 'warning'
        else:
            return 'pending'
    
    class Meta:
        ordering = ['vehicle_number']


class VehicleStock(models.Model):
    """Floating stock assigned to each vehicle"""
    vehicle = models.ForeignKey(Vehicle, on_delete=models.CASCADE, related_name='stock_items')
    product = models.ForeignKey(Product, on_delete=models.CASCADE)
    quantity = models.DecimalField(max_digits=15, decimal_places=2, default=0)

    class Meta:
        unique_together = ('vehicle', 'product')

    def __str__(self):
        return f"{self.vehicle.vehicle_number} - {self.product.name}: {self.quantity}"


class SalesBill(models.Model):
    """Header of the sales bill"""
    vehicle = models.ForeignKey(Vehicle, on_delete=models.CASCADE, related_name='bills')
    date = models.DateField(auto_now_add=True)
    shop_name = models.CharField(max_length=255)
    shop_code = models.CharField(max_length=50, blank=True, null=True)
    invoice_no = models.CharField(max_length=100, unique=True)
    subtotal = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    discount_total = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    net_total = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.invoice_no} - {self.shop_name} ({self.vehicle.vehicle_number})"
    rep = models.ForeignKey('Employee', on_delete=models.SET_NULL, null=True, blank=True, related_name='sales_bills')
    status = models.CharField(max_length=20, choices=[
        ('DRAFT', 'Draft'),
        ('COMPLETED', 'Completed'),
        ('CANCELLED', 'Cancelled'),
    ], default='DRAFT')
    completed_at = models.DateTimeField(null=True, blank=True)
    completed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='completed_bills')
    
    # Return tracking
    is_return = models.BooleanField(default=False)
    original_bill = models.ForeignKey('self', on_delete=models.SET_NULL, null=True, blank=True, related_name='return_bills')
    return_reason = models.CharField(max_length=50, choices=[
        ('DAMAGED', 'Damaged'),
        ('EXPIRED', 'Expired'),
        ('OTHER', 'Other'),
    ], null=True, blank=True)

    session = models.ForeignKey(
        'DailySession', 
        on_delete=models.SET_NULL, 
        null=True, 
        blank=True, 
        related_name='sales_bills'
    )
    foc_value = models.DecimalField(max_digits=12, decimal_places=2, default=0)  # Track FOC value

    def get_paid_amount(self):
        total = self.payments.aggregate(total=Sum('amount'))['total'] or 0
        return total

    def get_outstanding(self):
        return self.net_total - self.get_paid_amount()

    def is_fully_paid(self):
        return self.get_outstanding() == 0

    def get_payment_status(self):
        if self.is_fully_paid():
            return 'PAID'
        elif self.get_paid_amount() > 0:
            return 'PARTIAL'
        else:
            return 'OUTSTANDING'

    bill_discount_type = models.CharField(max_length=20, choices=[
        ('PERCENTAGE', 'Percentage'),
        ('FIXED', 'Fixed Amount'),
    ], null=True, blank=True)
    bill_discount_value = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    bill_discount_amount = models.DecimalField(max_digits=12, decimal_places=2, default=0)  # Calculated    


class SalesItem(models.Model):
    """Line items for the bill (positive = sale, negative = return)"""
    bill = models.ForeignKey(SalesBill, on_delete=models.CASCADE, related_name='items')
    product = models.ForeignKey(Product, on_delete=models.CASCADE)
    quantity = models.DecimalField(max_digits=12, decimal_places=2)
    rate = models.DecimalField(max_digits=12, decimal_places=2)
    total = models.DecimalField(max_digits=12, decimal_places=2)
    is_foc = models.BooleanField(default=False)

    def __str__(self):
        return f"{self.product.name} x {self.quantity}"

    discount_type = models.CharField(max_length=20, choices=[
        ('PERCENTAGE', 'Percentage'),
        ('FIXED', 'Fixed Amount'),
    ], null=True, blank=True)
    discount_value = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    discounted_rate = models.DecimalField(max_digits=12, decimal_places=2, default=0)

    return_reason = models.CharField(max_length=20, choices=[
        ('DAMAGED', 'Damaged'),
        ('EXPIRED', 'Expired'),
        ('OTHER', 'Other'),
    ], null=True, blank=True)


class Payment(models.Model):
    """Split payments: Cash, Credit, Cheque"""
    PAYMENT_TYPES = [
        ('Cash', 'Cash'),
        ('Credit', 'Credit'),
        ('Cheque', 'Cheque'),
    ]
    bill = models.ForeignKey(SalesBill, on_delete=models.CASCADE, related_name='payments')
    type = models.CharField(max_length=10, choices=PAYMENT_TYPES)
    amount = models.DecimalField(max_digits=12, decimal_places=2)

    def __str__(self):
        return f"{self.type}: {self.amount}"
    is_reversed = models.BooleanField(default=False)  # ✅ New field
    reversed_at = models.DateTimeField(null=True, blank=True)
    reversed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='reversed_payments')
    reversed_cheque = models.ForeignKey('Cheque', on_delete=models.SET_NULL, null=True, blank=True, related_name='reversed_payment')
    

class UserProfile(models.Model):
    """Extends the built-in User model with role-based permissions"""
    
    ROLE_CHOICES = [
        ('Admin', 'Admin - Full Access'),
        ('Accountant', 'Accountant - Sales & Reports'),
        ('Loader', 'Loader - Vehicle Loading Only'),
        ('Viewer', 'Viewer - Read Only Reports'),
    ]
    
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='profile')
    role = models.CharField(max_length=20, choices=ROLE_CHOICES, default='Viewer')
    phone = models.CharField(max_length=20, blank=True, null=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.user.username} - {self.role}"

    def has_permission(self, permission):
        if self.role == 'Admin':
            return True
        elif self.role == 'Users':          # New combined role
            # Everything Accountant + Loader could do
            return permission in [
                # Sales & Credit
                'view_sales', 'create_sales', 'view_reports',
                'view_products', 'view_vehicles', 'load_vehicle',
                'view_employees', 'view_expenses', 'manage_expenses',
                'view_purchases', 'view_suppliers',
                'view_customers', 'view_cheques',
                'view_credit_list', 'credit_collection',
                'view_transfers', 'manage_transfers',
                'view_vehicle_stock',
                'view_collections',
            ]
        elif self.role == 'Viewer':
            return permission in ['view_reports', 'view_products']
        return False


class Employee(models.Model):
    POSITION_CHOICES = [
        ('Rep', 'Sales Representative'),
        ('Driver', 'Driver'),
        ('Helper', 'Helper'),
        ('Accountant', 'Accountant'),
        ('Other', 'Other'),
    ]
    
    name = models.CharField(max_length=200)
    position = models.CharField(max_length=20, choices=POSITION_CHOICES)
    phone = models.CharField(max_length=20, blank=True, null=True)
    email = models.EmailField(blank=True, null=True)
    address = models.TextField(blank=True, null=True)
    hire_date = models.DateField(null=True, blank=True)
    salary = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)
    notes = models.TextField(blank=True, null=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    rep_code = models.CharField(max_length=20, unique=True, blank=True, null=True, help_text="Unique code for Sales Representatives (e.g., 008SIN)")
    rep_invoice_counter = models.IntegerField(default=0, help_text="Last invoice number used for this Rep")
    
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    def __str__(self):
        return f"{self.name} ({self.get_position_display()})"
    
    def save(self, *args, **kwargs):
        # If position is Rep and rep_code is empty, raise validation error (handled in form)
        super().save(*args, **kwargs)
    
    def get_next_invoice_number(self):
        """Generate the next invoice number for this rep."""
        if not self.rep_code:
            return None
        self.rep_invoice_counter += 1
        self.save()
        return f"{self.rep_code}-{self.rep_invoice_counter:03d}"
    

class Bank(models.Model):
    name = models.CharField(max_length=200, unique=True)
    branch = models.CharField(max_length=200, blank=True, null=True)
    account_no = models.CharField(max_length=50, blank=True, null=True)
    account_type = models.CharField(max_length=20, choices=[
        ('CURRENT', 'Current'),
        ('SAVINGS', 'Savings'),
    ], default='CURRENT')
    contact = models.CharField(max_length=100, blank=True, null=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.name

    class Meta:
        ordering = ['name']


class ChequePayment(models.Model):
    bill = models.OneToOneField(SalesBill, on_delete=models.CASCADE, related_name='cheque_detail')
    cheque_no = models.CharField(max_length=50)
    cheque_date = models.DateField()
    bank = models.ForeignKey(Bank, on_delete=models.CASCADE)
    amount = models.DecimalField(max_digits=12, decimal_places=2)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.cheque_no} - {self.bank.name}"

    
class OnlinePayment(models.Model):
    PAYMENT_METHODS = [
        ('Bank Transfer', 'Bank Transfer'),
        ('Mobile Wallet', 'Mobile Wallet'),
        ('Card', 'Card'),
        ('Other', 'Other'),
    ]
    
    bill = models.OneToOneField(SalesBill, on_delete=models.CASCADE, related_name='online_detail')
    payment_method = models.CharField(max_length=50, choices=PAYMENT_METHODS)
    reference_no = models.CharField(max_length=100, blank=True, null=True)
    amount = models.DecimalField(max_digits=12, decimal_places=2)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.payment_method} - {self.reference_no or 'N/A'}"

class MultiPayment(models.Model):
    bill = models.OneToOneField(SalesBill, on_delete=models.CASCADE, related_name='multi_detail')
    payment_type_1 = models.CharField(max_length=20, choices=[
        ('Cash', 'Cash'),
        ('Credit', 'Credit'),
        ('Cheque', 'Cheque'),
        ('Online', 'Online'),
    ])
    amount_1 = models.DecimalField(max_digits=12, decimal_places=2)
    payment_type_2 = models.CharField(max_length=20, choices=[
        ('Cash', 'Cash'),
        ('Credit', 'Credit'),
        ('Cheque', 'Cheque'),
        ('Online', 'Online'),
    ])
    amount_2 = models.DecimalField(max_digits=12, decimal_places=2)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.payment_type_1} + {self.payment_type_2}"
    
    # ==================== SUPPLIER ====================


class Supplier(models.Model):
    name = models.CharField(max_length=200)
    contact_person = models.CharField(max_length=100, blank=True, null=True)
    phone = models.CharField(max_length=20, blank=True, null=True)
    email = models.EmailField(blank=True, null=True)
    address = models.TextField(blank=True, null=True)
    tax_number = models.CharField(max_length=50, blank=True, null=True, help_text="Supplier's VAT/Tax number")
    notes = models.TextField(blank=True, null=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.name

    class Meta:
        ordering = ['name']


class Purchase(models.Model):
    PAYMENT_STATUS_CHOICES = [
    ('CREDIT', 'Credit Purchases'),
    ('COMPLETE', 'Complete Payment'),
    ]
    
    STATUS_CHOICES = [
        ('DRAFT', 'Draft'),
        ('PENDING', 'Pending'),
        ('RECEIVED', 'Received'),
        ('CANCELLED', 'Cancelled'),
    ]
    
    TAX_RATE_CHOICES = [
        (0, 'No Tax'),
        (5, 'VAT 5%'),
        (8, 'VAT 8%'),
        (12, 'VAT 12%'),
        (15, 'VAT 15%'),
    ]
    
    # Header Fields
    supplier = models.ForeignKey(Supplier, on_delete=models.PROTECT, related_name='purchases')
    rep = models.ForeignKey(Employee, on_delete=models.SET_NULL, null=True, blank=True, related_name='purchases')
    po_number = models.CharField(max_length=50, unique=True, blank=True, help_text="Your PO Number")
    invoice_no = models.CharField(max_length=100, help_text="Supplier's Invoice Number")
    purchase_date = models.DateField(auto_now_add=True)
    due_date = models.DateField(null=True, blank=True)
    tax_rate = models.IntegerField(choices=TAX_RATE_CHOICES, default=0)
    tax_invoice_no = models.CharField(max_length=100, blank=True, null=True, help_text="Tax Invoice Number")
    
    # Totals
    subtotal = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    tax_amount = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    total = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    paid_amount = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    
    # Status
    payment_status = models.CharField(max_length=20, choices=PAYMENT_STATUS_CHOICES, default='PENDING')
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='DRAFT')
    
    notes = models.TextField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='created_purchases')
    updated_at = models.DateTimeField(auto_now=True)
    received_at = models.DateTimeField(null=True, blank=True)
    received_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='received_purchases')

    @property
    def balance(self):
        """Calculate remaining balance (total - paid)"""
        return self.total - self.paid_amount
    
    def __str__(self):
        return f"{self.invoice_no} - {self.supplier.name}"
    
    def get_total_items(self):
        return self.items.aggregate(total=Sum('quantity'))['total'] or 0
    
    def get_profit_margin(self, product=None):
        """Calculate average profit margin"""
        if self.items.exists():
            total_cost = self.items.aggregate(total=Sum('total'))['total'] or 0
            total_retail = self.items.aggregate(total=Sum('retail_total'))['total'] or 0
            if total_cost > 0:
                return ((total_retail - total_cost) / total_cost) * 100
        return 0
    
    class Meta:
        ordering = ['-purchase_date', '-created_at']


class PurchaseItem(models.Model):
    purchase = models.ForeignKey(Purchase, on_delete=models.CASCADE, related_name='items')
    product = models.ForeignKey(Product, on_delete=models.PROTECT, related_name='purchase_items')
    quantity = models.DecimalField(max_digits=15, decimal_places=2)
    cost_price = models.DecimalField(max_digits=12, decimal_places=2)
    retail_price = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    wholesale_price = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    total = models.DecimalField(max_digits=15, decimal_places=2)  # quantity × cost_price
    retail_total = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    wholesale_total = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    is_foc = models.BooleanField(default=False)  # ✅ NEW FIELD
    
    def __str__(self):
        return f"{self.product.name} x {self.quantity}"
    
    def get_profit_margin(self):
        if self.cost_price > 0:
            return ((self.retail_price - self.cost_price) / self.cost_price) * 100
        return 0


class StockMovement(models.Model):
    MOVEMENT_TYPES = [
        ('PURCHASE', 'Purchase Received'),
        ('SALE', 'Sale'),
        ('RETURN', 'Return'),
        ('ADJUSTMENT', 'Stock Adjustment'),
        ('LOAD', 'Vehicle Load'),
    ]
    
    product = models.ForeignKey(Product, on_delete=models.CASCADE, related_name='stock_movements')
    movement_type = models.CharField(max_length=20, choices=MOVEMENT_TYPES)
    quantity = models.DecimalField(max_digits=15, decimal_places=2)
    previous_stock = models.DecimalField(max_digits=15, decimal_places=2)
    new_stock = models.DecimalField(max_digits=15, decimal_places=2)
    reference = models.CharField(max_length=200, blank=True, null=True, help_text="PO, Invoice, or Sale reference")
    notes = models.TextField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='stock_movements')
    
    def __str__(self):
        return f"{self.product.name} - {self.movement_type}: {self.quantity}"
    
    class Meta:
        ordering = ['-created_at']

        # ==================== CHEQUE MANAGEMENT ====================


class Cheque(models.Model):
    STATUS_CHOICES = [
        ('PENDING', 'Pending'),
        ('DEPOSITED', 'Deposited'),
        ('CLEARED', 'Cleared'),
        ('BOUNCED', 'Bounced'),
    ]
    
    # Cheque Details
    cheque_no = models.CharField(max_length=50)
    bank = models.ForeignKey(Bank, on_delete=models.PROTECT, related_name='cheques')
    cheque_date = models.DateField()
    amount = models.DecimalField(max_digits=12, decimal_places=2)
    
    # Customer & Sales Reference
    customer_name = models.CharField(max_length=255, help_text="Customer who gave the cheque")
    sales_bill = models.ForeignKey(SalesBill, on_delete=models.CASCADE, related_name='cheques')
    
    # Status Tracking
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='PENDING')
    deposit_date = models.DateField(null=True, blank=True)
    cleared_date = models.DateField(null=True, blank=True)
    bounce_reason = models.CharField(max_length=255, blank=True, null=True, help_text="Reason if bounced")
    
    notes = models.TextField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    def __str__(self):
        return f"{self.cheque_no} - {self.customer_name} ({self.amount})"
    
    def get_status_display_color(self):
        colors = {
            'PENDING': 'warning',
            'DEPOSITED': 'info',
            'CLEARED': 'success',
            'BOUNCED': 'danger',
        }
        return colors.get(self.status, 'secondary')
# Bounce info
    bounce_reason = models.CharField(max_length=255, blank=True, null=True)
    bounced_at = models.DateTimeField(null=True, blank=True)
    bounced_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='bounced_cheques')
    bank_charge_amount = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    bank_charge_expense = models.ForeignKey('Expense', on_delete=models.SET_NULL, null=True, blank=True, related_name='cheque_bounce_charges')
    
    class Meta:
        ordering = ['-created_at']

    created_at = models.DateTimeField(auto_now_add=True)    

        # ==================== CUSTOMER ====================


class Customer(models.Model):
    CUSTOMER_TYPES = [
        ('WHOLESALE', 'Wholesale'),
        ('RETAIL', 'Retail'),
        ('CORPORATE', 'Corporate'),
        ('OTHER', 'Other'),
    ]
    
    # Basic Info
    name = models.CharField(max_length=255, help_text="Customer/Shop name")
    code = models.CharField(max_length=50, unique=True, blank=True, help_text="Auto-generated if blank")
    customer_type = models.CharField(max_length=20, choices=CUSTOMER_TYPES, default='RETAIL')
    
    # Contact Info
    phone = models.CharField(max_length=20, blank=True, null=True)
    email = models.EmailField(blank=True, null=True)
    address = models.TextField(blank=True, null=True)
    city = models.CharField(max_length=100, blank=True, null=True)
    
    # Business Info
    tax_number = models.CharField(max_length=50, blank=True, null=True, help_text="VAT/TIN number")
    credit_limit = models.DecimalField(max_digits=12, decimal_places=2, default=0, help_text="Maximum credit allowed (0 = unlimited)")
    
    # Status
    is_active = models.BooleanField(default=True)
    notes = models.TextField(blank=True, null=True)
    
    # Timestamps
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    def __str__(self):
        return f"{self.code} - {self.name}"
    
    def save(self, *args, **kwargs):
        # Auto-generate code if not provided
        if not self.code:
            last_customer = Customer.objects.all().order_by('-id').first()
            if last_customer and last_customer.code:
                try:
                    last_num = int(last_customer.code.replace('CUST-', ''))
                    new_num = last_num + 1
                except:
                    new_num = 1
            else:
                new_num = 1
            self.code = f"CUST-{new_num:04d}"
        super().save(*args, **kwargs)
    
    def get_total_sales(self):
        """Get total sales value for this customer"""
        total = SalesBill.objects.filter(shop_code=self.code).aggregate(total=Sum('net_total'))['total']
        return total or 0
    
    def get_outstanding_balance(self):
        """Get current outstanding balance (credit - paid)"""
        total_credit = Payment.objects.filter(
            bill__shop_code=self.code,
            type='Credit'
        ).aggregate(total=Sum('amount'))['total'] or 0
        
        total_paid = Payment.objects.filter(
            bill__shop_code=self.code
        ).exclude(type='Credit').aggregate(total=Sum('amount'))['total'] or 0
        
        return total_credit - total_paid
    
    def get_available_credit(self):
        """Get available credit (credit limit - outstanding)"""
        if self.credit_limit == 0:
            return float('inf')  # Unlimited
        return self.credit_limit - self.get_outstanding_balance()
    
    class Meta:
        ordering = ['name']


class Expense(models.Model):
    CATEGORY_CHOICES = [
        # Vehicle-Related
        ('Fuel', 'Fuel'),
        ('Maintenance', 'Maintenance / Repairs'),
        ('Tires', 'Tires'),
        ('Spare Parts', 'Spare Parts'),
        ('Vehicle Insurance', 'Vehicle Insurance'),
        ('Vehicle Registration', 'Vehicle Registration'),
        ('Toll Charges', 'Toll Charges'),
        ('Parking Fees', 'Parking Fees'),
        ('Washing', 'Washing / Cleaning'),
        ('Driver Allowance', 'Driver Allowance'),
        
        # Employee-Related
        ('Salaries', 'Salaries'),
        ('Bonuses', 'Bonuses'),
        ('Overtime', 'Overtime'),
        ('Medical', 'Medical / Health'),
        ('Uniforms', 'Uniforms'),
        ('Training', 'Training'),
        ('Meals', 'Staff Meals'),
        
        # Operational
        ('Warehouse Rent', 'Warehouse Rent'),
        ('Electricity', 'Electricity'),
        ('Water', 'Water'),
        ('Internet', 'Internet'),
        ('Phone Bills', 'Phone Bills'),
        ('Office Supplies', 'Office Supplies'),
        ('Cleaning', 'Cleaning'),
        ('Security', 'Security'),
        ('Business Insurance', 'Business Insurance'),
        ('Rates & Taxes', 'Rates & Taxes'),
        
        # Sales & Marketing
        ('Advertising', 'Advertising'),
        ('Promotions', 'Promotions'),
        ('Sampling', 'Sampling'),
        ('Sales Materials', 'Sales Materials'),
        ('Customer Events', 'Customer Events'),
        
        # Administrative
        ('Bank Charges', 'Bank Charges'),
        ('Legal Fees', 'Legal Fees'),
        ('Accounting', 'Accounting'),
        ('Consulting', 'Consulting'),
        ('Licenses', 'Licenses'),
        ('Subscriptions', 'Subscriptions'),
        ('Memberships', 'Memberships'),
        
        # Miscellaneous
        ('Donations', 'Donations'),
        ('Gifts', 'Gifts'),
        ('Travel', 'Travel'),
        ('Postage', 'Postage'),
        ('Other', 'Other'),
    ]
    
    PAYMENT_METHODS = [
        ('Cash', 'Cash'),
        ('Credit', 'Credit'),
        ('Cheque', 'Cheque'),
        ('Online', 'Online'),
    ]
    
    STATUS_CHOICES = [
        ('PENDING', 'Pending'),
        ('APPROVED', 'Approved'),
        ('PAID', 'Paid'),
        ('REJECTED', 'Rejected'),
    ]
    
    # Existing fields
    vehicle = models.ForeignKey(Vehicle, on_delete=models.CASCADE, related_name='expenses', null=True, blank=True)
    date = models.DateField(auto_now_add=False, default=date.today)
    category = models.CharField(max_length=50, choices=CATEGORY_CHOICES)
    amount = models.DecimalField(max_digits=10, decimal_places=2, default=0)  # ✅ Added default
    note = models.CharField(max_length=255, blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    
    # New fields
    employee = models.ForeignKey('Employee', on_delete=models.SET_NULL, null=True, blank=True, related_name='expenses')
    payment_method = models.CharField(max_length=20, choices=PAYMENT_METHODS, default='Cash')
    invoice_no = models.CharField(max_length=100, blank=True, null=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='PENDING')
    approved_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='approved_expenses')
    approved_at = models.DateTimeField(null=True, blank=True)
    paid_at = models.DateTimeField(null=True, blank=True)
    rejected_reason = models.CharField(max_length=255, blank=True, null=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    def __str__(self):
        return f"{self.category} - {self.amount} ({self.date})"
    
    def get_status_display_color(self):
        colors = {
            'PENDING': 'warning',
            'APPROVED': 'info',
            'PAID': 'success',
            'REJECTED': 'danger',
        }
        return colors.get(self.status, 'secondary')

    session = models.ForeignKey(
        'DailySession', 
        on_delete=models.SET_NULL, 
        null=True, 
        blank=True, 
        related_name='expenses'
    )
    
class StockTransfer(models.Model):
    REASON_CHOICES = [
        ('BREAKDOWN', 'Vehicle Breakdown'),
        ('ROUTE_CHANGE', 'Route Change'),
        ('STOCK_BALANCE', 'Stock Balancing'),
        ('DRIVER_CHANGE', 'Driver Change'),
        ('OTHER', 'Other'),
    ]
    
    source_vehicle = models.ForeignKey(
        Vehicle, 
        on_delete=models.CASCADE, 
        related_name='transfers_out'
    )
    destination_vehicle = models.ForeignKey(
        Vehicle, 
        on_delete=models.CASCADE, 
        related_name='transfers_in'
    )
    product = models.ForeignKey(Product, on_delete=models.CASCADE)
    quantity = models.DecimalField(max_digits=15, decimal_places=2)
    transfer_date = models.DateTimeField(auto_now_add=True)
    transferred_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='stock_transfers')
    reason = models.CharField(max_length=20, choices=REASON_CHOICES, default='STOCK_BALANCE')
    notes = models.TextField(blank=True, null=True)
    
    def __str__(self):
        return f"{self.source_vehicle.vehicle_number} → {self.destination_vehicle.vehicle_number}: {self.product.name} x{self.quantity}"
    
    class Meta:
        ordering = ['-transfer_date']


class CreditCollection(models.Model):
    STATUS_CHOICES = [
        ('PENDING', 'Pending'),
        ('TAKEN', 'Taken for Collection'),
        ('COLLECTED', 'Collected'),
        ('NOT_COLLECTED', 'Not Collected'),
    ]
    
    REASON_CHOICES = [
        ('CUSTOMER_NOT_AVAILABLE', 'Customer Not Available'),
        ('REFUSED_TO_PAY', 'Customer Refused to Pay'),
        ('INSUFFICIENT_FUNDS', 'Insufficient Funds'),
        ('CUSTOMER_REQUESTED_DELAY', 'Customer Requested Delay'),
        ('OTHER', 'Other'),
    ]
    
    PAYMENT_METHODS = [
        ('Cash', 'Cash'),
        ('Cheque', 'Cheque'),
        ('Online', 'Online'),
    ]
    
    # Links
    sales_bill = models.ForeignKey(SalesBill, on_delete=models.CASCADE, related_name='collections')
    rep = models.ForeignKey(Employee, on_delete=models.SET_NULL, null=True, blank=True, related_name='collections')
    collected_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='collections_done')
    
    # Dates
    date_taken = models.DateField(null=True, blank=True)
    date_collected = models.DateField(null=True, blank=True)
    
    # Status
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='PENDING')
    
    # Collection details
    collection_amount = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    payment_method = models.CharField(max_length=20, choices=PAYMENT_METHODS, null=True, blank=True)
    not_collected_reason = models.CharField(max_length=50, choices=REASON_CHOICES, null=True, blank=True)
    notes = models.TextField(blank=True, null=True)
    
    # Timestamps
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    def __str__(self):
        return f"{self.sales_bill.invoice_no} - {self.status} ({self.sales_bill.shop_name})"
    
    def get_status_badge(self):
        badges = {
            'PENDING': 'warning',
            'TAKEN': 'info',
            'COLLECTED': 'success',
            'NOT_COLLECTED': 'danger',
        }
        return badges.get(self.status, 'secondary')
    
    class Meta:
        ordering = ['-created_at']


class DailySession(models.Model):
    STATUS_CHOICES = [
        ('OPEN', 'Open'),
        ('PENDING', 'Pending Review'),
        ('COMPLETED', 'Completed'),
        ('CANCELLED', 'Cancelled'),
    ]
    
    # Core Fields
    vehicle = models.ForeignKey(Vehicle, on_delete=models.CASCADE, related_name='sessions')
    rep = models.ForeignKey(Employee, on_delete=models.CASCADE, related_name='sessions')
    date = models.DateField(auto_now_add=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='OPEN')
    
    # Timestamps
    started_at = models.DateTimeField(auto_now_add=True)
    completed_at = models.DateTimeField(null=True, blank=True)
    
    # User Tracking
    started_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='started_sessions')
    completed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='completed_sessions')
    
    # Summary Fields (cached for quick display)
    total_sales = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    total_foc = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    total_expenses = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    total_bills = models.IntegerField(default=0)
    total_expense_count = models.IntegerField(default=0)
    
    # Session ID for display
    session_id = models.CharField(max_length=20, unique=True, blank=True)
    
    def __str__(self):
        return f"{self.session_id} - {self.rep.name} ({self.vehicle.vehicle_number})"
    
    def save(self, *args, **kwargs):
        if not self.session_id:
            # Generate session ID: SESS-YYYYMMDD-XXX
            date_str = datetime.now().strftime('%Y%m%d')
            last_session = DailySession.objects.filter(session_id__startswith=f'SESS-{date_str}').order_by('-session_id').first()
            if last_session:
                try:
                    last_num = int(last_session.session_id.split('-')[-1])
                    new_num = last_num + 1
                except:
                    new_num = 1
            else:
                new_num = 1
            self.session_id = f"SESS-{date_str}-{new_num:03d}"
        super().save(*args, **kwargs)
    
    def get_pending_bills(self):
        """Get all pending bills in this session"""
        return self.sales_bills.filter(status='DRAFT')
    
    def get_pending_expenses(self):
        """Get all pending expenses in this session"""
        return self.expenses.filter(status='PENDING')
    
    def get_completed_bills(self):
        """Get all completed bills in this session"""
        return self.sales_bills.filter(status='COMPLETED')
    
    def update_summary(self):
        """Update cached summary fields"""
        pending_bills = self.get_pending_bills()
        completed_bills = self.get_completed_bills()
        all_bills = pending_bills | completed_bills
        
        self.total_bills = all_bills.count()
        self.total_sales = all_bills.aggregate(total=Sum('net_total'))['total'] or 0
        self.total_foc = all_bills.aggregate(total=Sum('foc_value'))['total'] or 0
        self.total_expenses = self.expenses.filter(status='PAID').aggregate(total=Sum('amount'))['total'] or 0
        self.total_expense_count = self.expenses.filter(status='PAID').count()
        self.save()
    
    def is_active(self):
        return self.status == 'OPEN'
    
    def can_edit(self):
        return self.status in ['OPEN', 'PENDING']
    
    class Meta:
        ordering = ['-started_at']


class VehicleLoad(models.Model):
    vehicle = models.ForeignKey(Vehicle, on_delete=models.CASCADE, related_name='loads')
    product = models.ForeignKey(Product, on_delete=models.CASCADE)
    quantity = models.DecimalField(max_digits=15, decimal_places=2)
    loaded_at = models.DateTimeField(auto_now_add=True)
    notes = models.TextField(blank=True, null=True)

    def __str__(self):
        return f"{self.vehicle.vehicle_number} - {self.product.name} x{self.quantity}"

    class Meta:
        ordering = ['-loaded_at']


class StockMovementLog(models.Model):
    MOVEMENT_TYPES = [
        ('LOAD', 'Load to Vehicle'),
        ('UNLOAD', 'Unload from Vehicle'),
    ]
    
    vehicle = models.ForeignKey(Vehicle, on_delete=models.CASCADE, related_name='stock_movement_logs')
    product = models.ForeignKey(Product, on_delete=models.CASCADE)
    quantity = models.DecimalField(max_digits=15, decimal_places=2)
    movement_type = models.CharField(max_length=10, choices=MOVEMENT_TYPES)
    performed_by = models.ForeignKey(
        User, 
        on_delete=models.SET_NULL, 
        null=True, 
        blank=True, 
        related_name='stock_movement_logs'
    )
    notes = models.TextField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.vehicle.vehicle_number} - {self.movement_type} {self.product.name} x{self.quantity}"

    class Meta:
        ordering = ['-created_at']




