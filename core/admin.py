from django.contrib import admin
from .models import Product, WarehouseStock, VehicleStock, SalesBill, SalesItem, Payment, UserProfile, Employee, Bank, ChequePayment, OnlinePayment, MultiPayment, Supplier, Purchase, PurchaseItem, StockMovementLog
from .models import Employee
from .models import Category 
from .models import Bank, Cheque
from .models import Customer
from .models import Expense
from .models import Vehicle
from .models import StockTransfer

admin.site.register(Product)
admin.site.register(WarehouseStock)
admin.site.register(Vehicle)
admin.site.register(VehicleStock)
admin.site.register(SalesBill)
admin.site.register(SalesItem)
admin.site.register(Payment)
admin.site.register(UserProfile)
admin.site.register(Employee)
admin.site.register(Bank)
admin.site.register(ChequePayment)
admin.site.register(OnlinePayment)
admin.site.register(MultiPayment)
admin.site.register(Category)
admin.site.register(Supplier)
admin.site.register(Purchase)
admin.site.register(PurchaseItem)
admin.site.register(StockMovementLog)
admin.site.register(Cheque)
admin.site.register(Customer)
admin.site.register(Expense)
admin.site.register(StockTransfer)


