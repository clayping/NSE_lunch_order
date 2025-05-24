from django.contrib import admin
from .models import LunchConfig, Order

@admin.register(LunchConfig)
class LunchConfigAdmin(admin.ModelAdmin):
    list_display = ('price', 'subsidy', 'monthly_limit')
    ordering     = ('-id',)

@admin.register(Order)
class OrderAdmin(admin.ModelAdmin):
    list_display  = ('user', 'order_date', 'vendor', 'rice_size', 'quantity', 'price', 'subsidy', 'canceled')
    list_filter   = ('order_date', 'vendor', 'rice_size', 'canceled')
    search_fields = ('user__username',)
