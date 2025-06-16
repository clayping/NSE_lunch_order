from django.contrib import admin
from django.urls import reverse
from .models import LunchConfig, Order
from django.contrib.admin import AdminSite

@admin.register(LunchConfig)
class LunchConfigAdmin(admin.ModelAdmin):
    list_display = ('price', 'subsidy', 'monthly_limit')
    ordering     = ('-id',)

@admin.register(Order)
class OrderAdmin(admin.ModelAdmin):
    list_display  = ('user', 'order_date', 'vendor', 'rice_size', 'quantity', 'price', 'subsidy', 'canceled')
    list_filter   = ('order_date', 'vendor', 'rice_size', 'canceled')
    search_fields = ('user__username',)


class MyAdminSite(AdminSite):
    site_header = "NSランチ管理"
    def index(self, request, extra_context=None):
        if extra_context is None:
            extra_context = {}

            report_url = reverse('download_monthly_report')
            extra_context['report_link'] = format_html(
                '<a class="button" href="{}">月末レポートダウンロード</a>', report_url
            )
            return super().index(request, extra_context=extra_context)
