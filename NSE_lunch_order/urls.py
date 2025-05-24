"""
URL configuration for NSE_lunch_order project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/5.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""

from django.contrib import admin
from django.views.generic import RedirectView
from django.urls import path, include
from lunch.views import fax_order_pdf, today_order, monthly_calendar, toggle_order

urlpatterns = [
    path('', RedirectView.as_view(url='accounts/login/')),
    path('admin/', admin.site.urls),
    path('accounts/', include('allauth.urls')),

    path('fax-order/', fax_order_pdf,       name='fax_order_pdf'),
    path('order/',     today_order,        name='today_order'),

    # 今月表示と年月指定のいずれも同じビューを同じ名前で扱う
    path('calendar/',                    monthly_calendar, name='monthly_calendar'),
    path('calendar/<int:year>/<int:month>/', monthly_calendar, name='monthly_calendar'),

    path('api/toggle-order/', toggle_order, name='toggle_order'),
]
