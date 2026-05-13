from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static
from core import stats_views

urlpatterns = [
    path('admin/', admin.site.urls),

    path('stats/', stats_views.stats_landing, name='stats'),
    path('stats/workers/', stats_views.worker_stats, name='worker_stats'),
    path('stats/run-sheets/', stats_views.run_sheet_stats, name='run_sheet_stats'),

    path('', include('core.urls')),
    path('accounts/', include('django.contrib.auth.urls')),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
