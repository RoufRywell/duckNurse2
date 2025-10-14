"""DuckSpinner URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/4.0/topics/http/urls/
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
from django.urls import path
from django.conf import settings
from django.conf.urls.static import static
# 👇 Bu import'u ekleyin
from django.contrib.staticfiles.urls import staticfiles_urlpatterns 
from converter import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', views.home, name='home'),  # Ana sayfa için root URL
]

# 👇 Geliştirme ortamında (DEBUG=True) statik dosyaları sunmak için bu blok kullanılır.
if settings.DEBUG:
    # 1. Bu fonksiyon, INSTALLED_APPS içindeki tüm uygulamaların (converter dahil) 
    #    'static/' klasörlerini tarar ve URL'lerini oluşturur. (Sizin resimlerinizi sunar!)
    urlpatterns += staticfiles_urlpatterns() 

    # 2. Media dosyalarını sunar (Kullanıcı tarafından yüklenen dosyalar için gereklidir).
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)