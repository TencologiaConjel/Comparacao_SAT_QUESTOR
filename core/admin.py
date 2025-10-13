from django.contrib import admin

from django.contrib import admin
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from django.contrib.auth.models import User
from django import forms
from django.core.exceptions import ValidationError

# Form de criação que inclui e-mail
class UserCreationEmailForm(forms.ModelForm):
    password1 = forms.CharField(label='Senha', widget=forms.PasswordInput)
    password2 = forms.CharField(label='Confirmação de senha', widget=forms.PasswordInput)

    class Meta:
        model = User
        fields = ('username', 'email')  # agora o e-mail aparece no "Adicionar"

    def clean_email(self):
        email = (self.cleaned_data.get('email') or '').strip().lower()
        if not email:
            raise ValidationError('Informe um e-mail.')
        if User._default_manager.filter(email__iexact=email).exists():
            raise ValidationError('Já existe um usuário com este e-mail.')
        return email

    def clean(self):
        cleaned = super().clean()
        p1, p2 = cleaned.get('password1'), cleaned.get('password2')
        if p1 and p2 and p1 != p2:
            raise ValidationError('As senhas não conferem.')
        return cleaned

    def save(self, commit=True):
        user = super().save(commit=False)
        user.set_password(self.cleaned_data['password1'])
        if commit:
            user.save()
        return user

class CustomUserAdmin(BaseUserAdmin):
    add_form = UserCreationEmailForm
    add_fieldsets = (
        (None, {
            'classes': ('wide',),
            'fields': ('username', 'email', 'password1', 'password2'),
        }),
    )
    list_display = ('username', 'email', 'is_active', 'is_staff', 'is_superuser')
    search_fields = ('username', 'email')

admin.site.unregister(User)
admin.site.register(User, CustomUserAdmin)
# core/admin.py
from django.contrib import admin
from django.utils.safestring import mark_safe
import json

from .models import Empresa, SatRegistro, Documentos, LoginLog


@admin.register(Empresa)
class EmpresaAdmin(admin.ModelAdmin):
    list_display = ("nome", "cnpj")
    search_fields = ("nome", "cnpj")
    ordering = ("nome",)


@admin.register(SatRegistro)
class SatRegistroAdmin(admin.ModelAdmin):
    list_display = (
        "empresa", "sheet", "row",
        "descricao", "ncm", "cfop", "cest", "cst_csosn",
        "created_at",
    )
    search_fields = (
        "empresa__nome", "empresa__cnpj",
        "sheet", "descricao", "ncm", "cfop", "cest", "cst_csosn",
    )
    list_filter = ("sheet", "empresa")
    autocomplete_fields = ("empresa",)
    list_select_related = ("empresa",)
    date_hierarchy = "created_at"
    ordering = ("-created_at",)
    readonly_fields = ("created_at", "updated_at", "data_pretty")

    fieldsets = (
        (None, {"fields": ("empresa", "sheet", "row")}),
        ("Chaves de pesquisa", {"fields": ("descricao", "ncm", "cfop", "cest", "cst_csosn")}),
        ("Dados completos", {"fields": ("data_pretty",)}),
        ("Metadados", {"fields": ("created_at", "updated_at")}),
    )

    def data_pretty(self, obj):
        try:
            content = json.dumps(obj.data or {}, indent=2, ensure_ascii=False)
        except Exception:
            content = str(obj.data)
        return mark_safe(f"<pre style='max-width:100%;white-space:pre-wrap;'>{content}</pre>")
    data_pretty.short_description = "Dados (JSON)"


@admin.register(Documentos)
class DocumentosAdmin(admin.ModelAdmin):
    list_display = (
        "empresa", "chaveacesso", "modelodocumento", "tipodocumento",
        "dataemissao", "nomedoemitente", "cnpjdoemitente",
        "nomedodestinatario", "valornfe",
    )
    search_fields = (
        "empresa__nome", "empresa__cnpj",
        "chaveacesso", "cnpjdoemitente", "cpfdodestinatario", "cnpjdodestinatario",
        "nomedoemitente", "nomedodestinatario", "numerodocumento", "serie",
    )
    list_filter = ("situacao", "ufemitente", "ufdestinatario")
    autocomplete_fields = ("empresa",)
    list_select_related = ("empresa",)
    ordering = ("-dataemissao",) 

@admin.register(LoginLog)
class LoginLogAdmin(admin.ModelAdmin):
    list_display = ("email", "user", "success", "created_at")
    search_fields = ("email", "user__username", "user__email")
    list_filter = ("success",)
    date_hierarchy = "created_at"
    autocomplete_fields = ("user",)
    ordering = ("-created_at",)
