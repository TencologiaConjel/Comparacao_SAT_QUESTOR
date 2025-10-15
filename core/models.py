from django.db import models
from django.conf import settings

class LoginLog(models.Model):
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True,
        on_delete=models.SET_NULL, related_name="login_logs",
    )
    email = models.EmailField(max_length=254, db_index=True)
    success = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-created_at"]
        indexes = [
            models.Index(fields=["email"]),
            models.Index(fields=["success"]),
            models.Index(fields=["created_at"]),
        ]

    def __str__(self):
        ok = "OK" if self.success else "FALHA"
        return f"{self.email} - {ok} - {self.created_at:%Y-%m-%d %H:%M}"


class Empresa(models.Model):
    nome = models.CharField('Razão social', max_length=255)
    cnpj = models.CharField('CNPJ', max_length=18, unique=True)

    class Meta:
        ordering = ['nome']
        verbose_name = 'Empresa'
        verbose_name_plural = 'Empresas'

    def __str__(self):
        return f"{self.nome} ({self.cnpj})"


class EmpresaOwnedModel(models.Model):
    empresa = models.ForeignKey(Empresa, on_delete=models.CASCADE)
    class Meta:
        abstract = True


class Documentos(EmpresaOwnedModel):
    periododereferencia = models.CharField(max_length=255, blank=True, null=True)
    modelodocumento = models.CharField(max_length=255, blank=True, null=True)
    tipodocumento = models.CharField(max_length=255, blank=True, null=True)
    tipodeoperacaoentradaousaida = models.CharField(max_length=255, blank=True, null=True)
    situacao = models.CharField(max_length=255, blank=True, null=True)
    chaveacesso = models.CharField(max_length=255, blank=True, null=True)
    dataemissao = models.CharField(max_length=255, blank=True, null=True)
    cnpjoucpfdoemitente = models.CharField(max_length=255, blank=True, null=True)
    cnpjdoemitente = models.CharField(max_length=255, blank=True, null=True)
    cpfdoemitente = models.CharField(max_length=255, blank=True, null=True)
    nomedoemitente = models.CharField(max_length=255, blank=True, null=True)
    ufemitente = models.CharField(max_length=255, blank=True, null=True)
    cidadeemitente = models.CharField(max_length=255, blank=True, null=True)
    inscricaoestadualemitente = models.CharField(max_length=255, blank=True, null=True)
    cnpjoucpfdomunicipiodoemitente = models.CharField(max_length=255, blank=True, null=True)
    cnpjoucpfdoindicadordepermanenciadoproduto = models.CharField(max_length=255, blank=True, null=True)
    cnpjoucpfdodestinatario = models.CharField(max_length=255, blank=True, null=True)
    cnpjdodestinatario = models.CharField(max_length=255, blank=True, null=True)
    cpfdodestinatario = models.CharField(max_length=255, blank=True, null=True)
    nomedodestinatario = models.CharField(max_length=255, blank=True, null=True)
    ufdestinatario = models.CharField(max_length=255, blank=True, null=True)
    cidadedodestinatario = models.CharField(max_length=255, blank=True, null=True)
    inscricaoestadualdestinatario = models.CharField(max_length=255, blank=True, null=True)
    cnpjoucpfdomunicipiododestinatario = models.CharField(max_length=255, blank=True, null=True)
    dataentrada = models.CharField(max_length=255, blank=True, null=True)
    numerodocumento = models.CharField(max_length=255, blank=True, null=True)
    serie = models.CharField(max_length=255, blank=True, null=True)
    valorprodutos = models.CharField(max_length=255, blank=True, null=True)
    valortotal = models.CharField(max_length=255, blank=True, null=True)
    valorbcicms = models.CharField(max_length=255, blank=True, null=True)
    valoricms = models.CharField(max_length=255, blank=True, null=True)
    valorbcsimples = models.CharField(max_length=255, blank=True, null=True)
    valorsimples = models.CharField(max_length=255, blank=True, null=True)
    valorbcicmsst = models.CharField(max_length=255, blank=True, null=True)
    valoricmsst = models.CharField(max_length=255, blank=True, null=True)
    valorfcp = models.CharField(max_length=255, blank=True, null=True)
    valorbcfcpst = models.CharField(max_length=255, blank=True, null=True)
    valorfcpst = models.CharField(max_length=255, blank=True, null=True)
    valorpis = models.CharField(max_length=255, blank=True, null=True)
    valorcofins = models.CharField(max_length=255, blank=True, null=True)
    valorbcipi = models.CharField(max_length=255, blank=True, null=True)
    valoripi = models.CharField(max_length=255, blank=True, null=True)
    valoroutrasdespesas = models.CharField(max_length=255, blank=True, null=True)
    valorseguro = models.CharField(max_length=255, blank=True, null=True)
    valorfrete = models.CharField(max_length=255, blank=True, null=True)
    valordesconto = models.CharField(max_length=255, blank=True, null=True)
    valornfe = models.CharField(max_length=255, blank=True, null=True)
    valordevolucao = models.CharField(max_length=255, blank=True, null=True)
    valordespesasacessorias = models.CharField(max_length=255, blank=True, null=True)
    valoriss = models.CharField(max_length=255, blank=True, null=True)
    valoricmsdesonerado = models.CharField(max_length=255, blank=True, null=True)
    valoricmsoutriscredc = models.CharField(max_length=255, blank=True, null=True)
    valoricmsoutrisdesc = models.CharField(max_length=255, blank=True, null=True)
    valoricmsoutris = models.CharField(max_length=255, blank=True, null=True)
    valoricmsendebitoc = models.CharField(max_length=255, blank=True, null=True)
    valoricmsendebitos = models.CharField(max_length=255, blank=True, null=True)
    valoricmsenc = models.CharField(max_length=255, blank=True, null=True)
    valoricmsens = models.CharField(max_length=255, blank=True, null=True)
    valoricmsefetivo = models.CharField(max_length=255, blank=True, null=True)
    valoricmsgare = models.CharField(max_length=255, blank=True, null=True)
    valoricmsstgar = models.CharField(max_length=255, blank=True, null=True)
    valortotalpagamento = models.CharField(max_length=255, blank=True, null=True)
    valorirrfbc = models.CharField(max_length=255, blank=True, null=True)
    valorirrfretido = models.CharField(max_length=255, blank=True, null=True)
    valorprevidenciabc = models.CharField(max_length=255, blank=True, null=True)
    valorprevidenciaretido = models.CharField(max_length=255, blank=True, null=True)
    valortrocoformapagamento = models.CharField(max_length=255, blank=True, null=True)
    valoricsmmono = models.CharField(max_length=255, blank=True, null=True)
    valorqtdtributadamonoreten = models.CharField(max_length=255, blank=True, null=True)
    valoricmsmonoretenanterior = models.CharField(max_length=255, blank=True, null=True)
    valorqtdtributadamonoretenanterior = models.CharField(max_length=255, blank=True, null=True)
    valoricmsmonoretidaanterior = models.CharField(max_length=255, blank=True, null=True)
    ultimoeventodestinatario = models.CharField(max_length=255, blank=True, null=True)

    class Meta:
        verbose_name = 'Documentos'
        verbose_name_plural = 'Documentos'

from django.db import models
from django.db.models import Q

class SatRegistro(models.Model):
    data_emissao = models.DateField(db_index=True, blank=True, null=True)

    competencia = models.DateField(db_index=True, blank=True, null=True)

    empresa = models.ForeignKey(Empresa, on_delete=models.CASCADE, related_name="sat_registros")
    sheet   = models.CharField("Aba", max_length=100, db_index=True)
    row     = models.PositiveIntegerField("Linha (1-based)")
    data    = models.JSONField("Dados da linha")
    descricao  = models.CharField(max_length=255, null=True, blank=True, db_index=True)
    ncm        = models.CharField(max_length=20,  null=True, blank=True, db_index=True)
    cfop       = models.CharField(max_length=10,  null=True, blank=True, db_index=True)
    cest       = models.CharField(max_length=10,  null=True, blank=True, db_index=True)
    cst_csosn  = models.CharField(max_length=10,  null=True, blank=True, db_index=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Registro SAT (linha)"
        verbose_name_plural = "Registros SAT (linhas)"
        constraints = [
            models.UniqueConstraint(
                fields=["empresa", "competencia", "sheet", "row"],
                name="uniq_emp_comp_sheet_row",
                condition=~Q(competencia__isnull=True),
            ),
        ]
        indexes = [
            models.Index(fields=["empresa", "sheet"]),
            models.Index(fields=["empresa", "competencia"]),
            models.Index(fields=["empresa", "data_emissao"]),
            models.Index(fields=["empresa", "descricao"]),
            models.Index(fields=["empresa", "ncm"]),
            models.Index(fields=["empresa", "cfop"]),
            models.Index(fields=["empresa", "cest"]),
            models.Index(fields=["empresa", "cst_csosn"]),
        ]
