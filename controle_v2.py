"""
Bot de controle de gastos via Telegram — versão INTERATIVA com Google Sheets,
com controle separado por usuário e um controle compartilhado da Família.

--------------------------------------------------------------------------
COMO FUNCIONA (fluxo em árvore)
--------------------------------------------------------------------------
Cada comando abre uma sequência de perguntas com botões:

    /gasto       -> [Pessoal | Família] -> [Categoria] -> digita o valor
    /orcamento   -> [Pessoal | Família] -> [Todas | Categoria]
    /categorias  -> [Pessoal | Família]
    /extrato     -> [Pessoal | Família]

Em qualquer passo dá para tocar em "❌ Cancelar" ou enviar /cancelar.
Enviar outro comando no meio do fluxo reinicia a conversa.

ATALHOS (o formato antigo continua funcionando)
    /gasto [-F] <categoria> <valor>     ex: /gasto -F Farmácia 200
    /orcamento [-F] [categoria]
    /categorias -F
    /extrato -F
Se o comando vier com -F, a pergunta "Pessoal ou Família?" é pulada.

--------------------------------------------------------------------------
ESTRUTURA DA PLANILHA
--------------------------------------------------------------------------
Cada usuário tem o seu próprio par de abas, criado automaticamente na primeira
vez em que ele usa o bot:

    Orcamentos_<usuario>      Gastos_<usuario>

E existe um par de abas compartilhado da família:

    Orcamentos_Familia        Gastos_Familia

Aba de orçamentos:   A: Categoria | B: OrcamentoMensal
Aba de gastos:       A: Data | B: Categoria | C: Valor | D: Usuario

Toda aba nova de orçamentos é semeada com DEFAULT_BUDGETS; depois disso os
valores (e as categorias) são 100% editáveis direto na planilha.

--------------------------------------------------------------------------
MIGRAÇÃO DAS ABAS ANTIGAS
--------------------------------------------------------------------------
Se existirem as abas antigas "Orcamentos" e "Gastos" (versão de usuário único),
elas são RENOMEADAS na inicialização para "Orcamentos_<LEGACY_OWNER>" e
"Gastos_<LEGACY_OWNER>" — sem perder nenhum dado. O dono padrão é "rauhmones".
Para desativar, defina LEGACY_OWNER= (vazio) no .env.

--------------------------------------------------------------------------
CONFIGURAÇÃO (arquivo .env)
--------------------------------------------------------------------------
TELEGRAM_BOT_TOKEN        -> token do bot, obtido com o @BotFather
GOOGLE_SHEET_ID           -> ID da planilha (fica na URL, entre /d/ e /edit)
GOOGLE_CREDENTIALS_FILE   -> caminho do JSON da conta de serviço do Google
SHEET_ORCAMENTOS_NAME     -> (opcional) prefixo das abas de orçamentos. Padrão: Orcamentos
SHEET_GASTOS_NAME         -> (opcional) prefixo das abas de gastos. Padrão: Gastos
ALLOWED_USERS             -> (opcional) usuários autorizados (IDs ou usernames)
LEGACY_OWNER              -> (opcional) dono das abas antigas. Padrão: rauhmones

Instalação:  pip install -r requirements.txt
Execução:    python bot_gastos.py
"""

import logging
import os
import re
from datetime import datetime
from typing import Dict, List, NamedTuple, Optional, Tuple

import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from telegram import BotCommand, InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

# --------------------------------------------------------------------------
# CONFIGURAÇÃO INICIAL (via .env)
# --------------------------------------------------------------------------
load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")
SHEET_ORCAMENTOS_NAME = os.getenv("SHEET_ORCAMENTOS_NAME", "Orcamentos")
SHEET_GASTOS_NAME = os.getenv("SHEET_GASTOS_NAME", "Gastos")
# Lista de usuários autorizados. Pode conter IDs numéricos do Telegram ou
# nomes de usuário (sem @), separados por vírgula. Ex: "123456,anotheruser"
ALLOWED_USERS = [u.strip() for u in os.getenv("ALLOWED_USERS", "rauhmones,992630990").split(",") if u.strip()]
# Dono das abas "Orcamentos"/"Gastos" da versão antiga (usuário único).
LEGACY_OWNER = os.getenv("LEGACY_OWNER", "rauhmones").strip()

if not TELEGRAM_BOT_TOKEN:
    raise RuntimeError("Defina TELEGRAM_BOT_TOKEN no arquivo .env")
if not GOOGLE_SHEET_ID:
    raise RuntimeError("Defina GOOGLE_SHEET_ID no arquivo .env")

# Usados apenas para semear abas de orçamento novas (pessoais e da família).
DEFAULT_BUDGETS: Dict[str, float] = {
    "Alimentação": 1200.00,
    "Transporte": 400.00,
    "Lazer": 300.00,
    "Moradia": 2000.00,
    "Saúde": 300.00,
    "Educação": 250.00,
    "Outros": 200.00,
}

DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

CABECALHO_ORCAMENTOS = ["Categoria", "OrcamentoMensal"]
CABECALHO_GASTOS = ["Data", "Categoria", "Valor", "Usuario"]

# "Escopo" = de quem é o controle. Para usuários é o username/ID; para a
# família é este valor fixo.
FAMILIA_ESCOPO = "Familia"
FLAG_FAMILIA = "-f"  # comparada em minúsculas, então -F também funciona

_CHARS_INVALIDOS_ABA = re.compile(r"[\[\]:*?/\\]")

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


def normalizar_escopo(chave: str) -> str:
    """Transforma um username/ID em um sufixo seguro para nome de aba."""
    limpo = _CHARS_INVALIDOS_ABA.sub("_", chave.strip())[:80] or "sem_nome"
    # Evita que um usuário chamado "familia" colida com as abas da família.
    if limpo.lower() == FAMILIA_ESCOPO.lower():
        limpo += "_usuario"
    return limpo


def nome_aba_orcamentos(escopo: str) -> str:
    return f"{SHEET_ORCAMENTOS_NAME}_{escopo}"


def nome_aba_gastos(escopo: str) -> str:
    return f"{SHEET_GASTOS_NAME}_{escopo}"


# --------------------------------------------------------------------------
# CAMADA DE ACESSO AO GOOGLE SHEETS (inalterada)
# --------------------------------------------------------------------------
class PlanilhaGastos:
    """Encapsula toda a leitura/escrita na planilha do Google.

    Todos os métodos públicos recebem o `escopo` (username/ID normalizado, ou
    FAMILIA_ESCOPO) e criam as abas correspondentes automaticamente se ainda
    não existirem.
    """

    def __init__(self) -> None:
        creds = Credentials.from_service_account_file(
            GOOGLE_CREDENTIALS_FILE, scopes=SCOPES
        )
        cliente = gspread.authorize(creds)
        self.planilha = cliente.open_by_key(GOOGLE_SHEET_ID)
        # cache: escopo -> (aba_orcamentos, aba_gastos)
        self._abas: Dict[str, Tuple[gspread.Worksheet, gspread.Worksheet]] = {}

        self._migrar_abas_legadas()
        self.garantir_escopo(FAMILIA_ESCOPO)  # a aba da família sempre existe

    # ---------------------- ABAS ----------------------
    def _migrar_abas_legadas(self) -> None:
        """Renomeia as abas da versão antiga (usuário único) para o novo padrão."""
        if not LEGACY_OWNER:
            return
        escopo = normalizar_escopo(LEGACY_OWNER)
        existentes = {ws.title for ws in self.planilha.worksheets()}
        pares = (
            (SHEET_ORCAMENTOS_NAME, nome_aba_orcamentos(escopo)),
            (SHEET_GASTOS_NAME, nome_aba_gastos(escopo)),
        )
        for antigo, novo in pares:
            if antigo in existentes and novo not in existentes:
                logger.info("Migrando aba legada '%s' -> '%s'", antigo, novo)
                self.planilha.worksheet(antigo).update_title(novo)

    def _garantir_aba(self, nome: str, cabecalho: List[str]):
        """Retorna a worksheet pelo nome, criando-a (com cabeçalho) se não existir.

        Se a aba já existir, garante que todas as colunas do `cabecalho` estejam
        presentes (adiciona as que estiverem faltando)."""
        try:
            aba = self.planilha.worksheet(nome)
            existente = aba.row_values(1)
            missing = [h for h in cabecalho if h not in existente]
            if missing:
                logger.info("Atualizando cabeçalho da aba '%s', adicionando colunas: %s", nome, missing)
                current_len = len(existente)
                for i, col in enumerate(missing, start=1):
                    aba.update_cell(1, current_len + i, col)
        except gspread.exceptions.WorksheetNotFound:
            logger.info("Criando aba '%s'...", nome)
            aba = self.planilha.add_worksheet(title=nome, rows=200, cols=max(len(cabecalho), 3))
            aba.append_row(cabecalho)
        return aba

    def garantir_escopo(self, escopo: str):
        """Retorna (aba_orcamentos, aba_gastos) do escopo, criando o que faltar."""
        if escopo in self._abas:
            return self._abas[escopo]

        aba_orcamentos = self._garantir_aba(nome_aba_orcamentos(escopo), CABECALHO_ORCAMENTOS)
        aba_gastos = self._garantir_aba(nome_aba_gastos(escopo), CABECALHO_GASTOS)
        self._semear_orcamentos_se_vazio(aba_orcamentos)

        self._abas[escopo] = (aba_orcamentos, aba_gastos)
        return self._abas[escopo]

    @staticmethod
    def _semear_orcamentos_se_vazio(aba_orcamentos) -> None:
        """Se a aba de orçamentos só tiver o cabeçalho, popula com DEFAULT_BUDGETS."""
        valores = aba_orcamentos.get_all_values()
        if len(valores) <= 1:
            logger.info("Aba '%s' vazia — semeando valores padrão.", aba_orcamentos.title)
            linhas = [[cat, valor] for cat, valor in DEFAULT_BUDGETS.items()]
            aba_orcamentos.append_rows(linhas)

    # ---------------------- ORÇAMENTOS ----------------------
    def ler_orcamentos(self, escopo: str) -> Dict[str, float]:
        aba_orcamentos, _ = self.garantir_escopo(escopo)
        registros = aba_orcamentos.get_all_records()
        orcamentos = {}
        for linha in registros:
            categoria = str(linha.get("Categoria", "")).strip()
            if not categoria:
                continue
            try:
                orcamentos[categoria] = float(linha.get("OrcamentoMensal", 0) or 0)
            except (TypeError, ValueError):
                orcamentos[categoria] = 0.0
        return orcamentos

    @staticmethod
    def encontrar_categoria(nome_digitado: str, orcamentos: Dict[str, float]) -> Optional[str]:
        nome_normalizado = nome_digitado.strip().lower()
        for categoria in orcamentos:
            if categoria.lower() == nome_normalizado:
                return categoria
        return None

    # ---------------------- GASTOS ----------------------
    def registrar_gasto(
        self, escopo: str, categoria: str, valor: float, usuario: Optional[str] = None
    ) -> None:
        _, aba_gastos = self.garantir_escopo(escopo)
        agora = datetime.now().strftime(DATE_FORMAT)
        aba_gastos.append_row([agora, categoria, valor, usuario or ""])

    def gastos_do_mes(self, escopo: str) -> Dict[str, float]:
        """Total gasto no mês atual, por categoria (chave em minúsculas).

        Faz UMA leitura da aba, em vez de uma por categoria — importante para
        não estourar a cota da API do Google Sheets com vários usuários."""
        _, aba_gastos = self.garantir_escopo(escopo)
        agora = datetime.now()
        totais: Dict[str, float] = {}
        for linha in aba_gastos.get_all_records():
            data_str = str(linha.get("Data", ""))
            try:
                data_lancamento = datetime.strptime(data_str, DATE_FORMAT)
            except ValueError:
                continue
            if data_lancamento.month != agora.month or data_lancamento.year != agora.year:
                continue
            try:
                valor = float(linha.get("Valor", 0) or 0)
            except (TypeError, ValueError):
                continue
            chave = str(linha.get("Categoria", "")).strip().lower()
            totais[chave] = totais.get(chave, 0.0) + valor
        return totais

    def gasto_do_mes_por_categoria(self, escopo: str, categoria: str) -> float:
        return self.gastos_do_mes(escopo).get(categoria.strip().lower(), 0.0)

    def ultimos_lancamentos(self, escopo: str, limite: int = 15) -> List[Tuple[str, str, float, str]]:
        _, aba_gastos = self.garantir_escopo(escopo)
        registros = aba_gastos.get_all_records()
        recentes = registros[-limite:]
        resultado = []
        for linha in recentes:
            try:
                valor = float(linha.get("Valor", 0) or 0)
            except (TypeError, ValueError):
                valor = 0.0
            usuario = str(linha.get("Usuario", ""))
            resultado.append(
                (str(linha.get("Data", "")), str(linha.get("Categoria", "")), valor, usuario)
            )
        return resultado


# Instância única, criada na inicialização do processo.
planilha = PlanilhaGastos()


# --------------------------------------------------------------------------
# UTILITÁRIOS
# --------------------------------------------------------------------------
def _formatar_reais(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _identificador_usuario(user) -> str:
    """Username (sem @) se existir; caso contrário, o ID numérico do Telegram."""
    if user and user.username:
        return user.username.strip()
    return str(user.id)


def _parse_valor(bruto: str) -> Optional[float]:
    """Interpreta '45,90', '1.200,50', '45.90' ou '1.200'. Retorna None se inválido."""
    s = bruto.strip().replace("R$", "").replace(" ", "")
    if not s:
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")  # "1.200" -> milhar
    try:
        return float(s)
    except ValueError:
        return None


class Contexto(NamedTuple):
    """Resultado de interpretar quem enviou o comando e qual controle usar."""

    escopo: str          # qual par de abas usar
    familiar: bool       # True se o controle é o da família
    args: List[str]      # argumentos sem a flag -F
    usuario: str         # quem enviou a mensagem (vai na coluna "Usuario")
    titulo: str          # rótulo amigável para as respostas


def _montar_contexto(usuario: str, familiar: bool, args: Optional[List[str]] = None) -> Contexto:
    args = args or []
    if familiar:
        return Contexto(FAMILIA_ESCOPO, True, args, usuario, "👨‍👩‍👧 Família")
    return Contexto(normalizar_escopo(usuario), False, args, usuario, f"👤 {usuario}")


def _contexto(update: Update, args: List[str]) -> Contexto:
    """Interpreta um comando de atalho (com -F opcional)."""
    usuario = _identificador_usuario(update.effective_user)
    familiar = any(a.lower() == FLAG_FAMILIA for a in args)
    args_limpos = [a for a in args if a.lower() != FLAG_FAMILIA]
    return _montar_contexto(usuario, familiar, args_limpos)


async def verificar_permissao(update: Update) -> bool:
    """Verifica se o usuário que enviou o update está na lista ALLOWED_USERS."""
    user = getattr(update, "effective_user", None)
    if not user:
        return False
    user_id_str = str(user.id)
    username = (user.username or "").strip()
    for item in ALLOWED_USERS:
        if item == user_id_str or (username and (item == username or item == f"@{username}")):
            return True
    if getattr(update, "effective_message", None):
        await update.effective_message.reply_text("❌ Você não está autorizado a usar este bot.")
    return False


async def _enviar(update: Update, texto: str, teclado: Optional[InlineKeyboardMarkup] = None) -> None:
    """Responde à mensagem (comando) ou edita a mensagem do botão tocado."""
    query = update.callback_query
    if query:
        await query.answer()
        await query.edit_message_text(texto, reply_markup=teclado)
    else:
        await update.message.reply_text(texto, reply_markup=teclado)


# --------------------------------------------------------------------------
# TECLADOS (botões inline)
# --------------------------------------------------------------------------
BOTAO_CANCELAR = InlineKeyboardButton("❌ Cancelar", callback_data="cancelar")


def _teclado_escopo() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("👤 Pessoal", callback_data="esc:p"),
            InlineKeyboardButton("👨‍👩‍👧 Família", callback_data="esc:f"),
        ],
        [BOTAO_CANCELAR],
    ])


def _teclado_categorias(categorias: List[str], com_todas: bool = False) -> InlineKeyboardMarkup:
    # callback_data tem limite de 64 bytes; por isso usamos o índice da categoria.
    botoes = [InlineKeyboardButton(cat, callback_data=f"cat:{i}") for i, cat in enumerate(categorias)]
    linhas = [botoes[i:i + 2] for i in range(0, len(botoes), 2)]
    if com_todas:
        linhas.insert(0, [InlineKeyboardButton("📊 Todas as categorias", callback_data="cat:todas")])
    linhas.append([BOTAO_CANCELAR])
    return InlineKeyboardMarkup(linhas)


def _teclado_cancelar() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[BOTAO_CANCELAR]])


# --------------------------------------------------------------------------
# AÇÕES (lógica de cada comando, independente de como o usuário chegou nela)
# --------------------------------------------------------------------------
async def _acao_registrar_gasto(update: Update, ctx: Contexto, categoria: str, valor: float,
                                orcamentos: Dict[str, float]) -> None:
    planilha.registrar_gasto(ctx.escopo, categoria, valor, ctx.usuario)
    gasto_mes = planilha.gasto_do_mes_por_categoria(ctx.escopo, categoria)
    restante = orcamentos[categoria] - gasto_mes
    status = "✅" if restante >= 0 else "⚠️"
    await _enviar(
        update,
        f"{status} [{ctx.titulo}] Gasto registrado: {categoria} — {_formatar_reais(valor)}\n"
        f"Orçamento restante em {categoria}: {_formatar_reais(restante)} "
        f"(de {_formatar_reais(orcamentos[categoria])})",
    )


async def _acao_orcamento_todas(update: Update, ctx: Contexto) -> None:
    orcamentos = planilha.ler_orcamentos(ctx.escopo)
    gastos_mes = planilha.gastos_do_mes(ctx.escopo)
    linhas = [f"💰 [{ctx.titulo}] Saldo do orçamento por categoria (mês atual):"]
    total_orcado = 0.0
    total_gasto = 0.0
    for cat, orcamento_total in orcamentos.items():
        gasto_mes = gastos_mes.get(cat.lower(), 0.0)
        restante = orcamento_total - gasto_mes
        total_orcado += orcamento_total
        total_gasto += gasto_mes
        emoji = "✅" if restante >= 0 else "⚠️"
        linhas.append(f"{emoji} {cat}: {_formatar_reais(restante)} restante")
    linhas.append("")
    linhas.append(
        f"Total: {_formatar_reais(total_orcado - total_gasto)} restante "
        f"de {_formatar_reais(total_orcado)}"
    )
    await _enviar(update, "\n".join(linhas))


async def _acao_orcamento_categoria(update: Update, ctx: Contexto, categoria: str,
                                    orcamentos: Dict[str, float]) -> None:
    gasto_mes = planilha.gasto_do_mes_por_categoria(ctx.escopo, categoria)
    restante = orcamentos[categoria] - gasto_mes
    await _enviar(
        update,
        f"💰 [{ctx.titulo}] {categoria}: {_formatar_reais(restante)} restante "
        f"de {_formatar_reais(orcamentos[categoria])} "
        f"(gasto neste mês: {_formatar_reais(gasto_mes)})",
    )


async def _acao_categorias(update: Update, ctx: Contexto) -> None:
    orcamentos = planilha.ler_orcamentos(ctx.escopo)
    linhas = [f"📂 [{ctx.titulo}] Categorias e orçamento mensal total:"]
    for cat, orcamento_total in orcamentos.items():
        linhas.append(f"• {cat}: {_formatar_reais(orcamento_total)}")
    await _enviar(update, "\n".join(linhas))


async def _acao_extrato(update: Update, ctx: Contexto) -> None:
    lancamentos = planilha.ultimos_lancamentos(ctx.escopo, limite=15)
    if not lancamentos:
        await _enviar(update, f"[{ctx.titulo}] Nenhum gasto registrado ainda.")
        return
    linhas = [f"🧾 [{ctx.titulo}] Últimos gastos registrados:"]
    for data_str, cat, valor, usuario in reversed(lancamentos):
        usuario_txt = f" — {usuario}" if usuario else ""
        linhas.append(f"• {data_str} — {cat}: {_formatar_reais(valor)}{usuario_txt}")
    await _enviar(update, "\n".join(linhas))


# --------------------------------------------------------------------------
# FLUXO INTERATIVO (ConversationHandler)
# --------------------------------------------------------------------------
ESCOPO, CATEGORIA, VALOR = range(3)

ACAO_GASTO = "gasto"
ACAO_ORCAMENTO = "orcamento"
ACAO_CATEGORIAS = "categorias"
ACAO_EXTRATO = "extrato"

_TITULO_ACAO = {
    ACAO_GASTO: "💸 Novo gasto",
    ACAO_ORCAMENTO: "💰 Consultar orçamento",
    ACAO_CATEGORIAS: "📂 Categorias",
    ACAO_EXTRATO: "🧾 Extrato",
}


def _ctx_salvo(context: ContextTypes.DEFAULT_TYPE) -> Contexto:
    d = context.user_data
    return _montar_contexto(d["usuario"], d["familiar"])


async def _atalho(update: Update, acao: str, ctx: Contexto) -> bool:
    """Executa o comando direto se ele já veio completo. Retorna True se executou."""
    args = ctx.args

    if acao == ACAO_GASTO and len(args) >= 2:
        valor_bruto = args[-1]
        categoria_digitada = " ".join(args[:-1])
        valor = _parse_valor(valor_bruto)
        if valor is None:
            await update.message.reply_text(
                f"❌ Não consegui entender o valor '{valor_bruto}'. Use algo como 45,90."
            )
        elif valor <= 0:
            await update.message.reply_text("❌ O valor precisa ser positivo.")
        else:
            orcamentos = planilha.ler_orcamentos(ctx.escopo)
            categoria = planilha.encontrar_categoria(categoria_digitada, orcamentos)
            if categoria is None:
                await update.message.reply_text(
                    f"❌ Categoria '{categoria_digitada}' não encontrada.\n"
                    f"Categorias disponíveis: {', '.join(orcamentos.keys())}"
                )
            else:
                await _acao_registrar_gasto(update, ctx, categoria, valor, orcamentos)
        return True

    if acao == ACAO_ORCAMENTO and len(args) >= 1:
        categoria_digitada = " ".join(args)
        orcamentos = planilha.ler_orcamentos(ctx.escopo)
        categoria = planilha.encontrar_categoria(categoria_digitada, orcamentos)
        if categoria is None:
            await update.message.reply_text(
                f"❌ Categoria '{categoria_digitada}' não encontrada.\n"
                f"Categorias disponíveis: {', '.join(orcamentos.keys())}"
            )
        else:
            await _acao_orcamento_categoria(update, ctx, categoria, orcamentos)
        return True

    # /categorias -F e /extrato -F já sabem tudo o que precisam.
    if ctx.familiar and acao in (ACAO_CATEGORIAS, ACAO_EXTRATO):
        await (_acao_categorias if acao == ACAO_CATEGORIAS else _acao_extrato)(update, ctx)
        return True

    return False


async def _apos_escopo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Executado quando já sabemos se é Pessoal ou Família."""
    acao = context.user_data["acao"]
    ctx = _ctx_salvo(context)

    if acao == ACAO_CATEGORIAS:
        await _acao_categorias(update, ctx)
        return ConversationHandler.END
    if acao == ACAO_EXTRATO:
        await _acao_extrato(update, ctx)
        return ConversationHandler.END

    # gasto / orcamento: perguntar a categoria
    orcamentos = planilha.ler_orcamentos(ctx.escopo)
    context.user_data["cats"] = list(orcamentos.keys())
    pergunta = "Qual a categoria do gasto?" if acao == ACAO_GASTO else "Qual categoria você quer consultar?"
    await _enviar(
        update,
        f"{_TITULO_ACAO[acao]} — {ctx.titulo}\n\n{pergunta}",
        _teclado_categorias(context.user_data["cats"], com_todas=(acao == ACAO_ORCAMENTO)),
    )
    return CATEGORIA


def _iniciar(acao: str):
    """Fabrica o entry point de cada comando."""

    async def handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        if not await verificar_permissao(update):
            return ConversationHandler.END

        context.user_data.clear()
        ctx = _contexto(update, context.args or [])
        context.user_data.update(acao=acao, usuario=ctx.usuario, familiar=ctx.familiar)

        # Provisiona as abas do usuário no primeiro contato.
        planilha.garantir_escopo(normalizar_escopo(ctx.usuario))

        if await _atalho(update, acao, ctx):
            return ConversationHandler.END

        if ctx.familiar:  # veio com -F, então pula a pergunta de escopo
            return await _apos_escopo(update, context)

        await update.message.reply_text(
            f"{_TITULO_ACAO[acao]}\n\nDe qual controle?", reply_markup=_teclado_escopo()
        )
        return ESCOPO

    return handler


async def escolher_escopo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await verificar_permissao(update):
        return ConversationHandler.END
    context.user_data["familiar"] = update.callback_query.data == "esc:f"
    return await _apos_escopo(update, context)


async def escolher_categoria(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await verificar_permissao(update):
        return ConversationHandler.END
    ctx = _ctx_salvo(context)
    acao = context.user_data["acao"]
    escolha = update.callback_query.data.split(":", 1)[1]

    if escolha == "todas":
        await _acao_orcamento_todas(update, ctx)
        return ConversationHandler.END

    cats = context.user_data.get("cats", [])
    indice = int(escolha)
    if indice >= len(cats):  # botão de uma conversa antiga
        await _enviar(update, "⚠️ Essa lista expirou. Envie o comando novamente.")
        return ConversationHandler.END
    categoria = cats[indice]

    if acao == ACAO_ORCAMENTO:
        orcamentos = planilha.ler_orcamentos(ctx.escopo)
        await _acao_orcamento_categoria(update, ctx, categoria, orcamentos)
        return ConversationHandler.END

    # acao == gasto: falta o valor
    context.user_data["categoria"] = categoria
    await _enviar(
        update,
        f"{_TITULO_ACAO[ACAO_GASTO]} — {ctx.titulo}\n"
        f"Categoria: {categoria}\n\n"
        f"Digite o valor (ex: 45,90):",
        _teclado_cancelar(),
    )
    return VALOR


async def receber_valor(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await verificar_permissao(update):
        return ConversationHandler.END

    valor = _parse_valor(update.message.text)
    if valor is None:
        await update.message.reply_text(
            "❌ Não consegui entender esse valor. Digite algo como 45,90:",
            reply_markup=_teclado_cancelar(),
        )
        return VALOR
    if valor <= 0:
        await update.message.reply_text(
            "❌ O valor precisa ser positivo. Digite novamente:",
            reply_markup=_teclado_cancelar(),
        )
        return VALOR

    ctx = _ctx_salvo(context)
    categoria = context.user_data["categoria"]
    orcamentos = planilha.ler_orcamentos(ctx.escopo)
    if categoria not in orcamentos:  # categoria removida da planilha no meio do fluxo
        await update.message.reply_text("⚠️ Essa categoria não existe mais. Envie /gasto novamente.")
        return ConversationHandler.END

    await _acao_registrar_gasto(update, ctx, categoria, valor, orcamentos)
    return ConversationHandler.END


async def cancelar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await _enviar(update, "🚫 Operação cancelada.")
    return ConversationHandler.END


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await verificar_permissao(update):
        return
    planilha.garantir_escopo(normalizar_escopo(_identificador_usuario(update.effective_user)))
    await update.message.reply_text(
        "👋 Olá! Eu sou seu bot de controle de gastos (dados salvos no Google Sheets).\n\n"
        "Comandos disponíveis — cada um vai te guiando com botões:\n"
        "• /gasto — registra um gasto\n"
        "• /orcamento — saldo restante do mês (todas as categorias ou uma)\n"
        "• /categorias — categorias e orçamentos mensais\n"
        "• /extrato — últimos gastos registrados\n"
        "• /cancelar — cancela a operação em andamento\n\n"
        "Atalho: /gasto -F Farmácia 200 registra direto no controle da Família."
    )


async def tratar_erro_global(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Captura qualquer exceção não tratada nos handlers."""
    logger.error("Exceção não tratada ao processar um update:", exc_info=context.error)
    if isinstance(update, Update) and update.effective_message:
        await update.effective_message.reply_text(
            "⚠️ Ocorreu um erro inesperado ao processar esse comando. "
            "Tente novamente em alguns segundos."
        )


async def _registrar_menu(app: Application) -> None:
    """Preenche o menu de comandos (botão "/" do Telegram)."""
    await app.bot.set_my_commands([
        BotCommand("gasto", "Registrar um gasto"),
        BotCommand("orcamento", "Ver saldo do orçamento"),
        BotCommand("categorias", "Listar categorias"),
        BotCommand("extrato", "Últimos gastos"),
        BotCommand("cancelar", "Cancelar operação atual"),
    ])


# --------------------------------------------------------------------------
# INICIALIZAÇÃO
# --------------------------------------------------------------------------
def main() -> None:
    app = Application.builder().token(TELEGRAM_BOT_TOKEN).post_init(_registrar_menu).build()

    cb_cancelar = CallbackQueryHandler(cancelar, pattern=r"^cancelar$")

    fluxo = ConversationHandler(
        entry_points=[
            CommandHandler("gasto", _iniciar(ACAO_GASTO)),
            CommandHandler("orcamento", _iniciar(ACAO_ORCAMENTO)),
            CommandHandler("categorias", _iniciar(ACAO_CATEGORIAS)),
            CommandHandler("extrato", _iniciar(ACAO_EXTRATO)),
        ],
        states={
            ESCOPO: [
                CallbackQueryHandler(escolher_escopo, pattern=r"^esc:(p|f)$"),
                cb_cancelar,
            ],
            CATEGORIA: [
                CallbackQueryHandler(escolher_categoria, pattern=r"^cat:(\d+|todas)$"),
                cb_cancelar,
            ],
            VALOR: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receber_valor),
                cb_cancelar,
            ],
        },
        fallbacks=[CommandHandler("cancelar", cancelar), cb_cancelar],
        allow_reentry=True,  # um novo comando no meio do fluxo reinicia a conversa
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(fluxo)
    app.add_error_handler(tratar_erro_global)

    logger.info("Bot iniciado. Aguardando mensagens...")
    # drop_pending_updates evita reprocessar updates acumulados de uma
    # instância anterior ("Conflict: terminated by other getUpdates request").
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()