"""
Bot de controle de gastos via Telegram — versão com Google Sheets.

Estrutura esperada na planilha do Google (duas abas):

Aba "Orcamentos" (nome configurável via .env):
    A: Categoria        B: OrcamentoMensal
    Alimentação         1200
    Transporte          400
    ...

Aba "Gastos" (nome configurável via .env) — é o "banco de dados" de lançamentos,
alimentada automaticamente pelo bot a cada /gasto:
    A: Data              B: Categoria      C: Valor
    2026-09-28 14:03:11  Alimentação       45.90

Na primeira execução, se a planilha estiver vazia, o bot cria as duas abas
automaticamente e semeia a aba "Orcamentos" com valores padrão (definidos em
DEFAULT_BUDGETS abaixo) — assim já dá para testar sem editar nada na mão.
Depois disso, os orçamentos passam a ser 100% editáveis diretamente na planilha
(sem precisar tocar em código), e todo o histórico de gastos fica salvo na nuvem
— ou seja, o bot pode reiniciar, trocar de servidor, etc., sem perder nada.

O saldo de cada categoria é calculado somando os lançamentos de "Gastos" que
caem no mês/ano atual daquela categoria, e subtraindo do orçamento mensal.

--------------------------------------------------------------------------
CONFIGURAÇÃO (arquivo .env — veja .env.example)
--------------------------------------------------------------------------
TELEGRAM_BOT_TOKEN        -> token do bot, obtido com o @BotFather
GOOGLE_SHEET_ID           -> ID da planilha (fica na URL, entre /d/ e /edit)
GOOGLE_CREDENTIALS_FILE   -> caminho do JSON da conta de serviço do Google
SHEET_ORCAMENTOS_NAME     -> (opcional) nome da aba de orçamentos. Padrão: Orcamentos
SHEET_GASTOS_NAME         -> (opcional) nome da aba de gastos. Padrão: Gastos

--------------------------------------------------------------------------
COMO CONECTAR AO GOOGLE SHEETS (conta de serviço)
--------------------------------------------------------------------------
1. No Google Cloud Console, crie/abra um projeto e ative as APIs:
   "Google Sheets API" e "Google Drive API".
2. Crie uma "Service Account" (Conta de serviço) e gere uma chave em formato
   JSON. Salve esse arquivo (ex: credentials.json) na mesma pasta do bot.
3. Abra o JSON e copie o valor de "client_email"
   (algo como xxxx@xxxx.iam.gserviceaccount.com).
4. Na planilha do Google, clique em "Compartilhar" e adicione esse e-mail
   como Editor.
5. Preencha o .env com o caminho do JSON e o ID da planilha.

--------------------------------------------------------------------------
Instalação:
    pip install -r requirements.txt

Como rodar:
    python bot_gastos.py
"""

import logging
import os
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes

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
ALLOWED_USERS = [u.strip() for u in os.getenv("ALLOWED_USERS", "rauhmones").split(",") if u.strip()]

if not TELEGRAM_BOT_TOKEN:
    raise RuntimeError("Defina TELEGRAM_BOT_TOKEN no arquivo .env")
if not GOOGLE_SHEET_ID:
    raise RuntimeError("Defina GOOGLE_SHEET_ID no arquivo .env")

# Usados apenas para semear a aba de orçamentos quando ela estiver vazia.
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

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


# --------------------------------------------------------------------------
# CAMADA DE ACESSO AO GOOGLE SHEETS
# --------------------------------------------------------------------------
class PlanilhaGastos:
    """Encapsula toda a leitura/escrita na planilha do Google."""

    def __init__(self) -> None:
        creds = Credentials.from_service_account_file(
            GOOGLE_CREDENTIALS_FILE, scopes=SCOPES
        )
        cliente = gspread.authorize(creds)
        self.planilha = cliente.open_by_key(GOOGLE_SHEET_ID)
        self.aba_orcamentos = self._garantir_aba(
            SHEET_ORCAMENTOS_NAME, ["Categoria", "OrcamentoMensal"]
        )
        # Garantir que a aba Gastos tenha a coluna adicional "Usuario"
        self.aba_gastos = self._garantir_aba(
            SHEET_GASTOS_NAME, ["Data", "Categoria", "Valor", "Usuario"]
        )
        self._semear_orcamentos_se_vazio()

    def _garantir_aba(self, nome: str, cabecalho: List[str]):
        """Retorna a worksheet pelo nome, criando-a (com cabeçalho) se não existir.

        Se a aba já existir, garante que todas as colunas do `cabecalho` estejam
        presentes (adiciona as que estiverem faltando)."""
        try:
            aba = self.planilha.worksheet(nome)
            # verificar e adicionar colunas faltantes no cabeçalho
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

    def _semear_orcamentos_se_vazio(self) -> None:
        """Se a aba de orçamentos só tiver o cabeçalho, popula com DEFAULT_BUDGETS."""
        valores = self.aba_orcamentos.get_all_values()
        if len(valores) <= 1:
            logger.info("Aba de orçamentos vazia — semeando valores padrão.")
            linhas = [[cat, valor] for cat, valor in DEFAULT_BUDGETS.items()]
            self.aba_orcamentos.append_rows(linhas)

    # ---------------------- ORÇAMENTOS ----------------------
    def ler_orcamentos(self) -> Dict[str, float]:
        registros = self.aba_orcamentos.get_all_records()
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

    def encontrar_categoria(self, nome_digitado: str) -> Optional[str]:
        nome_normalizado = nome_digitado.strip().lower()
        for categoria in self.ler_orcamentos():
            if categoria.lower() == nome_normalizado:
                return categoria
        return None

    # ---------------------- GASTOS ----------------------
    def registrar_gasto(self, categoria: str, valor: float, usuario: Optional[str] = None) -> None:
        agora = datetime.now().strftime(DATE_FORMAT)
        linha = [agora, categoria, valor, usuario or ""]
        self.aba_gastos.append_row(linha)

    def gasto_do_mes_por_categoria(self, categoria: str) -> float:
        agora = datetime.now()
        total = 0.0
        for linha in self.aba_gastos.get_all_records():
            if str(linha.get("Categoria", "")).strip().lower() != categoria.lower():
                continue
            data_str = str(linha.get("Data", ""))
            try:
                data_lancamento = datetime.strptime(data_str, DATE_FORMAT)
            except ValueError:
                continue
            if data_lancamento.month == agora.month and data_lancamento.year == agora.year:
                try:
                    total += float(linha.get("Valor", 0) or 0)
                except (TypeError, ValueError):
                    pass
        return total

    def ultimos_lancamentos(self, limite: int = 15) -> List[Tuple[str, str, float, str]]:
        registros = self.aba_gastos.get_all_records()
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


def _formatar_reais(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


async def verificar_permissao(update: Update) -> bool:
    """Verifica se o usuário que enviou o update está na lista ALLOWED_USERS.

    Suporta comparação com o ID numérico do Telegram (str) ou com o
    nome de usuário (username) sem @.
    """
    user = getattr(update, "effective_user", None)
    if not user:
        return False
    user_id_str = str(user.id)
    username = (user.username or "").strip()
    for item in ALLOWED_USERS:
        if item == user_id_str or (username and (item == username or item == f"@{username}")):
            return True
    # se não autorizado, avisa o usuário quando possível
    if getattr(update, "effective_message", None):
        await update.effective_message.reply_text("❌ Você não está autorizado a usar este bot.")
    return False


# --------------------------------------------------------------------------
# HANDLERS DO TELEGRAM
# --------------------------------------------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await verificar_permissao(update):
        return
    mensagem = (
        "👋 Olá! Eu sou seu bot de controle de gastos (dados salvos no Google Sheets).\n\n"
        "Comandos disponíveis:\n"
        "• /gasto <categoria> <valor> — registra um gasto\n"
        "   Ex: /gasto Alimentação 45,90\n"
        "• /categorias — lista as categorias e orçamentos totais\n"
        "• /orcamento — mostra o saldo restante de todas as categorias (mês atual)\n"
        "• /orcamento <categoria> — mostra o saldo restante de uma categoria\n"
        "• /extrato — lista os últimos gastos registrados\n"
    )
    await update.message.reply_text(mensagem)


async def gasto(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await verificar_permissao(update):
        return
    args = context.args
    if len(args) < 2:
        await update.message.reply_text(
            "Uso correto: /gasto <categoria> <valor>\nEx: /gasto Lazer 50,00"
        )
        return

    valor_bruto = args[-1]
    categoria_digitada = " ".join(args[:-1])

    try:
        valor = float(valor_bruto.replace(".", "").replace(",", "."))
    except ValueError:
        await update.message.reply_text(
            f"❌ Não consegui entender o valor '{valor_bruto}'. Use algo como 45,90."
        )
        return

    if valor <= 0:
        await update.message.reply_text("❌ O valor precisa ser positivo.")
        return

    orcamentos = planilha.ler_orcamentos()
    categoria = planilha.encontrar_categoria(categoria_digitada)
    if categoria is None:
        categorias_disponiveis = ", ".join(orcamentos.keys())
        await update.message.reply_text(
            f"❌ Categoria '{categoria_digitada}' não encontrada.\n"
            f"Categorias disponíveis: {categorias_disponiveis}"
        )
        return

    # identificar usuário que registrou o gasto
    user = update.effective_user
    user_identifier = (user.username or "").strip() if user and user.username else str(user.id)

    planilha.registrar_gasto(categoria, valor, user_identifier)

    gasto_mes = planilha.gasto_do_mes_por_categoria(categoria)
    restante = orcamentos[categoria] - gasto_mes
    status = "✅" if restante >= 0 else "⚠️"

    await update.message.reply_text(
        f"{status} Gasto registrado: {categoria} — {_formatar_reais(valor)}\n"
        f"Orçamento restante em {categoria}: {_formatar_reais(restante)} "
        f"(de {_formatar_reais(orcamentos[categoria])})"
    )


async def categorias(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await verificar_permissao(update):
        return
    orcamentos = planilha.ler_orcamentos()
    linhas = ["📂 Categorias e orçamento mensal total:"]
    for cat, orcamento_total in orcamentos.items():
        linhas.append(f"• {cat}: {_formatar_reais(orcamento_total)}")
    await update.message.reply_text("\n".join(linhas))


async def orcamento(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await verificar_permissao(update):
        return
    args = context.args
    orcamentos = planilha.ler_orcamentos()

    if not args:
        linhas = ["💰 Saldo do orçamento por categoria (mês atual):"]
        total_orcado = 0.0
        total_gasto = 0.0
        for cat, orcamento_total in orcamentos.items():
            gasto_mes = planilha.gasto_do_mes_por_categoria(cat)
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
        await update.message.reply_text("\n".join(linhas))
        return

    categoria_digitada = " ".join(args)
    categoria = planilha.encontrar_categoria(categoria_digitada)
    if categoria is None:
        categorias_disponiveis = ", ".join(orcamentos.keys())
        await update.message.reply_text(
            f"❌ Categoria '{categoria_digitada}' não encontrada.\n"
            f"Categorias disponíveis: {categorias_disponiveis}"
        )
        return

    gasto_mes = planilha.gasto_do_mes_por_categoria(categoria)
    restante = orcamentos[categoria] - gasto_mes
    await update.message.reply_text(
        f"💰 {categoria}: {_formatar_reais(restante)} restante "
        f"de {_formatar_reais(orcamentos[categoria])} "
        f"(gasto neste mês: {_formatar_reais(gasto_mes)})"
    )


async def extrato(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await verificar_permissao(update):
        return
    lancamentos = planilha.ultimos_lancamentos(limite=15)
    if not lancamentos:
        await update.message.reply_text("Nenhum gasto registrado ainda.")
        return

    linhas = ["🧾 Últimos gastos registrados:"]
    for data_str, cat, valor, usuario in reversed(lancamentos):
        usuario_txt = f" — {usuario}" if usuario else ""
        linhas.append(f"• {data_str} — {cat}: {_formatar_reais(valor)}{usuario_txt}")
    await update.message.reply_text("\n".join(linhas))


async def erro_desconhecido(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Não reconheci esse comando. Digite /start para ver a lista de comandos."
    )


async def tratar_erro_global(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Captura qualquer exceção não tratada nos handlers.

    Sem isso, um erro inesperado (ex: instabilidade momentânea na API do
    Google Sheets, célula com formato inesperado, etc.) fica só registrado no
    log e o bot continua funcionando normalmente para os próximos comandos —
    em vez de deixar o processo em um estado instável.
    """
    logger.error("Exceção não tratada ao processar um update:", exc_info=context.error)
    if isinstance(update, Update) and update.effective_message:
        await update.effective_message.reply_text(
            "⚠️ Ocorreu um erro inesperado ao processar esse comando. "
            "Tente novamente em alguns segundos."
        )


# --------------------------------------------------------------------------
# INICIALIZAÇÃO
# --------------------------------------------------------------------------
def main() -> None:
    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("gasto", gasto))
    app.add_handler(CommandHandler("categorias", categorias))
    app.add_handler(CommandHandler("orcamento", orcamento))
    app.add_handler(CommandHandler("extrato", extrato))
    app.add_error_handler(tratar_erro_global)

    logger.info("Bot iniciado. Aguardando mensagens...")
    # drop_pending_updates evita que updates acumulados de uma instância
    # anterior (ex: se o processo antigo não foi encerrado corretamente)
    # sejam reprocessados e causem o erro "Conflict: terminated by other
    # getUpdates request".
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()