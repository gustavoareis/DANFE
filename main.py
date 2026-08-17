import re
import random
from datetime import datetime, timedelta
from pathlib import Path
import pdfplumber
from docx import Document

# Tenta importar bibliotecas de terceiros (Fallback mantido)
try:
    from num2words import num2words
    HAS_NUM2WORDS = True
except ImportError:
    HAS_NUM2WORDS = False

try:
    import holidays
    HAS_HOLIDAYS = True
except ImportError:
    HAS_HOLIDAYS = False
    print("Aviso: Biblioteca 'holidays' não encontrada. Feriados não serão descontados. Para instalar, rode: pip install holidays")

# Lista de meses (índice bate com o número do mês)
MESES = ["", "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", 
         "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]

# Constantes para conversão manual
UNIDADES = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove",
            "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"]
DEZENAS = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"]
CENTENAS = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"]

# Expressões regulares
RE_NF = re.compile(r"Nº\s*(\d+)")
RE_DATA_EMISSAO = re.compile(r"DATA DE EMISSÃO\n*(\d{2}/\d{2}/\d{4})")
RE_DATA_ANY = re.compile(r"(\d{2}/\d{2}/\d{4})")
RE_CNPJ = re.compile(r"CNPJ/CPF(?:[^\d]*)(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})")
RE_OBS = re.compile(r"OBS:\s*([^\n]+)", re.IGNORECASE)

def format_brl(valor: float) -> str:
    """Formata float para formato moeda BRL (ex: 1.234,56)."""
    return f"{valor:,.2f}".translate(str.maketrans(',.', '.,'))

def _converter_grupo(n: int) -> str:
    if n == 0: return ""
    if n == 100: return "cem"
    
    c, d, u = n // 100, (n % 100) // 10, n % 10
    partes = [CENTENAS[c]] if c > 0 else []
    
    if d == 1:
        partes.append(UNIDADES[d * 10 + u])
    else:
        if d > 1: partes.append(DEZENAS[d])
        if u > 0: partes.append(UNIDADES[u])
        
    return " e ".join(partes)

def valor_para_extenso_manual(valor: float) -> str:
    """Converte numérico para extenso (Fallback nativo)."""
    inteiro, centavos = int(valor), int(round((valor - int(valor)) * 100))
    if valor == 0: return "ZERO REAIS"
    
    partes = []
    if (m := inteiro // 1_000_000) > 0:
        partes.append(f"{_converter_grupo(m)} {'milhão' if m == 1 else 'milhões'}")
    if (k := (inteiro % 1_000_000) // 1_000) > 0:
        ext_k = _converter_grupo(k)
        partes.append("um mil" if ext_k == "um" else f"{ext_k} mil")
    if (u := inteiro % 1_000) > 0:
        partes.append(_converter_grupo(u))

    txt_int = ""
    if partes:
        txt_int = " ".join(partes[:-1]) + " e " + partes[-1] if len(partes) > 1 else partes[0]
        txt_int += " real" if inteiro == 1 else " reais"

    txt_cent = f"{_converter_grupo(centavos)} centavo{'s' if centavos > 1 else ''}" if centavos > 0 else ""
    return " e ".join(filter(None, [txt_int, txt_cent])).upper()

def valor_por_extenso(valor: float) -> str:
    if HAS_NUM2WORDS:
        try: return num2words(valor, lang='pt_BR', to='currency').upper()
        except Exception: pass
    return valor_para_extenso_manual(valor)

def is_dia_util(data_ref: datetime) -> bool:
    """Verifica se a data é um dia útil, excluindo fins de semana e feriados (BR, CE e Fortaleza)."""
    # 5 = Sábado, 6 = Domingo
    if data_ref.weekday() >= 5:
        return False
        
    if HAS_HOLIDAYS:
        # Carrega os feriados do Brasil e do estado do Ceará para o ano da data
        feriados = holidays.BR(subdiv='CE', years=data_ref.year)
        
        # Adiciona os feriados municipais fixos de Fortaleza
        feriados[datetime(data_ref.year, 4, 13).date()] = "Aniversário de Fortaleza"
        feriados[datetime(data_ref.year, 8, 15).date()] = "Nossa Senhora da Assunção"
        
        # Se a data estiver na lista de feriados, não é dia útil
        if data_ref.date() in feriados:
            return False
            
    return True

def extrair_dados_pdf(caminho_pdf: Path) -> dict:
    """Extrai informações e itens das páginas da NF."""
    dados = {"numero_nf": "000", "cliente": "", "cnpj": "", "data_emissao": datetime.now(), "itens": [], "obs": ""}
    texto_completo = ""

    with pdfplumber.open(caminho_pdf) as pdf:
        for page in pdf.pages:
            texto_completo += (page.extract_text() or "") + "\n"
            for table in page.extract_tables():
                for row in table:
                    if row and len(row) >= 9 and row[0] and row[0].isdigit():
                        try:
                            dados["itens"].append({
                                "descricao": (row[1] or "").replace("\n", " "),
                                "und": row[5] or "",
                                "qtd": float(row[6].translate(str.maketrans('', '', '.')) .replace(",", ".")),
                                "vlr_unit": float(row[7].translate(str.maketrans('', '', '.')) .replace(",", ".")),
                                "vlr_total": float(row[8].translate(str.maketrans('', '', '.')) .replace(",", "."))
                            })
                        except (ValueError, TypeError):
                            continue

    if match := RE_NF.search(texto_completo): dados["numero_nf"] = match.group(1)
    if match := RE_DATA_EMISSAO.search(texto_completo) or RE_DATA_ANY.search(texto_completo):
        dados["data_emissao"] = datetime.strptime(match.group(1), "%d/%m/%Y")
    if match := RE_CNPJ.search(texto_completo): dados["cnpj"] = match.group(1).strip()
    if match := RE_OBS.search(texto_completo): dados["obs"] = match.group(1).strip()

    linhas = texto_completo.split('\n')
    for i, linha in enumerate(linhas):
        if "NOME/RAZÃO SOCIAL" in linha:
            texto_restante = linha.split("NOME/RAZÃO SOCIAL")[-1].strip()
            if not texto_restante or "CNPJ" in texto_restante:
                if i + 1 < len(linhas):
                    cli = linhas[i+1].strip()
                    cli = re.sub(r'\s*\d{2}\.\d{3}\.\d{3}/.*', '', cli)
                    dados["cliente"] = cli.strip()
            else:
                dados["cliente"] = texto_restante
            break

    return dados

def reajustar_valores(itens: list, percentual: float = 0.03) -> tuple:
    """Aumenta o valor total, distribuindo nos itens de forma aleatória."""
    total_alvo = round(sum(i["vlr_total"] for i in itens) * (1 + percentual), 2)
    novos_itens, soma = [], 0.0

    for i in itens:
        fator = random.uniform(1.0 + max(0.005, percentual - 0.02), 1.0 + percentual + 0.02)
        vlr_unit = round(i["vlr_unit"] * fator, 2)
        vlr_total = round(vlr_unit * i["qtd"], 2)
        novos_itens.append({**i, "vlr_unit": vlr_unit, "vlr_total": vlr_total})
        soma += vlr_total

    if novos_itens and (dif := round(total_alvo - soma, 2)) != 0:
        ultimo = novos_itens[-1]
        ultimo["vlr_total"] = round(ultimo["vlr_total"] + dif, 2)
        if ultimo["qtd"] > 0:
            ultimo["vlr_unit"] = round(ultimo["vlr_total"] / ultimo["qtd"], 2)

    return novos_itens, total_alvo

def preencher_tabela(tabela, novos_itens: list):
    """Preenche tabela do Word baseada nas colunas dos cabeçalhos."""
    if not tabela.rows: return

    headers = [c.text.upper() for c in tabela.rows[0].cells]
    col = {
        "num": next((i for i, h in enumerate(headers) if any(x in h for x in ("ORDEM", "ITEM", "Nº"))), 0),
        "desc": next((i for i, h in enumerate(headers) if any(x in h for x in ("DESC", "PRODUTO", "SERVIÇO"))), 1),
        "und": next((i for i, h in enumerate(headers) if any(x in h for x in ("UND", "UNID"))), 2),
        "qtd": next((i for i, h in enumerate(headers) if any(x in h for x in ("QTD", "QUANT"))), 3),
        "unit": next((i for i, h in enumerate(headers) if any(x in h for x in ("UNIT", "PR. UNIT"))), 4),
        "total": next((i for i, h in enumerate(headers) if any(x in h for x in ("TOTAL", "PR. TOTAL"))), 5)
    }

    for idx, item in enumerate(novos_itens):
        if idx + 1 >= len(tabela.rows): tabela.add_row()
        cells = tabela.rows[idx + 1].cells
        
        if len(cells) > max(col.values()):
            cells[col["num"]].text = str(idx + 1)
            cells[col["desc"]].text = str(item["descricao"])
            cells[col["und"]].text = str(item["und"])
            cells[col["qtd"]].text = f"{item['qtd']:.0f}" if item['qtd'].is_integer() else f"{item['qtd']:.2f}".replace(".", ",")
            
            # Formatação atualizada com R$ nas tabelas
            cells[col["unit"]].text = f"R$ {format_brl(item['vlr_unit'])}"
            cells[col["total"]].text = f"R$ {format_brl(item['vlr_total'])}"

def aplicar_fonte_documento(doc, nome_fonte="Mongolian Baiti"):
    """Modifica fonte do doc inteiro usando achatamento de lista para performance."""
    if 'Normal' in doc.styles: doc.styles['Normal'].font.name = nome_fonte
    runs = [r for p in doc.paragraphs for r in p.runs] + \
           [r for t in doc.tables for row in t.rows for c in row.cells for p in c.paragraphs for r in p.runs]
    for run in runs: run.font.name = nome_fonte

def preencher_word(template_path: Path, output_path: Path, dados: dict, itens: list, total: float, data_fmt: str, is_jw=False):
    """Preenche e formata o documento Word."""
    doc = Document(template_path)
    val_str, ext_str = format_brl(total), valor_por_extenso(total)

    for p in doc.paragraphs:
        txt = p.text.upper()
        if txt.startswith("FORTALEZA") and " DE " in txt:
            p.text = f"{'FORTALEZA' if p.text.isupper() else 'Fortaleza'}, {data_fmt}"
        
        elif "CLIENTE:" in txt and dados.get("cliente"):
            partes = [f"Cliente: {dados['cliente']}"]
            if dados.get("obs"):
                partes.append(dados["obs"])
            if is_jw and dados.get("cnpj"):
                partes.append(f"CNPJ: {dados['cnpj']}")
            p.text = "\n".join(partes)
            
        elif "CNPJ:" in txt and not is_jw and dados.get("cnpj"):
            p.text = f"CNPJ: {dados['cnpj']}"
        
        elif is_jw and "VALOR DA PROPOSTA" in txt:
            p.text = f"Valor da Proposta R$ {val_str} ({ext_str})"
        elif not is_jw and "VALOR TOTAL DA PROPOSTA:" in txt:
            p.text = f"VALOR TOTAL DA PROPOSTA:  R$ {val_str}"

    if doc.tables:
        tabela = doc.tables[0]
        preencher_tabela(tabela, itens)
        if is_jw:
            for row in tabela.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for run in p.runs: run.bold = True

    aplicar_fonte_documento(doc)
    doc.save(output_path)

def processar_notas():
    entrada, saida = Path("notas_entrada"), Path("propostas_saida")
    tpl_lc = Path("template_lc.docx") if Path("template_lc.docx").exists() else Path("template.docx")
    tpl_jw = Path("template_jw.docx")

    if not entrada.exists():
        entrada.mkdir()
        return print(f"Pasta '{entrada}' criada. Adicione os PDFs nela e execute novamente.")

    saida.mkdir(exist_ok=True)
    if not tpl_lc.exists(): print(f"Aviso: Template '{tpl_lc}' não encontrado!")
    if not tpl_jw.exists(): print(f"Aviso: Template '{tpl_jw}' não encontrado!")

    arquivos = list(entrada.glob("*.pdf"))
    if not arquivos: return print("Nenhum arquivo PDF encontrado em 'notas_entrada'.")

    for arquivo in arquivos:
        print(f"\n--- Processando: {arquivo.name} ---")
        dados = extrair_dados_pdf(arquivo)
        
        # --- NOVA LÓGICA DE DATA: -3 dias úteis ignorando feriados ---
        dt = dados["data_emissao"]
        dias_sub = 0
        while dias_sub < 3:
            dt -= timedelta(days=1)
            # Só contabiliza se for dia útil e não for feriado
            if is_dia_util(dt):
                dias_sub += 1
                
        data_fmt = f"{dt.strftime('%d')} DE {MESES[dt.month]} DE {dt.year}"
        # -------------------------------------------------------------

        pasta_nf = saida / f"NF {dados['numero_nf']}"
        pasta_nf.mkdir(exist_ok=True)

        if tpl_lc.exists():
            itens_lc, tot_lc = reajustar_valores(dados["itens"], 0.03)
            preencher_word(tpl_lc, pasta_nf / f"LC COMERCIAL NF {dados['numero_nf']}.docx", dados, itens_lc, tot_lc, data_fmt)
            print(f"  ✓ Proposta LC criada (+3.0%) - Data orçada: {dt.strftime('%d/%m/%Y')}")

        if tpl_jw.exists():
            margem_jw = random.uniform(0.05, 0.08)
            itens_jw, tot_jw = reajustar_valores(dados["itens"], margem_jw)
            preencher_word(tpl_jw, pasta_nf / f"JW COMERCIAL NF {dados['numero_nf']}.docx", dados, itens_jw, tot_jw, data_fmt, is_jw=True)
            print(f"  ✓ Proposta JW criada (+{margem_jw*100:.2f}%) - Data orçada: {dt.strftime('%d/%m/%Y')}")

if __name__ == "__main__":
    processar_notas()