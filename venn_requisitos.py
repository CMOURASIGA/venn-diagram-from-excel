import pandas as pd
import matplotlib.pyplot as plt
from venn import venn
from itertools import chain, combinations
from tkinter import Tk, filedialog, simpledialog

# ======= ENTRADA INTERATIVA =======
root = Tk()
root.withdraw()
CAMINHO_ARQUIVO = filedialog.askopenfilename(title="Selecione o arquivo matriz.xlsx", filetypes=[("Excel files", "*.xlsx")])

if not CAMINHO_ARQUIVO:
    raise Exception("Nenhum arquivo selecionado. O script será encerrado.")

df = pd.read_excel(CAMINHO_ARQUIVO)
colunas = df.columns.tolist()
ID_COLUNA = colunas[0]

# Escolha dinâmica dos sistemas
sistemas_disponiveis = colunas[2:]
texto_opcoes = "\n".join(f"{i+1}: {s}" for i, s in enumerate(sistemas_disponiveis))
entrada = simpledialog.askstring("Seleção de Sistemas", f"Digite os números dos sistemas separados por vírgula:\n{texto_opcoes}")

if not entrada:
    raise Exception("Nenhuma seleção feita. O script será encerrado.")

indices = [int(i.strip())-1 for i in entrada.split(',') if i.strip().isdigit() and 0 <= int(i.strip())-1 < len(sistemas_disponiveis)]
SISTEMAS = [sistemas_disponiveis[i] for i in indices]

if len(SISTEMAS) < 2:
    raise Exception("Selecione pelo menos dois sistemas para análise de interseção.")

# ======= PROCESSAMENTO =======
df_filtrado = df[df[SISTEMAS].sum(axis=1) > 0].copy()

# Criar conjunto de IDs por sistema individual
data = {
    sistema: set(df_filtrado[df_filtrado[sistema] == 1][ID_COLUNA])
    for sistema in SISTEMAS
}

# Gerar todas as combinações possíveis de sistemas para interseção exata
def powerset(iterable):
    s = list(iterable)
    return chain.from_iterable(combinations(s, r) for r in range(1, len(s)+1))

intersecoes_map = {}
for combo in powerset(SISTEMAS):
    inclusos = set.intersection(*(data[s] for s in combo))
    excluidos = set.union(*(data[s] for s in SISTEMAS if s not in combo)) if len(combo) < len(SISTEMAS) else set()
    somente_essa_intersecao = inclusos - excluidos
    if somente_essa_intersecao:
        chave = '&'.join(sorted(combo))
        intersecoes_map[chave] = list(somente_essa_intersecao)

# DEBUG: Imprimir todas as interseções e requisitos encontrados
print("\n===== INTERSEÇÕES ENCONTRADAS =====")
for chave, ids in sorted(intersecoes_map.items(), key=lambda x: (-len(x[1]), x[0])):
    print(f"{chave} ({len(ids)}): {ids}")

# ======= GRÁFICO COM VENN SOMENTE IDs =======
fig, ax = plt.subplots(figsize=(12, 10))
venn(data, ax=ax, legend_loc="upper left")

# Ocultar valores zerados e exibir apenas os IDs
for text in ax.texts:
    content = text.get_text()
    if content.isdigit():
        if int(content) == 0:
            text.set_visible(False)
        else:
            x, y = text.get_position()
            for key, ids in intersecoes_map.items():
                if len(ids) == int(content):
                    text.set_text("\n".join(ids))
                    text.set_fontsize(8)
                    break

plt.title("Interseções de Requisitos por Sistema\nAnálise de Cobertura Funcional", fontsize=13, weight='bold')
plt.tight_layout()
plt.show()