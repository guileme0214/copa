# Bolão Copa do Mundo 2026 ⚽

Planilha para o bolão da fase de grupos da Copa do Mundo 2026 (EUA, Canadá e México).

## Arquivos
- **`bolao_copa_2026.xlsx`** — a planilha pronta para usar (é o que você sobe no Google Sheets).
- `gerar_bolao.py` — script que gera a planilha (caso queira regerar ou editar).

## Como funciona
São os **72 jogos da fase de grupos** em ordem de data, um por linha. Para cada jogo:

| Coluna | O que é |
|--------|---------|
| Data | data do jogo |
| Grupo | grupo (A a L) |
| Time 1 | seleção da esquerda |
| Bandeira | bandeira do Time 1 |
| Placar Real T1 | resultado de verdade do Time 1 *(deixe em branco até o jogo acontecer)* |
| Palpite T1 | quantos gols você acha que o Time 1 faz |
| Palpite T2 | quantos gols você acha que o Time 2 faz |
| Placar Real T2 | resultado de verdade do Time 2 *(deixe em branco até o jogo acontecer)* |
| Bandeira | bandeira do Time 2 |
| Time 2 | seleção da direita |

> As colunas de **Data** e **Grupo** são extras só para ajudar a se organizar — pode apagar se quiser.

## Como subir no Google Sheets
1. Acesse [sheets.google.com](https://sheets.google.com) e crie uma planilha em branco (ou abra o Drive).
2. **Arquivo → Importar → Fazer upload** e selecione `bolao_copa_2026.xlsx`.
3. Em "Local de importação", escolha *Criar nova planilha* (ou substituir) e confirme.

Pronto — é só preencher os palpites. Para um bolão com várias pessoas, duplique a aba "Fase de Grupos" (botão direito na aba → *Duplicar*) e renomeie com o nome de cada participante.

## Observação sobre as bandeiras
As bandeiras são **emoji** (🇧🇷, 🇦🇷, ...), que aparecem na maioria dos aparelhos. As de **Escócia** (🏴󠁧󠁢󠁳󠁣󠁴󠁿) e **Inglaterra** (🏴󠁧󠁢󠁥󠁮󠁧󠁿) usam um emoji especial que pode não renderizar em alguns sistemas mais antigos — nesse caso aparece um quadrado/preto, mas o nome do país ao lado continua correto.

## Regerar a planilha
```bash
pip install openpyxl
python3 gerar_bolao.py
```
Os jogos ficam na lista `JOGOS` dentro de `gerar_bolao.py` (fonte: sorteio de 05/12/2025, confirmado por ESPN/Wikipedia). Os nomes dos times estão em português e podem ser ajustados ali (ex.: trocar "Países Baixos" por "Holanda").
