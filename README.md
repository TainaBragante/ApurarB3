# CALCULO DOS RESULTADOS DE AÇÕES - APURAÇÃO MENSAL B3.

## RESUMO:
- Lê notas de corretagem em PDF de uma pasta específica (uma pasta por mês do ano, ex.: .../2025/01.25 ou .../01.25).
- Extrai as operações de compra e venda, taxas e IRRF de cada nota usando uma abordagem em camadas de parsers.
- Classifica as operações em Day Trade, Swing Trade e FIIs.
- Detecta o tipo do ativo (ON, PN, FII, BDR, UNITS e ETF) pela descrição; ou pergunta ao usuário quando não identificado automaticamente.
- Separa o resultado das operações de swing por tipo para aplicar a regra de isenção de 20k (apenas ON e PN são elegíveis à isenção).
- O resultado das operações de day trade é calculado separadamente, sem isenção independentemente do tipo de ativo.
- Calcula os resultados líquidos por tipo de operação, considerando taxas e IRRF.
- Atualiza a planilha Excel padrão (APURAÇÃO B3 - AAAA.xlsx) com o resultado da apuração mensal.
- Monta a carteira final do mês e grava em uma nova aba "CARTEIRA_MM.AA" na planilha Excel padrão.
- Permite exportar um Excel de debug (memória de cálculo) com todas as operações lidas.

## REGRAS DE TIPO DE ATIVO E ISENÇÃO:
    ON  (Ordinária)    → TEM isenção 20k  → células B15 + G20 (se ≥20k) e células J20 e G20 (se <20k)
    PN  (Preferencial) → TEM isenção 20k  → células B15 + G20 (se ≥20k) e células J20 e G20 (se <20k)
    FII / FIAGRO       → NÃO tem isenção  → vai para B49 diretamente 
    BDR / DRN          → NÃO tem isenção  → vai para B15 diretamente
    UNITS / UNT / UN   → NÃO tem isenção  → vai para B15 diretamente
    ETF                → NÃO tem isenção  → vai para B15 diretamente

## REGRAS DE CÁLCULO DE RESULTADO E APURAÇÃO:
    - Soma as VENDAS de ON+PN do mês:
   - Se total de vendas ON+PN < R$20.000 → lucro vai para J20 (ganho isento).
   - Se total de vendas ON+PN ≥ R$20.000 → lucro vai para B15 (tributável a 15%).
    - Independente da regra acima - BDR, UNITS e ETF sempre vão para B15 sem isenção (tributável a 15%).
    - Independente da regra acima - FIIs sempre vão para B49 com sem isenção (tributável a 20%).
    - O resultado final da apuração mensal é a soma de B15 (lucro tributável -> ON+PN+BDR+UNITS+ETF) + B49 (lucro FII)
    - Em G20 ficam o total de alienações dentro do mês ON+PN+BDR+UNITS+ETF.

## REGRAS DE PARSER EM CAMADAS:
    1ª Camada → CorrePy parser padrão, tenta identificar o padrão SINACOR com CorrePy
    2ª Camada → Parser próprio específico para cada corretora com pdfplumber; para notas que não seguem o layout CorrePy
    3ª Camada → Parser genérico com pdfplumber, tenta extrair operações de notas não-CorrePy sem parser específico, usando heurísticas de texto
    4ª Camada → Parser manual assistido, para casos onde o parser genérico não consegue extrair as operações corretamente. O usuário deve digitar os dados em uma grade visual e o sistema valida a consistência antes de aceitar.