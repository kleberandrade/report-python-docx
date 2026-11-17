# Revisão de material didático

Foram inspecionados o README e os dois scripts originais, incluindo os links HTTP/HTTPS. Não há URL de artigo do Medium nem encurtador nesses arquivos. O único link didático externo identificado é o da instituição Fatec; os outros links apontam ao próprio repositório e a imagens dele.

Não foi identificado artigo correspondente cuja autoria e identidade pudessem ser verificadas. Portanto, nenhum artigo foi aberto em editor, alterado ou publicado. Não houve envio de email, newsletter ou mensagem. Não se afirma bloqueio editorial ou paywall: não houve artigo vinculado para testar esse acesso.

Se houver um artigo correspondente fora desses arquivos, a revisão concreta a aplicar é:

- Preservar português, URL, autoria e propósito da aula.
- Substituir instalação sem versão por Python 3.11+ e os comandos do README atual; usar python-docx 1.2.0 e openpyxl 3.1.5.
- Trocar a execução dos scripts com efeitos colaterais por `class-report examples/alpha.xlsx -o reports/alpha` e o equivalente beta com `--stage beta`.
- Corrigir a explicação de interesse em jogar: ler `Você jogaria este jogo?`, independentemente do preço.
- Explicar média e desvio padrão populacional, tabelas de frequência e comentários completos; não apresentar nuvens de palavras como saída atual.
- Explicar validação de cabeçalhos, notas numéricas de 0 a 10, dificuldade beta de 1 a 5, comentários ausentes e arquivos de saída estáveis.
- Remover credenciais e instruções de envio de email ou limpeza por curingas. Mostrar apenas exemplos sintéticos.
- Atualizar imagens com relatórios atuais após conferir a renderização; não chamar exemplos sintéticos de resultados reais da turma.
