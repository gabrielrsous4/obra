# Plano de Ação — Sistema Web

Sistema web (HTML/CSS/JS puro, sem backend) que reproduz fielmente a planilha
**PLANO DE AÇÃO (Gestão de Desempenho)**. O usuário preenche as ações pela tela
e, ao final, o próprio sistema gera:

- uma **planilha Excel (.xlsx)** idêntica ao modelo original (cores, logo,
  caixa de código, menus suspensos, fórmula de status e formatação
  condicional);
- um **PDF** com o mesmo layout, pronto para impressão/compartilhamento.

Os dados digitados ficam salvos automaticamente no navegador (`localStorage`),
então é possível fechar a aba e continuar depois sem perder o preenchimento.

## Como usar

Basta abrir o `index.html` em um navegador (funciona offline, não precisa de
servidor). Para publicar no GitHub Pages, é só subir a pasta como está e
apontar o Pages para a raiz do repositório.

Para testar localmente com um servidor simples:

```bash
python3 -m http.server 8000
# depois acesse http://localhost:8000
```

## Estrutura do projeto

```
index.html      Layout da tela (cabeçalho, formulário e tabela)
styles.css       Estilo visual (cores e proporções iguais à planilha original)
listas.js        Diretorias e departamentos (extraídos da aba "Listas" da planilha original)
app.js           Estado da aplicação, cálculo automático de status, persistência local
export.js        Geração do Excel (ExcelJS) e do PDF (jsPDF + AutoTable)
assets/logo.png  Logo extraída da planilha original
vendor/          Bibliotecas de terceiros (ExcelJS e jsPDF), incluídas localmente
                 para o site funcionar 100% offline/estático, sem CDN
```

## Fidelidade ao modelo original

- Cabeçalho com logo, título "PLANO DE AÇÃO (Gestão de Desempenho)" e caixa
  CÓDIGO/FORM./VERSÃO, iguais à planilha.
- Barra REUNIÃO / DIRETORIA / DEPARTAMENTO com as mesmas cores
  (`#0C5EA5` e `#07B1E5`) e as mesmas opções da aba "Listas" (departamento
  varia de acordo com a diretoria escolhida, tanto na tela quanto no Excel
  gerado, via lista suspensa em cascata).
- Tabela com as 10 colunas originais: Data Registro, Área/Obra, Assunto,
  Detalhamento Ação, Observações, Responsáveis, Data Prevista Término, Data
  Reprogramada, Data Conclusão e Status.
- **Status calculado automaticamente**, replicando exatamente a fórmula da
  planilha original:
  `=IF(ISBLANK(Detalhamento),"",IF(ISBLANK(DataConclusão),IF(DataPrevista-HOJE()>=0,"Em Andamento","Atrasado"),"Concluído"))`
- Cores de status iguais ao original: amarelo (Em Andamento), vermelho
  (Atrasado) e verde (Concluído) — tanto na tela quanto no Excel (formatação
  condicional real) e no PDF.
- O Excel gerado contém fórmulas vivas (não valores fixos), validações de
  lista (menus suspensos) e a aba "Listas" de apoio, podendo ser editado
  depois no Excel normalmente.

## Observação sobre o arquivo original

Na planilha original, o menu suspenso de "Departamento" estava quebrado
(referência `#REF!`). Neste sistema esse comportamento foi corrigido tanto na
tela quanto no Excel gerado (usando nomes definidos + `INDIRECT`), para que a
lista de departamentos realmente dependa da diretoria escolhida.
