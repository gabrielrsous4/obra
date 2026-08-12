// Dados extraídos da aba "Listas" da planilha original.
// DIRETORIAS: nome completo -> abreviação (usada para localizar a coluna de departamentos)
const DIRETORIAS = [
  { nome: "Gente e Gestão", abrev: "GG" },
  { nome: "Engenharia", abrev: "Engenharia" },
  { nome: "Desenvolvimento Imobiliário", abrev: "D.I" },
  { nome: "Comercial", abrev: "Comercial" },
  { nome: "Controladoria", abrev: "Controladoria" },
  { nome: "Jurídico e Compliance", abrev: "J.C" },
];

// DEPARTAMENTOS por abreviação de diretoria (colunas F..K da aba Listas)
const DEPARTAMENTOS_POR_DIRETORIA = {
  "GG": ["CEO","Diretoria","Gerência","Business Partner","Desenvolvimento","Gestão","Inovação","Operações G&G","Saúde","Segurança do Trabalho","Tecnologia da Informação"],
  "Engenharia": ["CEO","Diretoria","Gerência","Esteira de Negócios","Obras","Orçamentos","Planejamento de Obras","Qualidade","Suporte Técnico","Suprimentos","Excelência Operacional"],
  "D.I": ["CEO","Diretoria","Gerência","Incorporação","Novos Negócios","Projetos de Arquitetura"],
  "Comercial": ["CEO","Diretoria","Gerência","Inteligência de Mercado","Marketing","Operações Comerciais","Relacionamento com o Cliente"],
  "Controladoria": ["CEO","Diretoria","Gerência","Contabilidade","Controladoria","Crédito","Finanças","Fiscal","Planejamento Estratégico","Tesouraria","Tributário"],
  "J.C": ["CEO","Diretoria","Gerência","Jurídico","Controles Internos e Compliance"],
};
