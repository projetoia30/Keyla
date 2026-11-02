from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

d = Document()
for s in d.sections:
    s.top_margin = s.left_margin = Cm(3); s.bottom_margin = s.right_margin = Cm(2)
st = d.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
st.paragraph_format.line_spacing = 1.5; st.paragraph_format.space_after = Pt(0)

def P(t, center=False, bold=False, indent=True, size=None, ls=None, right=False, after=0):
    p = d.add_paragraph()
    # suporte a **negrito** simples
    parts = t.split('**')
    for i, part in enumerate(parts):
        r = p.add_run(part); r.bold = bold or (i % 2 == 1)
        if size: r.font.size = Pt(size)
    if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif right: p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if indent: p.paragraph_format.first_line_indent = Cm(1.25)
    if ls: p.paragraph_format.line_spacing = ls
    p.paragraph_format.space_after = Pt(after)
    return p

def H(t, lvl=1):
    p = d.add_paragraph(); r = p.add_run(t); r.bold = True
    p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.keep_with_next = True

def Q(t):
    p = d.add_paragraph(); r = p.add_run(t); r.font.size = Pt(10)
    p.paragraph_format.left_indent = Cm(4); p.paragraph_format.line_spacing = 1.0
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(12); p.paragraph_format.space_before = Pt(12)

def A(t):  # alínea
    p = P(t, indent=False); p.paragraph_format.left_indent = Cm(1.25)

P('A REMIÇÃO DE PENA PELO TRABALHO E PELO ESTUDO NO CONJUNTO PENAL DE FEIRA DE SANTANA-BA', center=True, bold=True, after=12)
P('Nilton Dias Oliveira¹', right=True, ls=1.0)
P('Alberto Azis²', right=True, ls=1.0)
P('Daiane Zappe Viana Veronese³', right=True, ls=1.0, after=12)
for n in ['¹ Graduando em Direito pela Unidade de Ensino Superior de Feira de Santana (UNEF). E-mail: [e-mail do autor].',
          '² [Titulação], Unidade de Ensino Superior de Feira de Santana (UNEF). E-mail: [e-mail do professor Alberto Azis].',
          '³ Mestre em Ciências Jurídico-Criminais pela Universidade de Coimbra, Unidade de Ensino Superior de Feira de Santana (UNEF). E-mail: advogadadaiane@hotmail.com.']:
    P(n, indent=False, size=10, ls=1.0)
P('[Atenção: conferir com os orientadores a titulação de cada um e a qual deles pertence a formação em Coimbra; no original, as notas estavam trocadas.]', indent=False, size=10, ls=1.0, after=12)

H('1 INTRODUÇÃO')
H('1.1 TEMA')
P('A remição de pena pelo trabalho e pelo estudo no Conjunto Penal de Feira de Santana-BA.')

H('1.2 PROBLEMA DE PESQUISA')
P('Quais fatores institucionais e burocráticos limitam o acesso das pessoas privadas de liberdade do Conjunto Penal de Feira de Santana à remição de pena pelo trabalho e pelo estudo?')

H('1.3 HIPÓTESE')
P('Supõe-se que, embora o ordenamento jurídico assegure a remição de pena pelo trabalho e pelo estudo de forma ampla, o acesso efetivo a esse direito no Conjunto Penal de Feira de Santana seja limitado por dois conjuntos de fatores: (a) estruturais, decorrentes da desproporção entre o número de custodiados e a quantidade de vagas em atividades laborais e educacionais; e (b) burocráticos, relacionados ao fluxo de registro de frequência, emissão de atestados, remessa ao juízo da execução e homologação judicial dos dias remidos.')

H('1.4 RELEVÂNCIA DO ESTUDO')
P('A relevância deste estudo baseia-se na necessidade de analisar a aplicabilidade prática da remição de pena pelo trabalho e pelo estudo, prevista no art. 126 da Lei de Execução Penal (LEP), na realidade do Conjunto Penal de Feira de Santana (CPFS). Em um sistema carcerário tensionado pela superlotação e pela morosidade processual, investigar os gargalos locais que dificultam a oferta da remição torna-se indispensável para verificar se o instituto cumpre sua função de integração social ou se apenas mantém indivíduos segregados, sem preparo para o retorno à vida em comunidade.')
P('Nessa ótica, a pesquisa destaca-se ao buscar caminhos que confiram maior efetividade à remição, permitindo que a pessoa privada de liberdade reduza o tempo de cárcere no limite legalmente admissível. Além do aspecto aritmético de liquidação da pena, a rotina produtiva e formativa estabelece um vínculo com o meio externo: ao desenvolver habilidades técnicas, disciplina laboral e senso de dignidade pessoal, contribui para romper o ciclo de isolamento e exclusão, favorecendo que o egresso restabeleça vínculos familiares e tenha acesso ao mercado formal de trabalho ao retornar ao convívio social.')
P('Por fim, o valor acadêmico e social desta pesquisa reside em diagnosticar os desafios estruturais e procedimentais do Conjunto Penal de Feira de Santana sob a ótica dos direitos humanos e da prevenção da reincidência, analisando em que medida a execução penal tem cumprido sua função preventiva e integradora. Acrescenta-se a escassez de estudos acadêmicos voltados especificamente à realidade prisional de Feira de Santana, o que confere originalidade à investigação.')

H('1.5 OBJETIVO GERAL')
P('Analisar os fatores institucionais e burocráticos que limitam o acesso das pessoas privadas de liberdade do Conjunto Penal de Feira de Santana à remição de pena pelo trabalho e pelo estudo.')

H('1.6 OBJETIVOS ESPECÍFICOS')
A('a) Compreender o funcionamento da remição de pena pelo trabalho e pelo estudo nos diferentes regimes de cumprimento de pena, destacando suas diferenças e particularidades;')
A('b) Analisar a possibilidade de cumulação da remição pelo trabalho e pelo estudo, à luz do art. 126, § 3º, da LEP, e sua importância para a formação e a ressocialização da pessoa privada de liberdade;')
A('c) Mapear as frentes de trabalho e as atividades educacionais oferecidas aos internos do Conjunto Penal de Feira de Santana, identificando as oportunidades disponíveis para a remição;')
A('d) Descrever o fluxo administrativo e judicial de reconhecimento da remição no Conjunto Penal de Feira de Santana, desde o registro de frequência até a homologação pelo juízo da execução;')
A('e) Identificar os principais obstáculos enfrentados pelas pessoas privadas de liberdade para acessar as atividades de trabalho e estudo que possibilitam a remição da pena.')

H('1.7 CONSIDERAÇÕES METODOLÓGICAS')
P('Quanto à natureza, esta pesquisa é aplicada, com abordagem predominantemente qualitativa e finalidades exploratória e descritiva. O trabalho tem como objetivo analisar a remição de pena pelo trabalho e pelo estudo, conectando os fundamentos teóricos e normativos da execução penal à realidade do Conjunto Penal de Feira de Santana (CPFS). O método de abordagem é o dedutivo.')
P('Os procedimentos técnicos serão desenvolvidos a partir de dois eixos principais: pesquisa bibliográfica e documental, de cunho teórico-dogmático, e estudo de caso, voltado à realidade prática da unidade prisional.')
P('A pesquisa bibliográfica envolverá a análise de livros, artigos científicos, dissertações, teses e outras produções acadêmicas relacionadas ao Direito Penal, ao Direito Processual Penal e, sobretudo, à Execução Penal. Serão examinadas diferentes perspectivas doutrinárias sobre as finalidades da pena, a ressocialização, os direitos fundamentais das pessoas privadas de liberdade e a remição de pena pelo trabalho e pelo estudo.')
P('No campo documental e normativo, serão analisados a Constituição da República Federativa do Brasil de 1988, a Lei nº 7.210/1984 (Lei de Execução Penal), especialmente os artigos 126 a 130, as Leis nº 12.433/2011 e nº 13.163/2015, as resoluções e recomendações do Conselho Nacional de Justiça (CNJ), os atos normativos da Secretaria de Administração Penitenciária e Ressocialização do Estado da Bahia (SEAP-BA), bem como súmulas e precedentes do Superior Tribunal de Justiça (STJ) e do Supremo Tribunal Federal (STF). Esse passo tem por finalidade construir o referencial teórico e jurídico que permita compreender o conceito de remição e os mecanismos que viabilizam esse direito no sistema prisional.')
P('O estudo de caso terá como unidade de análise o Conjunto Penal de Feira de Santana e utilizará as seguintes fontes de dados: (a) dados oficiais do Sistema de Informações do Departamento Penitenciário Nacional (SISDEPEN/SENAPPEN) referentes à unidade e ao Estado da Bahia; (b) informações institucionais solicitadas à SEAP-BA e à direção da unidade, inclusive mediante pedido fundamentado na Lei nº 12.527/2011 (Lei de Acesso à Informação), sobre número de custodiados, vagas de trabalho e de educação, convênios e fluxo de remessa dos atestados; (c) decisões da Vara de Execuções Penais da Comarca de Feira de Santana e do Tribunal de Justiça da Bahia sobre remição, consultadas em bases públicas; e (d) observação sistemática do fluxo administrativo, registrada em diário de campo com roteiro previamente definido. O recorte temporal compreenderá o período de [definir, ex.: 2022 a 2026].')
P('A pesquisa contará com a experiência profissional do pesquisador, que atua como policial penal na unidade, o que proporciona acesso contextualizado à realidade institucional. A observação permitirá identificar aspectos dos procedimentos administrativos necessários à concessão e à homologação da remição, como o controle de frequência nas atividades laborais e educacionais, a elaboração e o envio da documentação ao juízo da execução e eventuais dificuldades decorrentes de fatores estruturais, administrativos ou de segurança.')
P('Reconhece-se que a dupla condição de servidor e pesquisador pode gerar vieses de interpretação. Para mitigá-los, a observação seguirá roteiro estruturado, os registros serão confrontados com os dados documentais e oficiais (triangulação de fontes) e não serão utilizadas informações protegidas por sigilo funcional nem dados que permitam identificar pessoas privadas de liberdade ou servidores.')
H('1.7.1 Aspectos éticos')
P('Por envolver ambiente prisional e possível interação com seres humanos, a etapa empírica será realizada somente após: (a) autorização formal da SEAP-BA e da direção do Conjunto Penal de Feira de Santana; e (b) apreciação e aprovação do projeto por Comitê de Ética em Pesquisa, nos termos da Resolução CNS nº 510/2016. Serão observados o sigilo das informações, o anonimato dos sujeitos, as normas institucionais vigentes e os direitos fundamentais das pessoas privadas de liberdade. Caso as autorizações não sejam obtidas no prazo do cronograma, o estudo de caso será conduzido exclusivamente a partir de fontes documentais públicas e de dados obtidos via Lei de Acesso à Informação.')

H('2 REFERENCIAL TEÓRICO: EXECUÇÃO PENAL E REMIÇÃO DE PENA')
H('2.1 FUNDAMENTOS CONSTITUCIONAIS E A JURISDICIONALIZAÇÃO DA EXECUÇÃO PENAL')
P('A execução penal brasileira passou por profunda reestruturação conceitual a partir da promulgação da Lei de Execução Penal (Lei nº 7.210/1984) e da Constituição Federal de 1988. O cumprimento da sanção penal deixou de ser visto como atividade meramente administrativa, pautada pela discricionariedade estatal, para assumir nítido caráter jurisdicionalizado, convertendo o condenado em efetivo sujeito de direitos no curso do processo executório.')
P('A Constituição assegura expressamente os princípios da dignidade da pessoa humana (art. 1º, III), da individualização da pena (art. 5º, XLVI) e do devido processo legal (art. 5º, LIV) como pilares indispensáveis da fase executiva. Ao analisar a consolidação da jurisdicionalização no processo penal executório, Sumida (2001, p. 124) destaca:')
Q('A "jurisdicionalização" da execução penal, em consonância com o escopo socializador que o legislador conferiu à sanção penal, atribuiu ao juiz da execução o poder de alterar, dentro dos limites da lei, o título executivo. Admitindo-se, portanto, um processo penal em duas fases [...]: de conhecimento [...] e a de execução, quando se efetiva a sanção e busca a sua adequação aos resultados do tratamento penitenciário.')
P('Nessa mesma perspectiva de dinamismo da sanção no curso de seu cumprimento, Nucci (2008, p. 399) leciona:')
Q('A sentença condenatória não é estática, mas dinâmica. Um título executivo judicial, na órbita penal, é mutável. Um réu condenado ao cumprimento da pena de reclusão de doze anos, em regime inicial fechado, pode cumpri-la exatamente em doze anos, no regime fechado [...] ou cumpri-la em menor tempo, valendo-se de benefícios específicos (remissão [sic], comutação, progressão de regime, livramento condicional etc.).')
P('O art. 1º da LEP estabelece a dupla finalidade da execução penal: efetivar as disposições da sentença ou decisão criminal e proporcionar condições para a harmônica integração social do condenado. Contudo, a doutrina aponta grave descompasso entre a norma e a realidade do sistema carcerário brasileiro. Tratando da crise do sistema penitenciário, Dotti (1998, p. 56 apud SILVA, 2015, p. 7) adverte:')
Q('Essa disfuncionalidade dos sistemas parciais, que levou à crise da execução penal, demonstrou a necessidade de uma política geral de governo e a intervenção efetiva da comunidade para reduzir os índices alarmantes da criminalidade violenta.')
P('Complementando essa abordagem crítica, Silva (2015) ressalta que a inaplicabilidade fática das garantias legais transforma o cárcere em ambiente de degradação humana, o que reforça a necessidade de efetivar institutos ressocializadores, a exemplo da remição da pena. Esse quadro foi reconhecido pelo Supremo Tribunal Federal no julgamento da Arguição de Descumprimento de Preceito Fundamental nº 347, que declarou o estado de coisas inconstitucional do sistema prisional brasileiro (BRASIL, 2023).')
P('Sob a ótica da criminologia crítica e do garantismo penal, Zaffaroni e Roig sustentam que o direito penitenciário deve atuar como instrumento de contenção do poder punitivo do Estado e de preservação dos direitos fundamentais da pessoa presa (ZAFFARONI et al., 2003; ROIG, 2016). [Inserir as páginas consultadas.] Diante do estado de coisas inconstitucional, Roig (2016, p. 242) defende interpretação ampliativa favorável ao custodiado:')
Q('[...] considerando que o trabalho e a educação são direitos sociais (art. 6º da CF) e que a remição é instituto concebido para o benefício das pessoas presas, sua interpretação, aplicação e extensão devem ser as mais amplas possíveis, inclusive com a admissão da analogia in bonam partem.')

H('2.2 O INSTITUTO DA REMIÇÃO DE PENA: CONCEITO, BALIZAS LEGAIS E REQUISITOS DE CÔMPUTO')
P('Prevista nos artigos 126 a 130 da Lei de Execução Penal, com a redação dada pela Lei nº 12.433/2011, a remição define-se como o direito do condenado de abreviar o tempo de cumprimento da pena mediante o trabalho ou o estudo. A leitura, embora não prevista expressamente na LEP, foi incorporada por interpretação extensiva do conceito de estudo, conforme se verá no item 2.3. Quanto à etimologia e à distinção conceitual do instituto, Marcão (2021, p. 94) assinala:')
Q('[...] a palavra remição vem de redimere que no latim significa reparar, compensar, ressarcir. É preciso não confundir remição com remissão; esta, segundo o léxico, significa a ação de perdoar.')
P('Em definição clássica quanto à finalidade do instituto, Mirabete (2004, p. 517) conceitua a remição nos seguintes termos:')
Q('Pode-se definir a remição, nos termos da lei brasileira, como um direito do condenado em reduzir pelo trabalho prisional o tempo de duração da pena privativa de liberdade cumprida em regime fechado ou semiaberto. Trata-se de um meio de abreviar ou extinguir parte da pena. Oferece-se ao preso um estímulo para corrigir-se, abreviando o tempo de cumprimento da sanção para que possa passar ao regime de liberdade condicional ou à liberdade definitiva.')
P('Registre-se que essa definição, anterior à Lei nº 12.433/2011, refere-se apenas ao trabalho, pois a remição pelo estudo só foi positivada em 2011, após o seu reconhecimento jurisprudencial pela Súmula 341 do STJ. Reforçando a dimensão humanitária do benefício, Dias (apud MIRABETE, 2004) assinala que a remição se revela um instrumento completo, [conferir a página na edição de 2004 ou manter a edição de 2000 nas duas citações]')
Q('[...] pois reeduca o delinquente, prepara-o para sua reincorporação à sociedade, proporciona-lhe meios para reabilitar-se diante de si mesmo e diante da sociedade, disciplina sua vontade, favorece sua família e sobretudo abrevia a condenação, condicionando esta ao próprio esforço do apenado.')
P('Sob o aspecto de sua natureza jurídica e de seu papel na individualização da pena, Leal (1993, p. 459) leciona:')
Q('Ao concedê-la, o Estado leva em consideração o resultado positivo do mérito demonstrado pelo condenado em aceitar o trabalho prisional e observar as demais regras de disciplina prisional. Em consequência, e tendo em vista o princípio constitucional da individualização da pena, considera de boa política criminal abdicar de parte da sanção criminal imposta na sentença criminal. Assim sendo, a remição é um instituto de natureza penal que opera como uma causa extintiva da punibilidade [...]. [Conferir a transcrição no original.]')
P('A qualificação da remição como causa extintiva da punibilidade, contudo, não é pacífica. Parte da doutrina entende que o tempo remido é computado como pena efetivamente cumprida, para todos os efeitos (art. 128 da LEP), e não como causa de extinção da punibilidade nos moldes do art. 107 do Código Penal. O cômputo do benefício opera-se segundo as proporções fixadas no art. 126 da LEP.')
P('**Remição pelo trabalho.** Aplica-se aos regimes fechado e semiaberto, à razão de 1 (um) dia de pena a cada 3 (três) dias de trabalho (art. 126, § 1º, II), observada a jornada de seis a oito horas diárias (art. 33). Nos termos da Súmula 562 do STJ, "é possível a remição de parte do tempo de execução da pena quando o condenado, em regime fechado ou semiaberto, desempenha atividade laborativa, ainda que extramuros". O STJ não admite a remição pelo trabalho no regime aberto, pois nele o trabalho é requisito do próprio regime (art. 114, I, da LEP). Quanto à remuneração, o STF, no julgamento da ADPF 336, reconheceu a constitucionalidade do art. 29 da LEP, que fixa remuneração não inferior a três quartos do salário mínimo.')
P('**Remição pelo estudo.** Garante o abatimento de 1 (um) dia de pena a cada 12 (doze) horas de frequência escolar em atividade de ensino fundamental, médio, inclusive profissionalizante, ou superior, ou ainda de requalificação profissional, divididas em, no mínimo, 3 (três) dias (art. 126, § 1º, I). As atividades podem ser presenciais ou a distância, desde que certificadas pelas autoridades educacionais competentes (§ 2º). A conclusão do ensino fundamental, médio ou superior durante o cumprimento da pena acrescenta 1/3 (um terço) ao tempo remido pelo estudo (§ 5º). Diferentemente do trabalho, a remição pelo estudo estende-se ao regime aberto, ao semiaberto e ao livramento condicional (§ 6º).')
P('**Cumulação.** As horas diárias de trabalho e de estudo podem ser cumuladas para fins de remição, desde que definidas de forma a se compatibilizarem (art. 126, § 3º). Também o preso impossibilitado, por acidente, de prosseguir no trabalho ou nos estudos continuará a beneficiar-se com a remição (§ 4º).')
P('**Perda dos dias remidos.** Em caso de falta grave, o juiz poderá revogar até 1/3 (um terço) do tempo remido, observado o art. 57 da LEP, recomeçando a contagem a partir da data da infração disciplinar (art. 127). A regra, introduzida pela Lei nº 12.433/2011, substituiu a perda integral prevista na redação original e exige decisão fundamentada quanto à fração aplicada.')
P('**Procedimento.** A autoridade administrativa encaminhará mensalmente ao juízo da execução cópia do registro de todos os condenados que estejam trabalhando ou estudando, com informação dos dias de trabalho ou das horas de frequência escolar (art. 129). A remição será declarada pelo juiz da execução, ouvidos o Ministério Público e a defesa (art. 126, § 8º). [Conferir a redação atual dos dispositivos.]')

H('2.3 A EXPANSÃO HERMENÊUTICA: REMIÇÃO PELA LEITURA E NORMATIVAS DO CNJ')
P('Para além das modalidades tradicionais de trabalho e estudo formal, a jurisprudência e a normatização infralegal ampliaram o instituto. No campo laboral, o STJ tem admitido o artesanato como atividade apta à remição pelo trabalho, desde que devidamente supervisionado e registrado pela administração prisional [conferir precedente]. No campo educacional, a principal ampliação foi a remição pela leitura, inicialmente orientada pela Recomendação nº 44/2013 do CNJ e consolidada pela Resolução nº 391/2021 do CNJ, que a inseriu entre as práticas sociais educativas, ao lado das atividades educacionais complementares e da aprovação em exames nacionais de certificação, como o ENCCEJA e o ENEM.')
P('Fundamentando o papel da formação intelectual no cárcere, Nucci (2021, p. 208) afirma:')
Q('[...] sendo a reeducação uma das finalidades da pena, não há dúvida que o trabalho e o estudo são fortes instrumentos para tanto, impedindo a ociosidade perniciosa no cárcere.')
P('No mesmo sentido, Marcão (2021, p. 94) enfatiza os reflexos do incentivo intelectual: [conferir a página, pois a citação anterior do mesmo autor indica a mesma página]')
Q('[...] não raras vezes o estudo acarretará melhores e mais sensíveis efeitos no presente e no futuro do sentenciado, vale dizer, durante o período de cumprimento de pena e no momento da reinserção social, do que o trabalho propriamente dito.')
P('Em termos operacionais, a Resolução nº 391/2021 do CNJ estabelece prazo para a leitura de cada obra [conferir o prazo exato no texto da resolução], seguida da apresentação de relatório de leitura. A aprovação do relatório garante a remição de 4 (quatro) dias de pena por obra, limitada a 12 (doze) obras por ano, o que permite remir até 48 (quarenta e oito) dias anualmente. A resolução prevê, ainda, adaptações para pessoas em processo de alfabetização e com deficiência.')

H('2.4 O TRABALHO PENITENCIÁRIO E A DIVERGÊNCIA SOBRE A "REMIÇÃO FICTA"')
P('O trabalho do condenado é definido pelo art. 28 da LEP como dever social e condição de dignidade humana, com finalidade educativa e produtiva, constituindo também direito do preso (art. 41, II). Enaltecendo a função integradora do trabalho carcerário, Barros (2001, p. 188) pontua que "o preso, como trabalhador, identifica-se com a sociedade. O homem livre trabalha, o preso também".')
P('Contudo, quando o Estado não fornece vagas suficientes de trabalho ou estudo, instaura-se relevante controvérsia doutrinária e jurisprudencial acerca da chamada "remição ficta" ou por ociosidade forçada, que pode ser organizada em três posições.')
P('**Regra geral nos tribunais superiores: vedação.** O STF e o STJ consolidaram o entendimento de que a efetiva dedicação a atividade produtiva ou educativa constitui requisito indispensável ao benefício (STF, HC 124.520/RO) [conferir relator e data]. A Quinta Turma do STJ assentou essa orientação:')
Q('1. O benefício da remição da pena pelo trabalho ou pelo estudo, consoante se denota do art. 126 da LEP, pressupõe que os reeducandos demonstrem a efetiva dedicação a trabalho ou estudo, com finalidade, portanto, produtiva ou educativa, dada a sua finalidade ressocializadora. 2. A suposta omissão estatal em propiciar ao apenado padrões mínimos previstos no ordenamento jurídico não pode ser utilizada como causa a ensejar a concessão ficta de um benefício que depende de um real envolvimento da pessoa do apenado em seu progresso educativo e ressocializador [...]. (BRASIL, 2017).')
P('**Exceções.** A primeira é legal: o art. 126, § 4º, da LEP assegura a continuidade do benefício ao preso impossibilitado de prosseguir no trabalho ou no estudo por acidente. A segunda é jurisprudencial: durante a pandemia de Covid-19, o STJ, em julgamento sob o rito dos recursos repetitivos (Tema 1.120, REsp 1.953.607/SC), admitiu a remição em favor dos presos que já exerciam atividades laborais ou educacionais e foram impedidos de continuá-las em razão da suspensão determinada pelo próprio Estado [conferir a tese no portal do STJ]. Essa decisão não autorizou a remição ficta de forma ampla, mas reconheceu que a interrupção involuntária, imposta pelo poder público a quem já estava inserido na atividade, não pode prejudicar o apenado.')
P('**Corrente garantista.** Defendida por autores como Rogério Greco e Rodrigo Duque Estrada Roig, sustenta que a omissão estatal não pode recair em prejuízo do custodiado [inserir obra e página de Greco]. Argumenta-se que a impossibilidade de trabalhar ou estudar por falta de vagas configura situação análoga à do acidente (art. 126, § 4º, da LEP), impondo-se a analogia in bonam partem para autorizar a compensação do tempo de ociosidade involuntária (ROIG, 2016).')

H('2.5 O CONJUNTO PENAL DE FEIRA DE SANTANA (CPFS): GARGALOS INSTITUCIONAIS E BUROCRÁTICOS')
P('A análise dogmática da remição ganha relevo empírico ao ser confrontada com a realidade do Conjunto Penal de Feira de Santana. Em um sistema penitenciário tensionado pela superlotação e pela morosidade processual, a investigação dos gargalos locais revela-se indispensável para aferir se o instituto cumpre sua função ressocializadora ou se mantém os custodiados apenas segregados, sem preparo para o retorno à vida comunitária.')
P('Com base no referencial teórico e na hipótese formulada, a pesquisa partirá da suposição de que dois entraves centrais limitam a efetividade da remição na unidade, os quais serão verificados no estudo de caso:')
A('a) **gargalo estrutural:** a provável desproporção entre o contingente carcerário e a quantidade de vagas disponíveis nas atividades laborais internas, como manutenção e serviços gerais, e no ensino formal. Essa desproporção, se confirmada, contraria o dever estatal de assegurar trabalho e assistência educacional (arts. 17 a 21, 28, 31 e 41, II, da LEP) e de manter salas de aula nos estabelecimentos penais (art. 83, § 4º), gerando filas de espera e impondo a ociosidade involuntária aos apenados;')
A('b) **gargalo burocrático:** a tramitação do fluxo documental. Embora o art. 129 da LEP determine o envio mensal dos registros de trabalho e estudo ao juízo da execução, o trâmite interno abrange o controle diário de frequência nos pavilhões, oficinas e salas de aula, a emissão dos atestados pela direção da unidade e a remessa ao juízo, com posterior manifestação do Ministério Público e da defesa e homologação judicial. Atrasos em qualquer etapa dessa cadeia podem retardar a declaração dos dias remidos e o cálculo de liquidação da pena, postergando a análise de benefícios como a progressão de regime e o livramento condicional.')
P('Para verificar essas suposições, serão levantados, sempre que possível: a capacidade e a população da unidade; o número de vagas de trabalho e de educação e a lista de espera; os convênios vigentes; o tempo médio entre o registro de frequência e a homologação judicial; e a existência de projetos de remição pela leitura. A comparação entre esses indicadores permitirá confirmar ou refutar a hipótese da pesquisa.')

H('3 CRONOGRAMA')
t = d.add_table(rows=1, cols=7); t.style = 'Table Grid'
hdr = ['Etapa', 'Mês 1', 'Mês 2', 'Mês 3', 'Mês 4', 'Mês 5', 'Mês 6']
for i, h in enumerate(hdr):
    c = t.rows[0].cells[i]; c.text = h; c.paragraphs[0].runs[0].bold = True
rows = [('Revisão bibliográfica', 'X', 'X', 'X', '', '', ''),
        ('Pré-banca e ajustes do projeto', 'X', '', '', '', '', ''),
        ('Pedido de autorização à SEAP e ao CEP', 'X', 'X', '', '', '', ''),
        ('Pedido via Lei de Acesso à Informação', '', 'X', '', '', '', ''),
        ('Coleta documental e observação', '', '', 'X', 'X', '', ''),
        ('Análise dos dados', '', '', '', 'X', 'X', ''),
        ('Redação final', '', '', '', '', 'X', 'X'),
        ('Revisão e defesa', '', '', '', '', '', 'X')]
for r in rows:
    cells = t.add_row().cells
    for i, v in enumerate(r): cells[i].text = v
for row in t.rows:
    for c in row.cells:
        for p in c.paragraphs:
            p.paragraph_format.line_spacing = 1.0
            for r_ in p.runs: r_.font.size = Pt(10)
P('Fonte: elaborado pelo autor (2026). [Ajustar os meses ao calendário da UNEF.]', indent=False, size=10, ls=1.0)

H('REFERÊNCIAS')
refs = [
'BARROS, Carmen Silvia de Moraes. A individualização da pena na execução penal. São Paulo: Revista dos Tribunais, 2001.',
'BRASIL. Constituição da República Federativa do Brasil de 1988. Brasília, DF: Presidência da República, 1988.',
'BRASIL. Lei nº 7.210, de 11 de julho de 1984. Institui a Lei de Execução Penal. Brasília, DF: Presidência da República, 1984.',
'BRASIL. Lei nº 12.433, de 29 de junho de 2011. Altera a Lei nº 7.210, de 11 de julho de 1984 (Lei de Execução Penal), para dispor sobre a remição de parte do tempo de execução da pena por estudo ou por trabalho. Brasília, DF: Presidência da República, 2011.',
'BRASIL. Lei nº 12.527, de 18 de novembro de 2011. Regula o acesso a informações. Brasília, DF: Presidência da República, 2011.',
'BRASIL. Lei nº 13.163, de 9 de setembro de 2015. Modifica a Lei nº 7.210, de 11 de julho de 1984, para instituir o ensino médio nas penitenciárias. Brasília, DF: Presidência da República, 2015.',
'BRASIL. Superior Tribunal de Justiça. Súmula nº 341. Brasília, DF: STJ, 2007.',
'BRASIL. Superior Tribunal de Justiça. Súmula nº 562. Brasília, DF: STJ, 2016.',
'BRASIL. Superior Tribunal de Justiça (5. Turma). Habeas Corpus nº 425.155/MG. Relator: Min. Ribeiro Dantas. Julgado em 23 nov. 2017. [Conferir no portal do STJ.]',
'BRASIL. Superior Tribunal de Justiça (3. Seção). Recurso Especial nº 1.953.607/SC (Tema Repetitivo 1.120). [Conferir relator e data no portal do STJ.]',
'BRASIL. Supremo Tribunal Federal. Habeas Corpus nº 124.520/RO. [Conferir relator e data no portal do STF.]',
'BRASIL. Supremo Tribunal Federal. Arguição de Descumprimento de Preceito Fundamental nº 336/DF. [Conferir relator e data no portal do STF.]',
'BRASIL. Supremo Tribunal Federal (Plenário). Arguição de Descumprimento de Preceito Fundamental nº 347/DF. Redator do acórdão: Min. Luís Roberto Barroso. Julgado em 4 out. 2023. [Conferir.]',
'CONSELHO NACIONAL DE JUSTIÇA. Recomendação nº 44, de 26 de novembro de 2013. Brasília, DF: CNJ, 2013.',
'CONSELHO NACIONAL DE JUSTIÇA. Resolução nº 391, de 10 de maio de 2021. Estabelece procedimentos e diretrizes a serem observados pelo Poder Judiciário para o reconhecimento do direito à remição de pena por meio de práticas sociais educativas em unidades de privação de liberdade. Brasília, DF: CNJ, 2021.',
'CONSELHO NACIONAL DE SAÚDE. Resolução nº 510, de 7 de abril de 2016. Brasília, DF: CNS, 2016.',
'DOTTI, René Ariel. Bases e alternativas para o sistema de penas. 2. ed. São Paulo: Revista dos Tribunais, 1998. [Conferir a obra citada por Silva.]',
'GRECO, Rogério. [Título da obra consultada]. [Edição]. [Local]: [Editora], [ano].',
'LEAL, João José. [Título da obra consultada]. [Local]: [Editora], 1993.',
'MARCÃO, Renato. Curso de execução penal. 19. ed. São Paulo: Saraiva, 2021. [Conferir edição.]',
'MIRABETE, Julio Fabbrini. Execução penal: comentários à Lei nº 7.210, de 11-7-1984. 11. ed. São Paulo: Atlas, 2004.',
'NUCCI, Guilherme de Souza. Manual de processo penal e execução penal. 5. ed. São Paulo: Revista dos Tribunais, 2008. [Conferir título e edição.]',
'NUCCI, Guilherme de Souza. Curso de execução penal. 4. ed. Rio de Janeiro: Forense, 2021. [Conferir edição.]',
'ROIG, Rodrigo Duque Estrada. Execução penal: teoria crítica. 2. ed. São Paulo: Saraiva, 2016.',
'SILVA, Mônica de Souza Alves. [Título do trabalho]. 2015. [Tipo de trabalho e instituição].',
'SUMIDA, Edson Tadashi. [Título da obra ou artigo]. [Local]: [Editora ou periódico], 2001.',
'ZAFFARONI, Eugenio Raúl; BATISTA, Nilo; ALAGIA, Alejandro; SLOKAR, Alejandro. Direito penal brasileiro: teoria geral do direito penal. Rio de Janeiro: Revan, 2003. v. 1.',
]
for r_ in refs:
    P(r_, indent=False, ls=1.0, after=12).alignment = WD_ALIGN_PARAGRAPH.LEFT

for s in d.sections:
    fp = s.header.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = fp.add_run()
    for tag, txt in (('begin', None), (None, 'PAGE'), ('end', None)):
        e = OxmlElement('w:fldChar' if tag else 'w:instrText')
        if tag: e.set(qn('w:fldCharType'), tag)
        else: e.text = txt
        r._r.append(e)

d.save('/home/user/Keyla/Projeto_TCC1_Nilton_Corrigido.docx'); print('ok')
