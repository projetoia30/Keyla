from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

d = Document()
for s in d.sections:
    s.top_margin = s.left_margin = Cm(3); s.bottom_margin = s.right_margin = Cm(2)
st = d.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
pf = st.paragraph_format; pf.line_spacing = 1.5; pf.space_after = Pt(0)

def P(t, center=False, bold=False, indent=True, size=None, just=True):
    p = d.add_paragraph(); r = p.add_run(t); r.bold = bold
    if size: r.font.size = Pt(size)
    if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        if just: p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if indent: p.paragraph_format.first_line_indent = Cm(1.25)
    return p

def H(t, lvl=1):
    p = d.add_paragraph(); r = p.add_run(t.upper() if lvl == 1 else t); r.bold = True
    r.font.size = Pt(12); p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(12)
    p.style = d.styles['Heading %d' % lvl]
    for run in p.runs:
        run.font.name = 'Times New Roman'; run.font.size = Pt(12); run.font.color.rgb = None; run.bold = True
        run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

def Q(t):  # citação longa
    p = d.add_paragraph(); r = p.add_run(t); r.font.size = Pt(10)
    p.paragraph_format.left_indent = Cm(4); p.paragraph_format.line_spacing = 1.0
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.space_after = Pt(12); p.paragraph_format.space_before = Pt(12)

def brk(): d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

def toc():
    p = d.add_paragraph(); r = p.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = 'TOC \\o "1-2" \\h \\z \\u'
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = 'Clique com o botão direito e escolha "Atualizar campo" para gerar o sumário.'
    f3 = OxmlElement('w:fldChar'); f3.set(qn('w:fldCharType'), 'end')
    for e in (f1, it, f2, t, f3): r._r.append(e)

# CAPA
for t in ['[NOME DA INSTITUIÇÃO]', '[CURSO DE DIREITO]', '', '', '[NOME DO(A) AUTOR(A)]', '', '', '', '']:
    P(t, center=True, bold=True)
P('REMIÇÃO DE PENA PELO TRABALHO E PELO ESTUDO: ALCANCE E LIMITES DA RESSOCIALIZAÇÃO NO CONJUNTO PENAL DE FEIRA DE SANTANA – BA', center=True, bold=True)
for t in ['', '', '', '', '', '', 'FEIRA DE SANTANA – BA', '2026']: P(t, center=True, bold=True)
brk()
# FOLHA DE ROSTO
P('[NOME DO(A) AUTOR(A)]', center=True, bold=True)
for _ in range(5): P('')
P('REMIÇÃO DE PENA PELO TRABALHO E PELO ESTUDO: ALCANCE E LIMITES DA RESSOCIALIZAÇÃO NO CONJUNTO PENAL DE FEIRA DE SANTANA – BA', center=True, bold=True)
P('')
q = d.add_paragraph(); q.paragraph_format.left_indent = Cm(8); q.paragraph_format.line_spacing = 1.0
q.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
q.add_run('Trabalho de Conclusão de Curso apresentado ao Curso de Direito da [Instituição], como requisito parcial para obtenção do título de Bacharel em Direito.\n\nOrientador(a): Prof.(a) [Nome].').font.size = Pt(10)
for _ in range(6): P('')
P('FEIRA DE SANTANA – BA', center=True, bold=True); P('2026', center=True, bold=True)
brk()

# RESUMO
P('RESUMO', center=True, bold=True)
d.add_paragraph().paragraph_format.line_spacing = 1.0
r = P('O presente trabalho analisa o instituto da remição de pena pelo trabalho e pelo estudo, previsto nos artigos 126 a 130 da Lei de Execução Penal (Lei nº 7.210/1984), com as alterações promovidas pelas Leis nº 12.433/2011 e nº 13.163/2015, tendo como recorte empírico o Conjunto Penal de Feira de Santana, na Bahia. Parte-se do problema de pesquisa: em que medida a remição por trabalho e estudo tem sido efetivada na referida unidade prisional como instrumento de ressocialização? Adota-se abordagem qualiquantitativa, com pesquisa bibliográfica, documental e jurisprudencial, valendo-se de dados oficiais do Sistema de Informações do Departamento Penitenciário Nacional (SISDEPEN), da Secretaria de Administração Penitenciária e Ressocialização da Bahia (SEAP-BA) e de normativas do Conselho Nacional de Justiça (CNJ), notadamente a Resolução nº 391/2021. Os resultados indicam que, embora o arcabouço normativo brasileiro seja avançado e a jurisprudência do Superior Tribunal de Justiça tenha ampliado as hipóteses de remição, a efetividade do instituto é limitada pela insuficiência de vagas de trabalho e de educação, pela superlotação e por entraves burocráticos no registro e na homologação dos dias remidos. Conclui-se que a remição constitui instrumento relevante de individualização da pena e de redução da reincidência, mas depende de políticas públicas estruturadas, de parcerias interinstitucionais e da expansão da remição pela leitura para alcançar sua finalidade ressocializadora.', indent=False)
r.paragraph_format.line_spacing = 1.0
p = P('Palavras-chave: remição de pena; execução penal; ressocialização; trabalho prisional; educação prisional.', indent=False); p.paragraph_format.line_spacing = 1.0
brk()
P('ABSTRACT', center=True, bold=True)
r = P('This study analyzes the institute of sentence remission through work and study, established in articles 126 to 130 of the Brazilian Penal Execution Law (Law No. 7,210/1984), as amended by Laws No. 12,433/2011 and No. 13,163/2015, with an empirical focus on the Feira de Santana Penal Complex, in the State of Bahia. The research problem is: to what extent has remission through work and study been implemented in that prison unit as an instrument of social reintegration? A qualitative and quantitative approach is adopted, based on bibliographic, documentary and case-law research, using official data from SISDEPEN, from the Bahia Penitentiary Administration Department (SEAP-BA) and regulations issued by the National Council of Justice (CNJ), particularly Resolution No. 391/2021. The results indicate that, although the Brazilian legal framework is advanced and the Superior Court of Justice has broadened the possibilities of remission, its effectiveness is limited by the shortage of work and educational opportunities, overcrowding and bureaucratic obstacles. It is concluded that remission is a relevant tool for individualizing punishment and reducing recidivism, but depends on structured public policies, inter-institutional partnerships and the expansion of remission through reading.', indent=False)
r.paragraph_format.line_spacing = 1.0
p = P('Keywords: sentence remission; penal execution; social reintegration; prison labor; prison education.', indent=False); p.paragraph_format.line_spacing = 1.0
brk()

P('LISTA DE ABREVIATURAS E SIGLAS', center=True, bold=True)
for s_ in ['CNJ – Conselho Nacional de Justiça', 'CPFS – Conjunto Penal de Feira de Santana', 'CF – Constituição Federal', 'DEPEN – Departamento Penitenciário Nacional', 'ENCCEJA – Exame Nacional para Certificação de Competências de Jovens e Adultos', 'ENEM – Exame Nacional do Ensino Médio', 'LEP – Lei de Execução Penal', 'SEAP-BA – Secretaria de Administração Penitenciária e Ressocialização da Bahia', 'SENAPPEN – Secretaria Nacional de Políticas Penais', 'SISDEPEN – Sistema de Informações do Departamento Penitenciário Nacional', 'STF – Supremo Tribunal Federal', 'STJ – Superior Tribunal de Justiça', 'TJBA – Tribunal de Justiça do Estado da Bahia', 'VEP – Vara de Execuções Penais']:
    P(s_, indent=False, just=False)
brk()
P('SUMÁRIO', center=True, bold=True); toc(); brk()

# 1
H('1 INTRODUÇÃO')
for t in [
'O sistema prisional brasileiro figura entre os maiores do mundo em números absolutos e convive, há décadas, com superlotação, déficit de vagas e condições de custódia que levaram o Supremo Tribunal Federal a reconhecer, na Arguição de Descumprimento de Preceito Fundamental nº 347, o "estado de coisas inconstitucional" do sistema carcerário nacional. Nesse cenário, a execução penal enfrenta o desafio de conciliar a retribuição pelo delito com a finalidade, expressamente declarada no artigo 1º da Lei de Execução Penal (LEP), de "proporcionar condições para a harmônica integração social do condenado e do internado" (BRASIL, 1984).',
'Entre os instrumentos que a legislação oferece para concretizar essa finalidade, destaca-se a remição de pena, instituto pelo qual o condenado que cumpre pena em regime fechado ou semiaberto pode abreviar parte do tempo de execução por meio do trabalho ou do estudo (art. 126 da LEP). Trata-se de mecanismo que, ao mesmo tempo, estimula a ocupação produtiva, eleva a escolaridade da população carcerária, contribui para a disciplina interna das unidades e funciona como incentivo concreto à mudança de trajetória de vida.',
'O município de Feira de Santana, segundo maior da Bahia e importante entroncamento rodoviário do Nordeste, abriga o Conjunto Penal de Feira de Santana (CPFS), uma das principais unidades prisionais do interior do Estado, vinculada à Secretaria de Administração Penitenciária e Ressocialização da Bahia (SEAP-BA). A unidade concentra presos provisórios e condenados de diversos regimes, oriundos de Feira de Santana e de municípios da região, o que a torna recorte relevante para a investigação da efetividade da remição no contexto baiano.',
'Diante disso, formula-se o seguinte problema de pesquisa: em que medida a remição de pena por trabalho e estudo tem sido efetivada no Conjunto Penal de Feira de Santana como instrumento de ressocialização? Parte-se da hipótese de que, embora o marco normativo seja amplo e progressivo, a efetividade da remição na unidade é limitada por fatores estruturais — insuficiência de vagas de trabalho e de educação, superlotação, déficit de servidores — e por entraves procedimentais no registro e na homologação judicial dos dias remidos.',
'O objetivo geral consiste em analisar a efetividade da remição de pena por trabalho e estudo no Conjunto Penal de Feira de Santana. São objetivos específicos: (a) examinar os fundamentos teóricos da pena e da finalidade ressocializadora da execução penal; (b) analisar o regime jurídico da remição, sua evolução legislativa e a interpretação conferida pelos tribunais superiores; (c) caracterizar o Conjunto Penal de Feira de Santana e as políticas de trabalho e educação nele desenvolvidas; e (d) identificar os principais obstáculos à concretização do instituto e propor medidas para o seu aperfeiçoamento.',
'A pesquisa justifica-se em três planos. No plano social, a remição relaciona-se diretamente à redução da reincidência e, por consequência, à segurança pública. No plano jurídico, a ampliação jurisprudencial do instituto — especialmente quanto à remição pela leitura e às atividades extramuros — demanda sistematização. No plano local, há escassez de estudos acadêmicos voltados especificamente à realidade prisional de Feira de Santana, o que confere originalidade à investigação.',
'O trabalho está estruturado em oito seções. Após esta introdução, a seção 2 descreve a metodologia. A seção 3 aborda as teorias da pena e a ressocialização. A seção 4 examina o regime jurídico da remição. A seção 5 trata da remição pela leitura e da jurisprudência dos tribunais superiores. A seção 6 caracteriza o sistema prisional baiano e o Conjunto Penal de Feira de Santana. A seção 7 apresenta a análise crítica da efetividade do instituto e, por fim, a seção 8 reúne as considerações finais.']: P(t)

H('2 METODOLOGIA')
for t in [
'Quanto à abordagem, a pesquisa é qualiquantitativa: qualitativa na interpretação do ordenamento jurídico, da doutrina e da jurisprudência; quantitativa na análise de dados oficiais sobre população carcerária e acesso a trabalho e educação. Quanto aos objetivos, é exploratória e descritiva, pois busca conhecer e descrever a realidade da remição na unidade estudada (GIL, 2019).',
'Quanto aos procedimentos técnicos, adota-se a pesquisa bibliográfica, fundada em obras de referência do Direito Penal, da Execução Penal e da Criminologia, bem como em artigos publicados em periódicos científicos; a pesquisa documental, com base na legislação federal e estadual, nas resoluções e recomendações do Conselho Nacional de Justiça e nos relatórios do SISDEPEN e da SEAP-BA; e a pesquisa jurisprudencial, com levantamento de julgados do STF, do STJ e do TJBA sobre remição.',
'O método de abordagem é o dedutivo, partindo-se das premissas gerais do ordenamento jurídico e da teoria da pena para a análise do caso concreto do Conjunto Penal de Feira de Santana (MARCONI; LAKATOS, 2017). O recorte temporal compreende, prioritariamente, o período posterior à Lei nº 12.433/2011, que instituiu a remição pelo estudo, até a atualidade.',
'Registra-se que a pesquisa não envolve coleta direta de dados com seres humanos, razão pela qual dispensa submissão a Comitê de Ética em Pesquisa, nos termos da Resolução CNS nº 510/2016. Eventual desdobramento empírico, com entrevistas a internos, servidores ou operadores do Direito, exigirá a referida aprovação prévia.']: P(t)

H('3 A PENA E A FINALIDADE RESSOCIALIZADORA')
H('3.1 Teorias da pena', 2)
for t in [
'A compreensão da remição pressupõe o debate acerca das finalidades da pena. As teorias absolutas, associadas a Kant e Hegel, concebem a pena como retribuição justa ao mal praticado, desvinculada de qualquer finalidade utilitária. As teorias relativas, por sua vez, atribuem à pena uma função preventiva, dividida em prevenção geral — voltada à sociedade, seja pela intimidação (negativa), seja pela reafirmação da validade da norma (positiva) — e prevenção especial — voltada ao próprio condenado, seja pela neutralização (negativa), seja pela ressocialização (positiva) (BITENCOURT, 2023).',
'As teorias mistas ou unificadoras procuram conciliar retribuição e prevenção. No Brasil, o artigo 59 do Código Penal determina que a pena seja estabelecida "conforme seja necessário e suficiente para reprovação e prevenção do crime", o que evidencia a adoção de uma concepção eclética. Na fase executória, contudo, a LEP privilegia a prevenção especial positiva, ao estabelecer como objetivo a integração social do condenado.',
'Na perspectiva crítica, Foucault (2014) demonstra que a prisão moderna nasceu associada à disciplina dos corpos e ao trabalho como técnica de docilização, ao passo que Baratta (2011) denuncia a seletividade do sistema penal e o caráter muitas vezes ilusório do discurso ressocializador. Essas leituras são indispensáveis para que a remição não seja tratada como mero benefício formal, mas analisada à luz das condições reais de cumprimento da pena.']: P(t)
H('3.2 A ressocialização na Lei de Execução Penal', 2)
for t in [
'A LEP, editada em 1984 sob forte influência das Regras Mínimas da ONU para o Tratamento de Reclusos, estruturou a execução penal brasileira em torno da individualização da pena e da assistência ao preso. O artigo 10 estabelece que a assistência ao preso é dever do Estado, "objetivando prevenir o crime e orientar o retorno à convivência em sociedade", compreendendo as dimensões material, à saúde, jurídica, educacional, social e religiosa (art. 11).',
'O trabalho do condenado é definido como "dever social e condição de dignidade humana", com finalidade educativa e produtiva (art. 28), sendo remunerado em valor não inferior a três quartos do salário mínimo (art. 29). A assistência educacional compreende a instrução escolar e a formação profissional (art. 17), sendo obrigatório o ensino fundamental (art. 18) e, desde a Lei nº 13.163/2015, também ofertado o ensino médio (art. 18-A).',
'Como observa Nucci (2023), a ressocialização não pode ser imposta, mas deve ser oferecida: cabe ao Estado proporcionar oportunidades reais de trabalho e estudo, e ao condenado a decisão de aproveitá-las. A remição atua exatamente nesse ponto, ao converter a adesão voluntária às atividades em benefício concreto de abreviação da pena. No plano internacional, as Regras de Mandela (ONU, 2015) reforçam que o tempo de prisão deve ser utilizado para assegurar a reinserção social da pessoa privada de liberdade.']: P(t)

H('4 A REMIÇÃO DE PENA NO ORDENAMENTO JURÍDICO BRASILEIRO')
H('4.1 Conceito e natureza jurídica', 2)
for t in [
'Remição, do latim redimere, significa resgatar. No âmbito da execução penal, consiste no desconto de parte do tempo de cumprimento da pena em razão do trabalho ou do estudo realizado pelo condenado. Não se confunde com "remissão" (perdão), grafia que designa instituto distinto do Direito Civil e Tributário.',
'Quanto à natureza jurídica, a doutrina majoritária a qualifica como direito subjetivo do condenado, uma vez preenchidos os requisitos legais, e não como mero benefício concedido discricionariamente (MIRABETE; FABBRINI, 2023). Nos termos do artigo 128 da LEP, o tempo remido é computado como pena cumprida para todos os efeitos, inclusive para fins de progressão de regime e de livramento condicional.']: P(t)
H('4.2 Evolução legislativa', 2)
for t in [
'Em sua redação original, a LEP previa apenas a remição pelo trabalho, na proporção de um dia de pena por três dias trabalhados, e determinava a perda integral dos dias remidos em caso de falta grave. A remição pelo estudo, embora não prevista expressamente, passou a ser admitida pela jurisprudência, culminando na edição da Súmula 341 do STJ (2007): "A frequência a curso de ensino formal é causa de remição de parte do tempo de execução de pena sob regime fechado ou semiaberto".',
'A Lei nº 12.433/2011 representou o principal marco de reforma do instituto: positivou a remição pelo estudo, admitiu a modalidade presencial e a distância, estendeu a remição pelo estudo ao regime aberto e ao livramento condicional, permitiu a cumulação de trabalho e estudo, criou o acréscimo de 1/3 pela conclusão de nível de ensino e limitou a perda dos dias remidos por falta grave a, no máximo, 1/3. Posteriormente, a Lei nº 13.163/2015 instituiu o ensino médio nas unidades prisionais, ampliando as possibilidades de remição educacional.']: P(t)
H('4.3 Remição pelo trabalho', 2)
for t in [
'Nos termos do artigo 126, § 1º, II, da LEP, o condenado em regime fechado ou semiaberto remirá um dia de pena a cada três dias de trabalho. A jornada deve observar o artigo 33 da LEP, não inferior a seis nem superior a oito horas diárias, com descanso aos domingos e feriados. O STJ consolidou o entendimento de que as horas excedentes à jornada podem ser computadas para fins de remição, à razão de um dia a cada dezoito horas acumuladas.',
'A Súmula 562 do STJ pacificou que "é possível a remição de parte do tempo de execução da pena quando o condenado, em regime fechado ou semiaberto, desempenha atividade laborativa, ainda que extramuros". O trabalho pode ser interno (manutenção, cozinha, limpeza, oficinas) ou externo, inclusive mediante convênios com o poder público e com a iniciativa privada. O preso impossibilitado de prosseguir no trabalho por acidente continua a beneficiar-se da remição (art. 126, § 4º).',
'A remição pelo trabalho não alcança, em regra, o condenado em regime aberto, pois nesse regime o trabalho é pressuposto do próprio regime (art. 114, I, da LEP). Essa distinção não se aplica à remição pelo estudo, expressamente estendida ao regime aberto pelo § 6º do artigo 126.']: P(t)
H('4.4 Remição pelo estudo', 2)
for t in [
'O artigo 126, § 1º, I, da LEP estabelece a remição de um dia de pena a cada doze horas de frequência escolar — atividade de ensino fundamental, médio, profissionalizante, superior ou de requalificação profissional — divididas, no mínimo, em três dias. As atividades podem ser desenvolvidas presencialmente ou por ensino a distância, devendo ser certificadas pelas autoridades educacionais competentes (art. 126, § 2º).',
'O § 5º do artigo 126 prevê que o tempo a remir em função das horas de estudo será acrescido de 1/3 no caso de conclusão do ensino fundamental, médio ou superior durante o cumprimento da pena, desde que certificada pelo órgão competente. O STJ tem reconhecido que a aprovação no ENCCEJA ou no ENEM, ainda que obtida por estudo autodidata, autoriza a remição, conforme parâmetros fixados pelo CNJ, considerando a carga horária correspondente ao nível de ensino certificado.',
'Admite-se, ainda, a cumulação das horas de trabalho e de estudo, desde que compatibilizadas (art. 126, § 3º), o que permite ao condenado ampliar o tempo remido, sem que a mesma atividade seja contabilizada em duplicidade.']: P(t)
H('4.5 Perda dos dias remidos e falta grave', 2)
for t in [
'Segundo o artigo 127 da LEP, com a redação dada pela Lei nº 12.433/2011, "em caso de falta grave, o juiz poderá revogar até 1/3 (um terço) do tempo remido, observado o disposto no art. 57, recomeçando a contagem a partir da data da infração disciplinar". A alteração substituiu a perda integral, que havia sido considerada constitucional pela Súmula Vinculante nº 9 do STF, por um regime proporcional e fundamentado.',
'A decisão de revogação deve ser motivada, considerando a natureza, os motivos, as circunstâncias e as consequências da falta, bem como a pessoa do faltoso e seu tempo de prisão (art. 57). A revogação exige prévio procedimento disciplinar com garantia de contraditório e ampla defesa, inclusive com defesa técnica, conforme orientação do STJ.',
'Por fim, o artigo 129 impõe à autoridade administrativa o dever de encaminhar mensalmente ao juízo da execução a relação dos presos que trabalham ou estudam, com informações sobre dias de trabalho ou horas de frequência escolar. A remição é declarada pelo juiz da execução, ouvidos o Ministério Público e a defesa (art. 126, § 8º). O artigo 130, por sua vez, tipifica como falsidade ideológica a declaração ou o atestado falso de prestação de serviço para fins de remição.']: P(t)

H('5 REMIÇÃO PELA LEITURA E ENTENDIMENTOS DOS TRIBUNAIS SUPERIORES')
H('5.1 A Resolução CNJ nº 391/2021', 2)
for t in [
'A remição pela leitura não está expressamente prevista na LEP, mas foi construída a partir de interpretação extensiva do conceito de estudo, inicialmente por meio da Portaria Conjunta nº 276/2012 (CJF/DEPEN), aplicável ao sistema penitenciário federal, e da Recomendação CNJ nº 44/2013. Atualmente, a matéria é disciplinada pela Resolução CNJ nº 391/2021, que estabelece procedimentos e diretrizes para o reconhecimento do direito à remição por meio de práticas sociais educativas.',
'A resolução prevê que a pessoa privada de liberdade poderá remir quatro dias de pena por obra lida e avaliada, mediante elaboração de relatório de leitura ou resenha, limitado a doze obras por ano, totalizando até 48 dias de remição anuais. A avaliação deve ser realizada por comissão designada, observando critérios pedagógicos e respeitando as condições de letramento do leitor, inclusive com adaptações para pessoas com deficiência e em processo de alfabetização.',
'A norma também reconhece a remição por atividades educacionais complementares — culturais, esportivas, de capacitação profissional e de saúde, entre outras — e pela aprovação em exames nacionais de certificação, consolidando um modelo ampliado de educação no cárcere. Para unidades com baixa oferta de vagas escolares, como é a realidade de grande parte do sistema baiano, a remição pela leitura representa alternativa de baixo custo e alto impacto.']: P(t)
H('5.2 Jurisprudência do STF e do STJ', 2)
for t in [
'O STJ tem desempenhado papel central na ampliação do instituto. Além das Súmulas 341 e 562, já mencionadas, a Corte consolidou entendimentos no sentido de que: (a) a remição pela leitura é legítima, por constituir atividade de estudo; (b) a aprovação no ENEM ou no ENCCEJA enseja remição ainda que o apenado não esteja matriculado em curso formal; (c) as horas de trabalho excedentes à jornada legal devem ser consideradas; e (d) a revogação dos dias remidos exige fundamentação concreta quanto à fração aplicada.',
'Durante a pandemia de Covid-19, discutiu-se a chamada remição ficta, isto é, o cômputo de dias remidos em favor de presos que, já inseridos em atividades laborais ou educacionais, foram impedidos de exercê-las em razão da suspensão das atividades nas unidades. O STJ, em sede de recursos repetitivos (Tema 1.120), firmou entendimento favorável à remição nessa hipótese excepcional, por não ser possível imputar ao condenado a interrupção causada pelo próprio Estado, mantendo-se, contudo, a regra geral de que a remição pressupõe o efetivo exercício da atividade. [Recomenda-se conferir a tese e o número do tema no portal do STJ antes do depósito.]',
'O STF, por sua vez, ao reconhecer o estado de coisas inconstitucional na ADPF 347, determinou a elaboração de planos nacional e estaduais para enfrentamento da crise prisional, nos quais a ampliação do acesso a trabalho e educação constitui eixo estruturante. Tais decisões reforçam que a efetividade da remição depende não apenas da atuação judicial, mas de políticas públicas consistentes.']: P(t)

H('6 O CONJUNTO PENAL DE FEIRA DE SANTANA')
H('6.1 Panorama do sistema prisional baiano', 2)
for t in [
'A Bahia possui uma das maiores populações carcerárias do Nordeste, distribuída entre unidades da capital — com destaque para o Complexo Penitenciário da Mata Escura, em Salvador — e conjuntos penais no interior, como os de Feira de Santana, Vitória da Conquista, Jequié, Itabuna, Juazeiro e Teixeira de Freitas. A gestão é atribuída à SEAP-BA, que coordena as políticas de custódia, ressocialização e assistência ao egresso.',
'Os relatórios do SISDEPEN indicam, de forma recorrente, que apenas uma parcela minoritária da população prisional brasileira está envolvida em atividades laborais ou educacionais, e que o perfil predominante é de pessoas jovens, negras e com baixa escolaridade, a maioria sem ensino fundamental completo. [Inserir aqui os dados mais recentes do SISDEPEN referentes à Bahia: população total, déficit de vagas e percentuais em trabalho e estudo.] Esse perfil evidencia o potencial transformador da remição pelo estudo, mas também a dimensão do desafio.']: P(t)
H('6.2 Caracterização da unidade', 2)
for t in [
'O Conjunto Penal de Feira de Santana, localizado no município homônimo, é unidade de grande porte destinada à custódia de presos provisórios e condenados, de ambos os sexos, contando com módulos específicos, inclusive para mulheres. Atende à demanda da comarca de Feira de Santana e de municípios circunvizinhos do Portal do Sertão, sendo a execução penal acompanhada pela Vara de Execuções Penais da comarca. [Inserir capacidade oficial, população atual e ano de referência conforme SEAP-BA/SISDEPEN.]',
'A unidade enfrenta problemas comuns às grandes unidades do sistema brasileiro, frequentemente noticiados e registrados em inspeções do CNJ, da Defensoria Pública e do Mecanismo Nacional de Prevenção e Combate à Tortura: ocupação acima da capacidade, limitações de infraestrutura e déficit de pessoal, fatores que repercutem diretamente na capacidade de oferta de atividades laborais e educacionais.']: P(t)
H('6.3 Oferta de trabalho e educação na unidade', 2)
for t in [
'No campo educacional, a oferta de ensino nas unidades prisionais baianas ocorre por meio de parceria entre a SEAP-BA e a Secretaria da Educação do Estado (SEC-BA), com turmas de Educação de Jovens e Adultos (EJA), além da aplicação do ENCCEJA e do ENEM para Pessoas Privadas de Liberdade (ENEM PPL). Projetos de remição pela leitura, com bibliotecas e comissões avaliadoras, vêm sendo implementados no Estado em consonância com a Resolução CNJ nº 391/2021. [Inserir número de internos matriculados e de participantes de projetos de leitura no CPFS.]',
'No campo laboral, predominam as atividades internas de apoio — limpeza, cozinha, manutenção e serviços gerais —, complementadas por oficinas e cursos de qualificação profissional, bem como por convênios com órgãos públicos e empresas para trabalho interno ou externo, nos moldes do Programa Começar de Novo (CNJ) e da Política Nacional de Trabalho no âmbito do Sistema Prisional (Decreto nº 9.450/2018). [Inserir número de vagas de trabalho e principais convênios vigentes.]',
'A comparação entre o número de internos e o número de vagas disponíveis em trabalho e estudo constitui o indicador central de efetividade da remição na unidade, pois revela quantos presos, em tese aptos ao benefício, permanecem sem acesso a ele por ausência de oferta estatal.']: P(t)

H('7 EFETIVIDADE DA REMIÇÃO: ANÁLISE CRÍTICA')
H('7.1 Obstáculos estruturais e institucionais', 2)
for t in [
'A análise do marco normativo e da realidade da unidade permite identificar um descompasso entre a "lei nos livros" e a "lei em ação". O primeiro e mais relevante obstáculo é a insuficiência de vagas: se o Estado não oferta trabalho e estudo em número compatível com a população, a remição converte-se em privilégio de poucos, gerando desigualdade entre presos em situação jurídica idêntica e, não raro, disputa interna pelas vagas existentes.',
'O segundo obstáculo é a superlotação, que compromete espaços que poderiam ser destinados a salas de aula, bibliotecas e oficinas, e direciona o efetivo de agentes para atividades de segurança, em detrimento da escolta e da supervisão das atividades ressocializadoras. O terceiro é de natureza procedimental: falhas no controle de frequência, atrasos no envio mensal dos registros ao juízo (art. 129 da LEP) e demora na homologação judicial retardam a progressão de regime e frustram a finalidade de incentivo do instituto.',
'Há, ainda, obstáculos de natureza cultural: a resistência de parte dos operadores e da sociedade em reconhecer a remição como direito, a percepção do trabalho prisional apenas como mão de obra barata e a baixa escolaridade dos internos, que dificulta o acesso à remição pela leitura sem acompanhamento pedagógico adequado.']: P(t)
H('7.2 Atuação da VEP, da Defensoria Pública e do Ministério Público', 2)
for t in [
'A efetivação da remição depende de articulação entre os atores da execução penal. À Vara de Execuções Penais cabe declarar os dias remidos e fiscalizar a regularidade das informações; à Defensoria Pública, que assiste a maior parte da população carcerária, cabe requerer a remição e impugnar revogações desproporcionais; ao Ministério Público, fiscalizar a execução e as condições da unidade; e ao Conselho da Comunidade, colaborar na captação de parcerias.',
'A implantação do Sistema Eletrônico de Execução Unificado (SEEU), promovida pelo CNJ, representa avanço importante, ao permitir o cálculo automatizado da pena e o acompanhamento em tempo real dos marcos de benefícios. Sua efetividade, contudo, depende da alimentação tempestiva dos dados de trabalho e estudo pela administração prisional, o que reforça a necessidade de integração entre os sistemas da SEAP-BA e do Poder Judiciário.']: P(t)
H('7.3 Remição e reincidência', 2)
for t in [
'Estudos nacionais sobre reincidência criminal, a exemplo das pesquisas conduzidas pelo Instituto de Pesquisa Econômica Aplicada (IPEA, 2015) e pelo DEPEN em parceria com universidades, indicam que a elevação da escolaridade e a qualificação profissional estão associadas a menores taxas de retorno ao sistema prisional. A remição, ao incentivar essas atividades, produz efeitos que transcendem a abreviação da pena, alcançando a empregabilidade do egresso e a redução dos custos sociais do encarceramento.',
'Nesse sentido, a remição deve ser compreendida não como benesse, mas como política pública de segurança. Cada vaga de trabalho ou estudo ofertada no Conjunto Penal de Feira de Santana representa, potencialmente, menor tempo de encarceramento, menor ociosidade e maior chance de reinserção social — resultado que interessa ao condenado, à unidade prisional e à comunidade feirense.',
'A partir da análise empreendida, propõem-se as seguintes medidas: (a) ampliação de convênios com o poder público municipal e com empresas da região, aproveitando a vocação comercial e logística de Feira de Santana; (b) expansão das turmas de EJA e de cursos profissionalizantes, inclusive em parceria com instituições de ensino superior locais, como a Universidade Estadual de Feira de Santana (UEFS); (c) estruturação de bibliotecas e de comissões permanentes de remição pela leitura; (d) informatização e integração do controle de frequência ao SEEU; e (e) capacitação de servidores para o acompanhamento das atividades ressocializadoras.']: P(t)

H('8 CONSIDERAÇÕES FINAIS')
for t in [
'O presente trabalho buscou responder em que medida a remição de pena por trabalho e estudo tem sido efetivada no Conjunto Penal de Feira de Santana como instrumento de ressocialização. Verificou-se que o ordenamento jurídico brasileiro dispõe de marco normativo avançado, reforçado pelas Leis nº 12.433/2011 e nº 13.163/2015, pela Resolução CNJ nº 391/2021 e por uma jurisprudência do STJ marcadamente ampliativa, que reconheceu a remição extramuros, pela leitura, pela aprovação em exames nacionais e, excepcionalmente, durante a suspensão de atividades na pandemia.',
'Constatou-se, entretanto, que a efetividade do instituto na unidade estudada encontra limites significativos, sobretudo na insuficiência de vagas de trabalho e estudo frente à população custodiada, na superlotação e em entraves procedimentais no registro e na homologação dos dias remidos. Confirma-se, assim, a hipótese inicialmente formulada: a remição é juridicamente consolidada, mas materialmente restrita.',
'Conclui-se que a remição constitui instrumento relevante de individualização da pena, de valorização da dignidade humana e de redução da reincidência, mas somente alcançará sua finalidade ressocializadora se acompanhada de políticas públicas estruturadas, de investimentos em educação e trabalho prisional e de articulação entre a SEAP-BA, o Poder Judiciário, a Defensoria Pública, o Ministério Público e a sociedade civil feirense.',
'Como sugestão para pesquisas futuras, recomenda-se a realização de estudo empírico com entrevistas a internos, servidores e operadores do Direito do Conjunto Penal de Feira de Santana, mediante aprovação em Comitê de Ética, bem como o acompanhamento longitudinal de egressos beneficiados pela remição, a fim de mensurar o impacto do instituto sobre a reincidência no contexto local.']: P(t)
brk()

P('REFERÊNCIAS', center=True, bold=True)
refs = [
'BARATTA, Alessandro. Criminologia crítica e crítica do direito penal: introdução à sociologia do direito penal. 6. ed. Rio de Janeiro: Revan, 2011.',
'BITENCOURT, Cezar Roberto. Tratado de direito penal: parte geral. 29. ed. São Paulo: SaraivaJur, 2023. v. 1.',
'BITENCOURT, Cezar Roberto. Falência da pena de prisão: causas e alternativas. 5. ed. São Paulo: Saraiva, 2017.',
'BRASIL. Constituição da República Federativa do Brasil de 1988. Brasília, DF: Presidência da República, 1988.',
'BRASIL. Decreto-Lei nº 2.848, de 7 de dezembro de 1940. Código Penal. Rio de Janeiro, 1940.',
'BRASIL. Lei nº 7.210, de 11 de julho de 1984. Institui a Lei de Execução Penal. Brasília, DF: Presidência da República, 1984.',
'BRASIL. Lei nº 12.433, de 29 de junho de 2011. Altera a Lei nº 7.210, de 11 de julho de 1984 (Lei de Execução Penal), para dispor sobre a remição de parte do tempo de execução da pena por estudo ou por trabalho. Brasília, DF, 2011.',
'BRASIL. Lei nº 13.163, de 9 de setembro de 2015. Modifica a Lei nº 7.210, de 11 de julho de 1984, para instituir o ensino médio nas penitenciárias. Brasília, DF, 2015.',
'BRASIL. Decreto nº 9.450, de 24 de julho de 2018. Institui a Política Nacional de Trabalho no âmbito do Sistema Prisional. Brasília, DF, 2018.',
'BRASIL. Superior Tribunal de Justiça. Súmula nº 341. A frequência a curso de ensino formal é causa de remição de parte do tempo de execução de pena sob regime fechado ou semiaberto. Brasília, DF, 2007.',
'BRASIL. Superior Tribunal de Justiça. Súmula nº 562. É possível a remição de parte do tempo de execução da pena quando o condenado, em regime fechado ou semiaberto, desempenha atividade laborativa, ainda que extramuros. Brasília, DF, 2016.',
'BRASIL. Supremo Tribunal Federal. Arguição de Descumprimento de Preceito Fundamental nº 347/DF. Relator: Min. Marco Aurélio. Redator do acórdão: Min. Luís Roberto Barroso. Brasília, DF, 2023.',
'BRASIL. Ministério da Justiça e Segurança Pública. Secretaria Nacional de Políticas Penais. SISDEPEN: levantamento nacional de informações penitenciárias. Brasília, DF: SENAPPEN, [ano de referência]. Disponível em: https://www.gov.br/senappen. Acesso em: [data].',
'BAHIA. Secretaria de Administração Penitenciária e Ressocialização. Unidades prisionais. Salvador: SEAP, [ano]. Disponível em: http://www.seap.ba.gov.br. Acesso em: [data].',
'CONSELHO NACIONAL DE JUSTIÇA. Recomendação nº 44, de 26 de novembro de 2013. Dispõe sobre atividades educacionais complementares para fins de remição da pena pelo estudo e estabelece critérios para a admissão pela leitura. Brasília, DF: CNJ, 2013.',
'CONSELHO NACIONAL DE JUSTIÇA. Resolução nº 391, de 10 de maio de 2021. Estabelece procedimentos e diretrizes a serem observados pelo Poder Judiciário para o reconhecimento do direito à remição de pena por meio de práticas sociais educativas em unidades de privação de liberdade. Brasília, DF: CNJ, 2021.',
'CONSELHO NACIONAL DE SAÚDE. Resolução nº 510, de 7 de abril de 2016. Brasília, DF: CNS, 2016.',
'FOUCAULT, Michel. Vigiar e punir: nascimento da prisão. 42. ed. Petrópolis: Vozes, 2014.',
'GIL, Antonio Carlos. Métodos e técnicas de pesquisa social. 7. ed. São Paulo: Atlas, 2019.',
'INSTITUTO DE PESQUISA ECONÔMICA APLICADA. Reincidência criminal no Brasil: relatório de pesquisa. Rio de Janeiro: IPEA, 2015.',
'MARCÃO, Renato. Curso de execução penal. 20. ed. São Paulo: SaraivaJur, 2023.',
'MARCONI, Marina de Andrade; LAKATOS, Eva Maria. Fundamentos de metodologia científica. 8. ed. São Paulo: Atlas, 2017.',
'MIRABETE, Julio Fabbrini; FABBRINI, Renato N. Execução penal. 15. ed. Indaiatuba: Foco, 2023.',
'NUCCI, Guilherme de Souza. Curso de execução penal. 6. ed. Rio de Janeiro: Forense, 2023.',
'ORGANIZAÇÃO DAS NAÇÕES UNIDAS. Regras mínimas das Nações Unidas para o tratamento de reclusos (Regras de Nelson Mandela). Brasília, DF: CNJ, 2016.',
'ROIG, Rodrigo Duque Estrada. Execução penal: teoria crítica. 5. ed. São Paulo: Thomson Reuters Brasil, 2021.',
]
for r_ in refs:
    p = P(r_, indent=False, just=False); p.paragraph_format.line_spacing = 1.0; p.paragraph_format.space_after = Pt(12)

# numeração de página
for s in d.sections:
    fp = s.footer.paragraphs[0] if False else s.header.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = fp.add_run()
    for tag, txt in (('begin', None), (None, 'PAGE'), ('end', None)):
        if tag:
            e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), tag)
        else:
            e = OxmlElement('w:instrText'); e.text = txt
        r._r.append(e)

d.save('/home/user/Keyla/TCC_Remicao_Pena_Feira_de_Santana.docx')
print('ok')
