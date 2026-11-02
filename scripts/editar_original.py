import sys, copy, docx
src, out = sys.argv[1], sys.argv[2]
d = docx.Document(src)

def rep(p, old, new):
    runs = p.runs; text = ''.join(r.text for r in runs)
    i = text.find(old)
    if i < 0: return False
    j = i + len(old); pos = 0; first = True
    for r in runs:
        a, b = pos, pos + len(r.text); pos = b
        if b <= i or a >= j: continue
        s, e = max(i, a) - a, min(j, b) - a
        if first: r.text = r.text[:s] + new + r.text[e:]; first = False
        else: r.text = r.text[:s] + r.text[e:]
    return True

H = '‑'
E = [
('A remição de pena pelo trabalho e pelo estudo no conjunto penal de Feira de Santana- Ba', 'A remição de pena pelo trabalho e pelo estudo no Conjunto Penal de Feira de Santana-BA.'),
('Como se dá o acesso efetivo dos detentos do conjunto Penal de Feira de Santana ao direito de remição de pena por estudo e trabalho e quais são os principais “gargalos” institucionais e burocráticos que limitam a absorção da população carceraria nessas atividades?',
 'Quais fatores institucionais e burocráticos limitam o acesso efetivo das pessoas privadas de liberdade do Conjunto Penal de Feira de Santana à remição de pena pelo trabalho e pelo estudo?'),
('pelo estudo, previstos no art. 126', 'pelo estudo, prevista no art. 126'),
('Em um sistema carcerário que costuma ficar tensionado', 'Em um sistema carcerário tensionado'),
('pelo trabalho e estudo torna'+H+'se indispensável para verificar se este instituto cumpre seu papel constitucional de reintegração ou se apenas mantém indivíduos segregados sem preparo',
 'pelo trabalho e pelo estudo torna-se indispensável para verificar se o instituto cumpre sua função de integração social ou se apenas mantém indivíduos segregados, sem preparo'),
('Nesta ótica, a pesquisa destaca'+H+'se', 'Nessa ótica, a pesquisa destaca-se'),
('reduza o tempo de cárcere e saia do presídio no menor prazo legalmente admissível. Muito além de um simples cálculo de liquidação da pena, a rotina produtiva e formativa cria uma ponte crucial com o meio externo: ao desenvolver habilidades técnicas, cultivar disciplina laboral e restaurar dignidade pessoal, rompe o ciclo de isolamento discriminante, viabilizando que o egresso reconquiste vínculos familiares saudáveis e encontre',
 'reduza o tempo de cárcere no limite legalmente admissível. Além do aspecto aritmético da liquidação da pena, a rotina produtiva e formativa estabelece um vínculo com o meio externo: ao desenvolver habilidades técnicas, disciplina laboral e senso de dignidade pessoal, contribui para romper o ciclo de isolamento e exclusão, favorecendo que o egresso restabeleça vínculos familiares e encontre'),
('diagnosticar os desafios estruturais do Conjunto', 'diagnosticar os desafios estruturais e procedimentais do Conjunto'),
('Verificar como se dá o acesso efetivo dos detentos do conjunto Penal de Feira de Santana ao direito de remição de pena por estudo e trabalho e quais são os principais “gargalos” institucionais e burocráticos que limitam a absorção da população carceraria nessas atividades.',
 'Analisar os fatores institucionais e burocráticos que limitam o acesso efetivo das pessoas privadas de liberdade do Conjunto Penal de Feira de Santana à remição de pena pelo trabalho e pelo estudo.'),
('destacando suas diferenças e particularidades.', 'destacando suas diferenças e particularidades;'),
('Demonstrar a possibilidade da pessoa privada de liberdade acumular a remição de pena pelo trabalho e pelo estudo, valorizando essas atividades como instrumentos importantes para sua formação e ressocialização.',
 'Analisar a possibilidade de cumulação da remição de pena pelo trabalho e pelo estudo (art. 126, § 3º, da LEP) e sua importância para a formação e a ressocialização da pessoa privada de liberdade;'),
('que possibilitam a remição da pena.', 'que possibilitam a remição da pena, inclusive no fluxo administrativo e judicial de registro e homologação dos dias remidos;'),
('Quanto à natureza desta pesquisa é aplicada', 'Quanto à natureza, esta pesquisa é aplicada'),
('e estudo de caso, que vai à observação direta da realidade prática na unidade prisional.', 'e estudo de caso, voltado à realidade prática da unidade prisional. O método de abordagem é o dedutivo.'),
('especialmente o artigo 126 e os seguintes, além de', 'especialmente os artigos 126 a 130, as Leis nº 12.433/2011 e nº 13.163/2015, além de'),
('estudo de caso centrado no Conjunto Penal de Feira de Santana.', 'estudo de caso centrado no Conjunto Penal de Feira de Santana, com base em dados oficiais do SISDEPEN e da SEAP-BA, em informações solicitadas à unidade, inclusive por meio da Lei nº 12.527/2011 (Lei de Acesso à Informação), e em decisões públicas da Vara de Execuções Penais da comarca, no período de [definir].'),
('será utilizada como fonte de dados para entender a realidade prática.', 'será utilizada como ponto de observação para compreender a realidade prática, com registros em diário de campo e confronto com os dados documentais (triangulação de fontes).'),
('os direitos fundamentais das pessoas privadas de liberdade. Assim,', 'os direitos fundamentais das pessoas privadas de liberdade, não sendo utilizados dados protegidos por sigilo funcional nem informações que permitam identificar internos ou servidores. Eventual interação direta com pessoas dependerá de autorização da SEAP-BA e de aprovação em Comitê de Ética em Pesquisa (Resolução CNS nº 510/2016). Assim,'),
('no curso do processo executório', 'no curso do processo executório.'),
('A Carta Magna assegura explicitamente', 'A Constituição Federal assegura expressamente'),
('como pilares indispensáveis da fase executiva', 'como pilares indispensáveis da fase executiva.'),
('(remissão, comutação', '(remissão [sic], comutação'),
('a dupla finalidade da sanção penal:', 'a dupla finalidade da execução penal:'),
('dos direitos fundamentais da pessoa presa.', 'dos direitos fundamentais da pessoa presa (ROIG, 2016; ZAFFARONI et al., 2003) [conferir obras e páginas].'),
('Diante do Estado de Coisas Inconstitucional que caracteriza o sistema prisional,', 'Diante do estado de coisas inconstitucional que caracteriza o sistema prisional, reconhecido pelo STF na ADPF 347,'),
('o direito subjetivo público do apenado de abater o tempo de sua condenação mediante o trabalho, o estudo ou a leitura.', 'o direito subjetivo do apenado de abater o tempo de sua condenação mediante o trabalho ou o estudo, incluída neste a leitura, por interpretação extensiva consolidada pelo CNJ.'),
('vem de redemere', 'vem de redimere'),
('mérito demonstrado pelo ordenamento', 'mérito demonstrado pelo condenado'),
('causa extintiva da punibilidade e que reduz a quantidade mínima de pena em corrente contínua.', 'causa extintiva da punibilidade [...].'),
('O cômputo do benefício opera-se', 'A qualificação da remição como causa extintiva da punibilidade, porém, não é pacífica, pois o art. 128 da LEP determina que o tempo remido seja computado como pena cumprida para todos os efeitos. Registre-se, ainda, que a definição de Mirabete é anterior à Lei nº 12.433/2011 e, por isso, menciona apenas o trabalho. O cômputo do benefício opera-se'),
('horas diárias (art. 126, § 1º, II, LEP).', 'horas diárias (art. 126, § 1º, II, e art. 33 da LEP), inclusive quando a atividade for exercida extramuros (Súmula 562 do STJ).'),
('O Superior Tribunal de Justiça veda a concessão da remição pelo trabalho no regime aberto, por constituir o trabalho um requisito ordinário deste regime.', 'O Superior Tribunal de Justiça não admite a remição pelo trabalho no regime aberto, por constituir o trabalho requisito do próprio regime (art. 114, I, da LEP). Em caso de falta grave, o juiz poderá revogar até 1/3 (um terço) do tempo remido (art. 127 da LEP).'),
('ensino fundamental, médio, superior ou profissionalizante,', 'ensino fundamental, médio, inclusive profissionalizante, ou superior, ou ainda de requalificação profissional,'),
('conforme o art. 126, § 5º, da LEP.', 'conforme o art. 126, § 5º, da LEP. Diferentemente do trabalho, a remição pelo estudo alcança também o regime aberto e o livramento condicional (§ 6º), e as horas de trabalho e de estudo podem ser cumuladas, desde que compatíveis (§ 3º).'),
('no âmbito das Práticas Sociais Educativas Complementares.', 'no âmbito das práticas sociais educativas. O artesanato, por sua vez, tem sido admitido pela jurisprudência como atividade laboral apta à remição, desde que supervisionado e registrado pela administração [conferir precedente].'),
('dispõe de 21 a 30 dias para a leitura', 'dispõe de [conferir o prazo na Resolução 391/2021] para a leitura'),
('dotado de finalidade educativa e produtiva.', 'dotado de finalidade educativa e produtiva, sendo também direito do preso (art. 41, II, da LEP).'),
('Corrente Majoritária (STF e STJ): Veda categoricamente', 'Corrente majoritária (STF e STJ): veda, como regra,'),
('O benefício da remição da pena pelo trabalho ou pelo estudo, consoante', '1. O benefício da remição da pena pelo trabalho ou pelo estudo, consoante'),
('por virtude de acidente.', 'por acidente. Registre-se, ainda, que o STJ, no julgamento do Tema Repetitivo 1.120 (REsp 1.953.607/SC), admitiu a remição em favor dos presos que já trabalhavam ou estudavam e foram impedidos de continuar as atividades em razão da suspensão imposta pelo Estado durante a pandemia de Covid-19 [conferir a tese no portal do STJ].'),
('Corrente Garantista: Defendida por doutrinadores como Rogério Greco e Rodrigo Duque Estrada Roig,', 'Corrente garantista: defendida por doutrinadores como Rogério Greco e Rodrigo Duque Estrada Roig [inserir obra e página],'),
('tencionado', 'tensionado'),
('No CPFS, identificam-se dois entraves centrais: Gargalo Estrutural: Manifesta-se', 'Parte-se da hipótese de que, no CPFS, dois entraves centrais limitam a efetividade da remição: a) gargalo estrutural: manifesta-se'),
('Essa desproporção viola a lógica dos arts. 83, § 4º, e 88 da LEP,', 'Essa desproporção, se confirmada, contraria o dever estatal de assegurar trabalho e assistência educacional (arts. 17 a 21, 28, 41, II, e 83, § 4º, da LEP),'),
('aos apenados; Gargalo Burocrático: Refere-se', 'aos apenados; b) gargalo burocrático: refere-se'),
('ao Juízo da Execução Penal (VEP), ao Ministério Público e à Defensoria Pública,', 'ao juízo da execução, com posterior manifestação do Ministério Público e da defesa (art. 126, § 8º),'),
('a progressão de regime e o livramento condiciona', 'a progressão de regime e o livramento condicional. Tais hipóteses serão verificadas no estudo de caso, a partir dos dados oficiais e documentais da unidade.'),
]
paras = d.paragraphs
fail = []
for old, new in E:
    ok = False
    for p in paras:
        if rep(p, old, new): ok = True; break
    if not ok:
        # tentativa sem diferenciar maiúscula inicial
        fail.append(old[:60])
# substituições globais
for p in paras:
    while rep(p, 'carceraria', 'carcerária'): pass
    while rep(p, 'conjunto Penal', 'Conjunto Penal'): pass
    while rep(p, 'p. 208) .', 'p. 208).'): pass
# pontos soltos (parágrafos que contêm apenas ".")
for p in paras:
    if p.text.strip() == '.':
        for r in p.runs: r.text = r.text.replace('.', '')

# rodapé (notas dos autores)
done = set()
for s in d.sections:
    for f in (s.footer, s.first_page_footer, s.even_page_footer):
        if id(f._element) in done: continue
        done.add(id(f._element))
        ps = f.paragraphs
        for p in ps:
            rep(p, 'Graduando/a em Direito, Unidade de Ensino Superior de Feira de Santana (UNEF)', 'Graduando em Direito, Unidade de Ensino Superior de Feira de Santana (UNEF).')
            rep(p, 'nome@email.com.br', 'E-mail: [e-mail do autor].')
            if 'advogadadaiane' in p.text and 'Coimbra' in p.text:
                new = copy.deepcopy(p._p)
                p._p.addnext(new)
                rep(p, 'Mestre em Ciências Jurídico Criminais (Universidade de Coimbra), Unidade de Ensino Superior de Feira de Santana (UNEF), advogadadaiane@hotmail.com.', '[Titulação do Prof. Alberto Azis], Unidade de Ensino Superior de Feira de Santana (UNEF), [e-mail].')
                from docx.text.paragraph import Paragraph
                q = Paragraph(new, p._parent)
                rep(q, 'Mestre em Ciências Jurídico Criminais', 'Mestre em Ciências Jurídico-Criminais')
                for r in q.runs:
                    if r.text.strip() == '2': r.text = r.text.replace('2', '3')
                break
d.save(out)
print('falhas:', fail)
