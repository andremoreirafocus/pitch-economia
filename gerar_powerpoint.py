"""Gera o PowerPoint executivo e uma prévia a partir do caso do pitch.

Requer python-pptx e Pillow. Todos os elementos do PPTX são editáveis.
"""

from pathlib import Path
from math import cos, sin, pi
import sys
import textwrap

DEPS = Path('/tmp/pitch-economia-pptx-deps')
if DEPS.exists():
    sys.path.insert(0, str(DEPS))

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / 'pitch_economia.md'
assert SOURCE.exists(), 'Arquivo de origem não encontrado.'
OUT = ROOT / 'pitch_economia_executivo_v3.pptx'
PDF_OUT = ROOT / 'pitch_economia_executivo_v3_previa.pdf'
PREVIEWS = Path('/tmp/pitch-economia-preview-v3')
PREVIEWS.mkdir(exist_ok=True)

NAVY = '102C36'
NAVY2 = '1D3C46'
GREEN = 'B8EC80'
TEAL = '1C7767'
CREAM = 'F5F4EE'
INK = '17343E'
GRAY = '62747A'
LINE = 'DCE3DE'
WHITE = 'FFFFFF'
PALE = 'E8F1E5'
AMBER = 'EEBB69'

CAPEX = 4_000_000
FCF = 805_000
YEARS = 15
WACC = .35 * .17 + .65 * .128 * (1 - .34)
FLOWS = [-CAPEX] + [FCF] * YEARS


def npv(rate, flows):
    return sum(cf / (1 + rate) ** i for i, cf in enumerate(flows))


def irr(flows):
    low, high = 0., 1.
    for _ in range(120):
        mid = (low + high) / 2
        if npv(mid, flows) > 0:
            low = mid
        else:
            high = mid
    return (low + high) / 2


IRR = irr(FLOWS)
NPV = npv(WACC, FLOWS)
assert round(WACC * 100, 1) == 11.4
assert round(IRR * 100, 1) == 18.6
assert round(NPV / 1e6, 2) == 1.65

prs = Presentation()
prs.slide_width = Inches(13.333333)
prs.slide_height = Inches(7.5)
prs.core_properties.title = 'LogiBrasil | Solarização do centro de distribuição'
prs.core_properties.subject = 'Pitch executivo de economia — caso fictício'
prs.core_properties.author = 'LogiBrasil — caso acadêmico'
prs.core_properties.keywords = 'WACC, TIR, VPL, energia solar, pitch de economia'
prs.core_properties.comments = 'Versão 3. Caso original LogiBrasil. Layout da versão 2; notas de fala reduzidas para dois minutos. Premissas ilustrativas; indicadores recalculados.'

SCALE = 120
FONT = Path('/usr/share/fonts/truetype/lato')
images = []
layout_issues = []


class Slide:
    def __init__(self, index, label, dark=False):
        self.slide = prs.slides.add_slide(prs.slide_layouts[6])
        self.dark = dark
        self.index = index
        self.bg = NAVY if dark else CREAM
        self.slide.background.fill.solid()
        self.slide.background.fill.fore_color.rgb = RGBColor.from_string(self.bg)
        self.im = Image.new('RGB', (1600, 900), '#' + self.bg)
        self.draw = ImageDraw.Draw(self.im)
        self.text(.65, .35, 5.4, .3, 'LOGIBRASIL  /  INVESTIMENTOS', 11, GREEN if dark else TEAL, True)
        self.text(9.0, .35, 3.67, .3, 'CASO FICTÍCIO  •  VALORES ILUSTRATIVOS', 9, 'C4D1D2' if dark else GRAY, align='right')
        self.line(.65, .89, 12.68, .89, NAVY2 if dark else LINE, 1)
        self.line(.65, 6.93, 12.68, 6.93, NAVY2 if dark else LINE, 1)
        self.text(.65, 7.08, 11.5, .22, 'ECONOMIA E FINANÇAS  /  ' + label.upper(), 9, 'B8CACB' if dark else GRAY)
        self.text(12.05, 7.02, .63, .32, f'{index:02d}', 11, GREEN if dark else TEAL, True, align='right')

    def rect(self, x, y, w, h, fill, radius=False, border=None):
        sh = self.slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fill)
        if border:
            sh.line.color.rgb = RGBColor.from_string(border)
            sh.line.width = Pt(.7)
        else:
            sh.line.fill.background()
        if radius:
            sh.adjustments[0] = .1
        box = (round(x * SCALE), round(y * SCALE), round((x + w) * SCALE), round((y + h) * SCALE))
        self.draw.rounded_rectangle(box, radius=10, fill='#' + fill, outline='#' + border if border else None) if radius else self.draw.rectangle(box, fill='#' + fill, outline='#' + border if border else None)
        return sh

    def circle(self, x, y, d, color):
        sh = self.slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(color)
        sh.line.fill.background()
        self.draw.ellipse((x*SCALE,y*SCALE,(x+d)*SCALE,(y+d)*SCALE), fill='#'+color)

    def line(self, x1, y1, x2, y2, color=LINE, width=1):
        sh = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        sh.line.color.rgb = RGBColor.from_string(color)
        sh.line.width = Pt(width)
        self.draw.line((x1*SCALE,y1*SCALE,x2*SCALE,y2*SCALE), fill='#'+color, width=max(1,round(width*SCALE/72)))

    def text(self, x, y, w, h, value, size=18, color=INK, bold=False, align='left'):
        box = self.slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        box.text_frame.clear()
        box.text_frame.word_wrap = True
        box.text_frame.margin_left = box.text_frame.margin_right = 0
        box.text_frame.margin_top = box.text_frame.margin_bottom = 0
        box.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        font = ImageFont.truetype(str(FONT / ('Lato-Bold.ttf' if bold else 'Lato-Regular.ttf')), round(size*SCALE/72))
        wrapped = []
        for i, para in enumerate(value.split('\n')):
            p = box.text_frame.paragraphs[0] if i == 0 else box.text_frame.add_paragraph()
            p.text = para
            p.font.name = 'Lato'
            p.font.size = Pt(size)
            p.font.bold = bold
            p.font.color.rgb = RGBColor.from_string(color)
            p.line_spacing = 1.12
            p.space_after = Pt(0)
            p.alignment = PP_ALIGN.RIGHT if align == 'right' else PP_ALIGN.CENTER if align == 'center' else PP_ALIGN.LEFT
            p.font.language_id = 1046
            words = para.split()
            cur = ''
            for word in words:
                trial = (cur+' '+word).strip()
                if font.getlength(trial) <= w*SCALE or not cur:
                    cur = trial
                else:
                    wrapped.append(cur)
                    cur = word
            wrapped.append(cur)
        lineheight = size/72*SCALE*1.12
        if len(wrapped)*lineheight > h*SCALE + 7:
            layout_issues.append((self.index, value[:70], round(len(wrapped)*lineheight/SCALE,2), h))
        for i, row in enumerate(wrapped):
            dx = x*SCALE
            if align == 'right':
                dx += w*SCALE-font.getlength(row)
            if align == 'center':
                dx += (w*SCALE-font.getlength(row))/2
            self.draw.text((dx, y*SCALE+i*lineheight), row, fill='#'+color, font=font, anchor='lt')
        return box

    def notes(self, value):
        self.slide.notes_slide.notes_text_frame.text = value

    def finish(self):
        self.im.save(PREVIEWS / f'slide_{self.index:02d}.png')
        images.append(self.im)


# 01 — Proposta: um único argumento e três premissas de escala.
s = Slide(1, 'Projeto', True)
s.text(.65, 1.15, 7.8, .32, '01  /  SOLARIZAÇÃO DO CENTRO DE DISTRIBUIÇÃO', 11, GREEN, True)
s.text(.65, 1.77, 8.1, 1.35, 'Energia solar para\nreduzir o custo do CD', 36, WHITE, True)
s.text(.65, 3.32, 6.8, .95, 'Usina de 1 MWp no CD de Cajamar para reduzir a despesa elétrica e a exposição a reajustes tarifários.', 19, 'D2DFDF')
# Ilustração geométrica de telhado solar, construída com formas editáveis.
s.circle(10.82, 1.61, .65, AMBER)
for angle in range(0, 360, 45):
    a = angle*pi/180
    cx, cy = 11.145, 1.935
    s.line(cx+.43*cos(a),cy+.43*sin(a),cx+.57*cos(a),cy+.57*sin(a),AMBER,1.5)
s.rect(8.68, 3.25, 3.42, .88, NAVY2)
s.line(8.45, 3.23, 10.39, 2.57, '86ADA6', 2)
s.line(10.39, 2.57, 12.32, 3.23, '86ADA6', 2)
for row in range(2):
    for col in range(4):
        s.rect(8.90+col*.72, 2.98+row*.34, .62, .25, TEAL, border='8AC3B0')
s.rect(9.05, 3.66, .48, .47, '3D5C63')
s.rect(9.82, 3.66, .48, .47, '3D5C63')
s.rect(10.59, 3.66, 1.02, .47, '3D5C63')
s.line(8.45, 4.15, 12.32, 4.15, '86ADA6', 1.3)
for x, value, label in [(.65,'R$ 4,0 mi','INVESTIMENTO INICIAL'),(4.73,'1 MWp','POTÊNCIA INSTALADA'),(8.81,'15 anos','HORIZONTE DE ANÁLISE')]:
    s.rect(x, 4.91, 3.87, 1.33, NAVY2, radius=True)
    s.text(x+.22, 5.08, 3.43, .63, value, 33, GREEN, True)
    s.text(x+.22, 5.84, 3.43, .23, label, 10, 'D2DFDF')
s.text(.65, 6.49, 11.8, .23, 'Objetivo complementar: reduzir emissões e trazer previsibilidade ao custo logístico.', 11, 'B8CACB')
s.notes('0:00–0:25 | PROJETO\n\nEste é um caso fictício da LogiBrasil. A proposta é instalar uma usina solar de um megawatt-pico no telhado do centro de distribuição de Cajamar, para reduzir a conta de energia. O investimento é de quatro milhões de reais, com análise de quinze anos.')
s.finish()

# 02 — Estrutura financeira: fontes → custos após tributos → WACC.
s = Slide(2, 'Estrutura de capital')
s.text(.65, 1.16, 9, .28, '02  /  FINANCIAMENTO', 11, TEAL, True)
s.text(.65, 1.66, 11.9, .7, 'Do custo de cada fonte ao custo médio de capital', 30, INK, True)
s.text(.65, 2.48, 8.05, .32, 'COMPOSIÇÃO DO INVESTIMENTO  /  R$ 4,0 MILHÕES', 11, GRAY, True)
s.rect(.65, 3.04, 2.8525, .54, TEAL)
s.rect(3.5025, 3.04, 5.2975, .54, GREEN)
s.text(.84, 3.16, 2.4, .3, '35%  /  PRÓPRIO', 12, WHITE, True)
s.text(3.70, 3.16, 4.8, .3, '65%  /  DÍVIDA', 12, INK, True)

# Os dois custos ficam alinhados por fonte, fora do quadro do resultado.
for x,w,label in [(.65,2.20,'FONTE'),(3.0,.65,'PESO'),(3.87,1.30,'VALOR'),(5.46,1.40,'CUSTO ANTES\nDOS TRIBUTOS'),(7.22,1.58,'CUSTO APÓS\nOS TRIBUTOS')]:
    s.text(x,3.89,w,.48,label,10,GRAY,True)
s.line(.65,4.49,8.80,4.49,LINE,1)
for y,source,weight,amount,before,after in [(4.72,'Capital próprio','35%','R$ 1,4 mi','17,0%','17,0%'),(5.37,'Financiamento verde','65%','R$ 2,6 mi','12,8%','8,45%')]:
    s.text(.65,y,2.20,.5,source,14,INK,True)
    s.text(3.0,y,.65,.35,weight,14,INK)
    s.text(3.87,y,1.36,.35,amount,14,INK)
    s.text(5.46,y,1.40,.39,before,18,GRAY)
    s.text(7.22,y,1.58,.39,after,18,TEAL,True)
s.line(.65,5.99,8.80,5.99,LINE,1)
s.text(.65,6.17,8.15,.3,'WACC = 35% × 17,0% + 65% × 8,448% = 11,4412% a.a.',12,TEAL,True)

# O quadro verde apresenta somente o resultado e sua interpretação.
s.rect(9.25,2.69,3.43,3.64,GREEN,radius=True)
s.text(9.52,2.99,2.90,.28,'RESULTADO PONDERADO',10,INK,True)
s.text(9.47,3.65,2.99,.9,'11,4%',46,INK,True)
s.text(9.52,4.61,2.90,.3,'WACC AO ANO',12,INK,True)
s.line(9.52,5.10,12.40,5.10,'8FBB65',1)
s.text(9.52,5.34,2.89,.66,'Custo médio do capital\npara avaliar o projeto.',13,INK)
s.text(.65,6.61,12.03,.22,'Taxas anuais. Premissa: Lucro Real, juros dedutíveis e aproveitamento integral do benefício fiscal marginal de 34%.',9,GRAY)
s.notes('0:25–1:05 | ESTRUTURA DE CAPITAL\n\nUsaremos trinta e cinco por cento de capital próprio e sessenta e cinco por cento de dívida. O capital próprio custa dezessete por cento ao ano; a dívida, doze vírgula oito. Assumindo aproveitamento integral do benefício fiscal de trinta e quatro por cento, a dívida passa a custar oito vírgula quarenta e cinco por cento. Ponderando esses custos, chegamos ao WACC de onze vírgula quatro por cento ao ano.')
s.finish()

# 03 — Retorno: VPL, caixa, payback e comparação TIR/WACC.
s = Slide(3, 'Retorno esperado')
s.text(.65, 1.16, 9, .28, '03  /  CRIAÇÃO DE VALOR', 11, TEAL, True)
s.text(.65, 1.66, 12.0, .7, 'Retorno acima do custo de capital', 32, INK, True)
for x, value, label in [(.65,'R$ 1,65 mi','VPL POSITIVO'),(4.73,'R$ 805 mil','CAIXA LIVRE POR ANO'),(8.81,'5,0 anos','PAYBACK SIMPLES')]:
    s.rect(x, 2.63, 3.87, 1.35, WHITE, radius=True)
    s.text(x+.23, 2.85, 3.43, .58, value, 29, TEAL, True)
    s.text(x+.23, 3.57, 3.43, .24, label, 10, GRAY, True)
s.text(.65, 4.42, 7.1, .28, 'TIR DO PROJETO  ×  WACC', 11, GRAY, True)
maxw = 5.5
for y,label,rate,color in [(5.04,'TIR',IRR,TEAL),(5.82,'WACC',WACC,'93AAA4')]:
    s.text(.65,y+.04,1.0,.3,label,13,INK,True)
    s.rect(1.75,y,maxw*rate/.22,.38,color)
    s.text(1.94+maxw*rate/.22,y,1.1,.4,f'{rate*100:.1f}%'.replace('.',','),18,INK,True)
s.rect(8.75, 4.43, 3.93, 1.92, PALE, radius=True)
s.text(9.01, 4.64, 3.4, .63, '+7,1 p.p.', 33, TEAL, True)
s.text(9.02, 5.36, 3.3, .73, 'TIR acima do WACC\nCaixa / investimento: 20,1% a.a.', 13, INK)
s.text(.65, 6.52, 12.03, .23, 'Fluxos do projeto: −R$ 4,0 mi no início; +R$ 805 mil ao fim de cada ano, por 15 anos. Sem valor residual.', 10, GRAY)
s.notes('1:05–1:35 | RETORNO\n\nProjetamos oitocentos e cinco mil reais de caixa livre por ano. A TIR é de dezoito vírgula seis por cento, acima do WACC. O VPL é positivo em um milhão e seiscentos e cinquenta mil reais, com payback simples de cinco anos. Assim, o retorno esperado supera o custo de capital.')
s.finish()

# 04 — Decisão: condições e riscos, sem repetição dos detalhes financeiros.
s = Slide(4, 'Recomendação', True)
s.text(.65, 1.16, 9, .28, '04  /  DECISÃO DE INVESTIMENTO', 11, GREEN, True)
s.text(.65, 1.75, 11.8, 1.32, 'Aprovar o projeto, condicionado\nà validação das premissas', 34, WHITE, True)
s.text(.65, 3.30, 11.5, .45, 'VPL positivo de R$ 1,65 milhão e TIR superior ao custo de capital.', 20, GREEN)
cards = [(.65,'01','FINANCIAMENTO','Contratar a dívida dentro\ndo custo previsto.'),(4.73,'02','GERAÇÃO DE ENERGIA','Validar a produção estimada\ne a garantia de desempenho.'),(8.81,'03','BENEFÍCIO FISCAL','Confirmar dedutibilidade\ne capacidade de aproveitamento.')]
for x,num,label,desc in cards:
    s.rect(x,4.23,3.87,1.88,NAVY2,radius=True)
    s.text(x+.23,4.44,.6,.42,num,20,GREEN,True)
    s.text(x+.23,5.02,3.41,.29,label,11,WHITE,True)
    s.text(x+.23,5.47,3.41,.58,desc,13,'D2DFDF')
s.text(.65,6.49,12.03,.23,'Riscos centrais: geração abaixo da estimativa e financiamento acima da taxa prevista.',11,'B8CACB')
s.notes('1:35–1:55 | RECOMENDAÇÃO\n\nRecomendo aprovar o projeto, condicionado à confirmação do orçamento, da geração de energia, da capacidade do telhado e das condições de financiamento. Com essas premissas validadas, o investimento pode reduzir custos e criar valor para a empresa.\n\n[Encerrar até 1:55. Reservar os cinco segundos finais para eventuais atrasos. Não avançar ao apêndice durante o pitch.]')
s.finish()

# 05 — Apêndice técnico para perguntas, fora dos dois minutos.
s = Slide(5, 'Apêndice • premissas e cálculos')
s.text(.65,1.16,9,.28,'APÊNDICE  /  APOIO À DISCUSSÃO',11,TEAL,True)
s.text(.65,1.66,12,.7,'Premissas que sustentam os indicadores',31,INK,True)
s.text(.65,2.58,5.4,.3,'MODELO DO PROJETO',11,TEAL,True)
rows = [('Investimento no início','R$ 4.000.000'),('Caixa livre anual¹','R$ 805.000'),('Horizonte / momento do fluxo','15 anos / fim de cada ano'),('Valor residual','Zero'),('Taxa de desconto sem arredondar','11,4412% a.a.')]
for i,(label,value) in enumerate(rows):
    y=3.11+i*.48
    s.text(.65,y,3.7,.28,label,12,GRAY)
    s.text(4.22,y,2.14,.3,value,12,INK,True,align='right')
    s.line(.65,y+.37,6.36,y+.37,LINE,.7)
s.rect(6.84,2.58,5.84,3.40,WHITE,radius=True)
s.text(7.10,2.82,5.3,.28,'FÓRMULAS E DEFINIÇÕES',11,TEAL,True)
for y,label,definition in [(3.32,'WACC','35% × 17% + 65% × 12,8% × (1 − 34%)'),(3.94,'VPL','−4.000.000 + Σ [805.000 / (1 + WACC)ᵗ]'),(4.56,'TIR','Taxa que zera o VPL; t = 1 até 15'),(5.18,'Payback','4.000.000 / 805.000 ≈ 5,0 anos')]:
    s.text(7.10,y,1.05,.27,label,12,INK,True)
    s.text(8.19,y,4.17,.45,definition,12,GRAY)
s.text(.65,6.04,12.03,.40,'¹ Após tributos operacionais e reinvestimentos; antes dos juros e da amortização. O fluxo anual é uma premissa, sem abertura operacional no texto-base.',11,GRAY)
s.text(.65,6.57,12.03,.22,'Fonte: pitch_economia.md. Cálculos recalculados. Custos de capital e benefício fiscal são premissas acadêmicas.',9,GRAY)
s.notes('SLIDE DE APOIO — não faz parte dos dois minutos.\n\nEste apêndice permite reproduzir os cálculos. Os fluxos são nominais, constantes e anuais. O valor residual e as variações adicionais de capital de giro são considerados nulos. O investimento inicial deve ser entendido como todo o desembolso inicial do modelo. Não foram fornecidos dados de geração, tarifas de energia, degradação dos painéis, manutenção, depreciação ou orçamento técnico; os R$ 805 mil são assumidos como caixa livre após tributos operacionais e reinvestimentos. Sem essa premissa não é possível validar os retornos.\n\nO benefício fiscal de 34% é uma hipótese de alíquota marginal e aproveitamento integral no Lucro Real. Não confundir com alíquota efetiva, nem com dedução do principal da dívida. O ROIC contábil e a TIR do acionista não são apresentados por ausência de detalhamento suficiente.\n\nO indicador caixa/investimento de 20,1% preserva o conteúdo chamado de ROI aproximado no original, mas explicita sua definição.\n\nBase: pitch_economia.md; empresa, projeto, local e valores são fictícios.')
s.finish()

if layout_issues:
    raise RuntimeError('Possível transbordamento de texto: '+repr(layout_issues))

prs.save(OUT)
# Reabrir o pacote confirma a integridade da estrutura do PowerPoint.
check = Presentation(OUT)
assert len(check.slides) == 5
for sl in check.slides:
    assert sl.has_notes_slide
    for shape in sl.shapes:
        assert shape.left >= 0 and shape.top >= 0
        assert shape.left + shape.width <= prs.slide_width + 10000
        assert shape.top + shape.height <= prs.slide_height + 10000

thumbs = Image.new('RGB',(1600,1570),'#E1E7E3')
for i,im in enumerate(images):
    thumb=im.resize((768,432))
    thumbs.paste(thumb,(22+(i%2)*790,22+(i//2)*510))
thumbs.save(PREVIEWS/'visao_geral.png')
images[0].save(PDF_OUT,save_all=True,append_images=images[1:],resolution=120)
print(f'Gerado: {OUT}')
print(f'Prévia: {PDF_OUT}')
print(f'Slides: {len(check.slides)} | TIR: {IRR:.8%} | WACC: {WACC:.6%} | VPL: R$ {NPV:.2f}')
print(f'Prévia visual: {PREVIEWS / "visao_geral.png"}')
