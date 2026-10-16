"""Gera data/planilha-base-exemplo.xlsx com dados FICTÍCIOS (empresas e valores inventados)
e tests/fixtures/seed.js usado pelos testes. Requer: pip install openpyxl"""
import json, os, datetime as dt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
D = lambda s: dt.datetime.strptime(s, '%Y-%m-%d') if s else None

HEAD = ['id', 'parceiro', 'tipo_servico', 'ferramenta', 'finalidade', 'vigencia_inicio', 'vigencia_fim',
        'indice_reajuste', 'mes_reajuste', 'sla', 'renovacao_tipo', 'valor_suporte', 'valor_orcado',
        'valor_pago_anterior', 'observacoes', 'gestor_nome', 'gestor_email', 'status_renovacao',
        'atualizado_em', 'atualizado_por']
# parceiro, tipo, ferramenta, finalidade, início, fim, mês reajuste, valor suporte, orçado, pago, observações
ROWS = [
    ('Orion Secure', 'Suporte 7x24', 'Antivírus de Servidores', 'Proteção básica de servidores de arquivos', '2024-03-01', '2026-02-28', '', '12 × R$ 4.000,00', 50000.00, 48000.00, 'Substituição planejada pela nova plataforma de endpoint.'),
    ('Atlas Backup', 'Suporte 7x24', 'Backup Imutável', 'Cópias imutáveis contra ransomware', '2025-04-01', '2028-03-31', 'abril', '36 × R$ 15.000,00', 540000.00, 500000.00, 'Inclui testes trimestrais de restauração.'),
    ('Pulsar Service', 'Service Desk N1 7x24', 'Service Desk N1', 'Atendimento e triagem de chamados', '2025-01-01', '2027-12-31', 'janeiro', 'R$ 40.000,00 / mês', 480000.00, 450000.00, 'Atendimento N1 de requisições de segurança.'),
    ('Nebula Logs', 'Suporte 5x8', 'Plataforma de Logs', 'Coleta e retenção de logs', '2025-06-01', '2028-05-31', 'junho', '', 400000.00, 380000.00, 'Suporte à plataforma e apoio em casos de uso.'),
    ('Gatekeeper Net', 'Suporte 7x24', 'NAC', 'Controle de acesso à rede', '2023-09-01', '2026-08-31', '', '36 × R$ 5.000,00', 60000.00, 60000.00, 'Renovação em negociação.'),
    ('Vértice Web', 'Suporte 7x24', 'Proxy Web', 'Filtro de navegação', '2025-03-01', '2027-08-31', '', 'R$ 200.000,00 / ano', 200000.00, 190000.00, 'Suporte para regras e troubleshooting.'),
    ('Radar Sec', 'Suporte 5x8', 'Scanner de Vulnerabilidades', 'Varredura de ativos', '2025-02-01', '2028-01-31', 'fevereiro', '', 150000.00, 140000.00, 'Licença por ativo.'),
    ('Identix', 'Suporte 7x24', 'Gestão de Identidades', 'Ciclo de vida de identidades', '2025-08-01', '2028-07-31', '', '36 × R$ 10.000,00', 120000.00, 110000.00, 'Integração com RH e diretório.'),
    ('Farol Intel', 'Suporte 7x24', 'Inteligência de Ameaças', 'Feeds e alertas de ameaças', '2025-05-01', '2026-04-30', '', '12 × R$ 10.000,00', 120000.00, 115000.00, 'Avaliar consolidação com outro fornecedor.'),
    ('Alerta Humano', 'Suporte 5x8', 'Conscientização', 'Treinamentos e simulações de phishing', '2025-09-01', '2026-08-31', '', '', 80000.00, 75000.00, 'Campanhas trimestrais.'),
    ('Escudo Net', 'MSS/Suporte 7x24', 'Anti-DDoS', 'Mitigação de ataques volumétricos', '2025-07-01', '2027-06-30', '', '24 × R$ 12.500,00', 150000.00, 140000.00, 'Proteção dos links de internet.'),
    ('Red Lab', 'Projeto', 'Pentest', 'Teste de intrusão anual', '2025-06-15', '2026-06-14', '', '', 90000.00, 85000.00, 'Escopo: aplicações externas.'),
    ('Muralha Dados', 'Suporte 5x8', 'DLP', 'Prevenção contra vazamento de dados', '2025-12-10', '2026-12-10', '', 'R$ 25.000,00', 250000.00, 230000.00, 'Manutenção de políticas e suporte.'),
    ('Chaveiro', 'Suporte 7x24', 'Cofre de Senhas', 'Credenciais privilegiadas', '2025-04-15', '2028-04-14', '', '36 × R$ 8.000,00', 96000.00, 90000.00, 'Manutenção e atualização do cofre.'),
    ('Ponte API', 'Suporte 7x24', 'Proteção de API', 'Segurança de APIs expostas', '', '', '', '', 300000.00, 200000.00, 'Vigência a confirmar com Compras.'),
    ('Móvel Seguro', '', 'MDM', '', '2025-07-01', '2028-06-30', '', '', None, None, ''),
]
contratos = []
for i, r in enumerate(ROWS, 1):
    p, t, f, fin, ini, fim, mes, vs, orc, pago, obs = r
    contratos.append([f'CT-{i:03d}', p, t, f, fin, D(ini), D(fim), 'IPCA', mes, 'Plano de Governança' if t else '',
                      'Renovação' if t else '', vs, orc, pago, obs, '', '', 'Em vigor', None, 'importação inicial'])

CONFIG = [
    ['JANELAS', '90,60,30', 'Dias antes do fim da vigência em que o alerta é enviado (separados por vírgula)'],
    ['NOTIFICAR_ADMINS', 'SIM', 'SIM/NAO — copiar os administradores do portal nos alertas'],
    ['CC_FIXO', '', 'E-mails sempre copiados nos alertas (lista de distribuição), separados por vírgula'],
    ['REENVIO_VENCIDOS', 'SIM', 'SIM/NAO — reenviar aviso semanal enquanto o contrato estiver vencido'],
    ['DIA_REENVIO', '1', 'Dia da semana do reenvio de vencidos (1=segunda … 7=domingo)'],
    ['HORA_GATILHO', '7', 'Hora do gatilho diário de alertas (0–23, America/Sao_Paulo)'],
    ['AUTO_REFRESH_SEGUNDOS', '60', 'Intervalo (s) em que o portal verifica se os dados mudaram (mín. 15)'],
    ['LOGO_FILE_ID', '', 'Opcional: ID de um arquivo no Drive (versão para fundo escuro) com o logo da organização'],
    ['ROTULO_ORCADO', 'Orçado 2026', 'Rótulo da coluna valor_orcado'],
    ['ROTULO_PAGO', 'Pago 2025', 'Rótulo da coluna valor_pago_anterior'],
    ['DOMINIO_PERMITIDO', 'example.com', 'Somente e-mails deste domínio podem solicitar acesso (vazio = qualquer)'],
]
SHEETS = {
    'Contratos': (HEAD, contratos),
    'Usuarios': (['email', 'nome', 'perfil', 'escopo', 'status', 'criado_em', 'aprovado_por', 'ultimo_acesso', 'tema'], []),
    'Solicitacoes': (['id', 'email', 'nome', 'area', 'perfil_solicitado', 'justificativa', 'status', 'criada_em', 'decidido_por', 'decidido_em', 'motivo'], []),
    'LogAlertas': (['chave', 'contrato_id', 'parceiro', 'ferramenta', 'janela', 'vigencia_fim', 'enviado_em', 'destinatarios', 'status', 'erro'], []),
    'Auditoria': (['quando', 'usuario', 'acao', 'detalhe'], []),
    'Config': (['chave', 'valor', 'descricao'], CONFIG),
}

wb = openpyxl.Workbook(); wb.remove(wb.active)
seed = {}
for name, (head, rows) in SHEETS.items():
    ws = wb.create_sheet(name); ws.append(head)
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='13241E')
    for r in rows: ws.append(r)
    ws.freeze_panes = 'A2'
    for i, h in enumerate(head, 1): ws.column_dimensions[get_column_letter(i)].width = max(14, len(h) + 4)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, dt.datetime): c.number_format = 'dd/mm/yyyy'
            elif isinstance(c.value, float): c.number_format = '#,##0.00'
    seed[name] = [head] + [[{'__d': [v.year, v.month, v.day]} if isinstance(v, dt.datetime) else ('' if v is None else v) for v in r] for r in rows]
os.makedirs(os.path.join(ROOT, 'data'), exist_ok=True)
wb.save(os.path.join(ROOT, 'data', 'planilha-base-exemplo.xlsx'))
with open(os.path.join(ROOT, 'tests', 'fixtures', 'seed.js'), 'w', encoding='utf-8') as fh:
    fh.write('// Gerado por scripts/gerar_planilha_exemplo.py — dados fictícios\nwindow.__SEED = ' + json.dumps(seed, ensure_ascii=False) + ';\n')
print(len(contratos), 'contratos fictícios')
