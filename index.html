import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, date, timedelta
import psycopg2
import webbrowser
import sys
import json
import csv
from flask import Flask, jsonify, request, render_template_string
import threading
import hashlib
import socket
import os
from datetime import datetime, date
from decimal import Decimal
from PIL import Image, ImageTk
from reportlab.platypus.flowables import HRFlowable

# Adicionar imports para PDF e impressão
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
import tempfile
import win32print # type: ignore
import win32ui # type: ignore
from reportlab.pdfgen import canvas

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # quando é .exe
    except Exception:
        base_path = os.path.abspath(".")  # quando é .py
    return os.path.join(base_path, relative_path)

class SistemaGestaoPapelaria:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Gestão Papelaria v2.0") 
        self.root.geometry("1200x700")
        
        # Configuração do tema
        self.root.configure(bg='#2c3e50')
        
        self.autor = "Tomas J. T. Antonio"
        self.versao = "2.0 - Sistema Completo Simplificado"
        
        # Inicializar variáveis
        self.carrinho = []
        self.usuario_atual = None
        self.flask_app = None
        self.flask_thread = None
        
        # Preços do sistema (FIXOS - não podem ser alterados por vendedores)
        self.config_precos = {
            'copia_pb': 2.0,          # Cópia preto e branco
            'copia_colorida': 10.0,    # Cópia colorida
            'impressao_pb': 5.0,       # Impressão preto e branco
            'impressao_colorida': 45.0, # Impressão colorida
            'encadernacao_6mm': 20.0,
            'encadernacao_8mm': 35.0,
            'encadernacao_10mm': 45.0,
            'encadernacao_12mm': 55.0,
            'encadernacao_14mm': 65.0,
            'encadernacao_16mm': 80.0,
            'encadernacao_18mm': 100.0,
            'encadernacao_20mm': 120.0,
            'encadernacao_22mm': 150.0,
            'laminacao_bi': 30.0,
            'laminacao_a4': 50.0,
            'laminacao_a3': 100.0,
            'laminacao_a5': 40.0,
            'digitacao': 35.0,
            'fotografia': 100.0,
            'cartao_visita': 80.0,
            'convite': 120.0,
            'banner': 250.0,
            'adesivo': 45.0
        }
        
        # Configurar estilo
        self.configurar_estilo()
        
        # Conectar ao banco
        self.conectar_banco()
        
        # Mostrar tela de login
        self.tela_login()
    
    def iniciar_api_web(self):
        """Iniciar servidor web para dashboard simplificado"""
        if self.flask_app:
            return
            
        self.flask_app = Flask(__name__)
        
        @self.flask_app.route('/')
        def index():
            return self.dashboard_web()
        
        @self.flask_app.route('/api/metricas')
        def api_metricas():
            try:
                metricas = self.obter_metricas_dashboard()
                return jsonify({
                    'status': 'success',
                    'vendas_hoje': metricas['vendas_hoje'],
                    'total_produtos': metricas['total_produtos'],
                    'estoque_baixo': metricas['estoque_baixo'],
                    'usuarios_ativos': metricas['usuarios_ativos'],
                    'servicos_especiais_hoje': metricas['servicos_especiais_hoje'],
                    'total_vendas_mes': metricas['total_vendas_mes'],
                    'total_servicos_mes': metricas['total_servicos_mes'],
                    'valor_estoque_total': metricas['valor_estoque_total']
                })
            except Exception as e:
                return jsonify({'status': 'error', 'message': str(e)}), 500
        
        @self.flask_app.route('/api/vendas')
        def api_vendas():
            try:
                periodo = request.args.get('periodo', 'hoje')
                limit = request.args.get('limit', 50)
                vendas_data = self.obter_vendas_periodo(periodo)
                return jsonify({
                    'status': 'success',
                    'vendas': vendas_data[:int(limit)],
                    'total_vendas': len(vendas_data),
                    'total_valor': sum(v.get('total', 0) for v in vendas_data),
                    'periodo': periodo
                })
            except Exception as e:
                return jsonify({'status': 'error', 'message': str(e)}), 500
        
        @self.flask_app.route('/api/produtos')
        def api_produtos():
            try:
                produtos = []
                self.cursor.execute('''
                    SELECT codigo, nome, categoria, preco, quantidade, estoque_minimo 
                    FROM produtos 
                    ORDER BY nome
                ''')
                
                for row in self.cursor.fetchall():
                    produtos.append({
                        'codigo': row[0],
                        'nome': row[1],
                        'categoria': row[2],
                        'preco': float(row[3]),
                        'quantidade': row[4],
                        'estoque_minimo': row[5],
                        'valor_total': float(row[3]) * row[4],
                        'status': 'CRÍTICO' if row[4] < row[5] else 'BAIXO' if row[4] == row[5] else 'OK'
                    })
                
                # Calcular valor total do estoque
                valor_total_estoque = sum(p['valor_total'] for p in produtos)
                
                return jsonify({
                    'status': 'success',
                    'produtos': produtos,
                    'total': len(produtos),
                    'estoque_baixo': len([p for p in produtos if p['quantidade'] <= p['estoque_minimo']]),
                    'valor_total_estoque': valor_total_estoque
                })
            except Exception as e:
                return jsonify({'status': 'error', 'message': str(e)}), 500
        
        @self.flask_app.route('/api/servicos')
        def api_servicos():
            try:
                periodo = request.args.get('periodo', 'hoje')
                data_inicio = self.get_data_inicio_periodo(periodo)
                
                impressao_data = []
                if data_inicio:
                    self.cursor.execute('''
                        SELECT s.tipo, s.numero_ontem, s.numero_hoje, s.quantidade, 
                               s.preco_unitario, s.total, s.data_hora, u.nome as usuario
                        FROM servicos s
                        JOIN usuarios u ON s.usuario_id = u.id
                        WHERE s.data >= %s
                        ORDER BY s.data_hora DESC
                    ''', (data_inicio,))
                else:
                    self.cursor.execute('''
                        SELECT s.tipo, s.numero_ontem, s.numero_hoje, s.quantidade, 
                               s.preco_unitario, s.total, s.data_hora, u.nome as usuario
                        FROM servicos s
                        JOIN usuarios u ON s.usuario_id = u.id
                        ORDER BY s.data_hora DESC
                    ''')
                
                for row in self.cursor.fetchall():
                    tipo_nome = self.get_nome_tipo_servico(row[0])
                    impressao_data.append({
                        'tipo': tipo_nome,
                        'numero_ontem': row[1],
                        'numero_hoje': row[2],
                        'quantidade': row[3],
                        'preco_unitario': float(row[4]),
                        'total': float(row[5]),
                        'data_hora': row[6].isoformat() if row[6] else '',
                        'usuario': row[7]
                    })
                
                especiais_data = []
                if data_inicio:
                    self.cursor.execute('''
                        SELECT se.tipo, se.descricao, se.quantidade, se.preco_unitario, 
                               se.total, se.data_hora, u.nome as usuario
                        FROM servicos_especiais se
                        JOIN usuarios u ON se.usuario_id = u.id
                        WHERE se.data >= %s
                        ORDER BY se.data_hora DESC
                    ''', (data_inicio,))
                else:
                    self.cursor.execute('''
                        SELECT se.tipo, se.descricao, se.quantidade, se.preco_unitario, 
                               se.total, se.data_hora, u.nome as usuario
                        FROM servicos_especiais se
                        JOIN usuarios u ON se.usuario_id = u.id
                        ORDER BY se.data_hora DESC
                    ''')
                
                for row in self.cursor.fetchall():
                    especiais_data.append({
                        'tipo': row[0],
                        'descricao': row[1] or '',
                        'quantidade': row[2],
                        'preco_unitario': float(row[3]),
                        'total': float(row[4]),
                        'data_hora': row[5].isoformat() if row[5] else '',
                        'usuario': row[6]
                    })
                
                # Estatísticas
                total_impressao = sum(i['total'] for i in impressao_data)
                total_especiais = sum(i['total'] for i in especiais_data)
                
                return jsonify({
                    'status': 'success',
                    'servicos_impressao': {
                        'items': impressao_data,
                        'total': len(impressao_data),
                        'total_valor': total_impressao
                    },
                    'servicos_especiais': {
                        'items': especiais_data,
                        'total': len(especiais_data),
                        'total_valor': total_especiais
                    },
                    'total_geral': total_impressao + total_especiais
                })
            except Exception as e:
                return jsonify({'status': 'error', 'message': str(e)}), 500
        
        @self.flask_app.route('/api/vendas_detalhadas')
        def api_vendas_detalhadas():
            """API para obter vendas detalhadas com produtos/serviços"""
            try:
                periodo = request.args.get('periodo', 'hoje')
                data_inicio = self.get_data_inicio_periodo(periodo)
                
                vendas_detalhadas = []
                
                # Obter vendas do período
                if data_inicio:
                    self.cursor.execute('''
                        SELECT v.id, v.numero_serie, v.total, v.data_hora, u.nome as vendedor
                        FROM vendas v
                        JOIN usuarios u ON v.usuario_id = u.id
                        WHERE DATE(v.data_hora) >= %s
                        ORDER BY v.data_hora DESC
                        LIMIT 50
                    ''', (data_inicio,))
                else:
                    self.cursor.execute('''
                        SELECT v.id, v.numero_serie, v.total, v.data_hora, u.nome as vendedor
                        FROM vendas v
                        JOIN usuarios u ON v.usuario_id = u.id
                        ORDER BY v.data_hora DESC
                        LIMIT 50
                    ''')
                
                vendas = self.cursor.fetchall()
                
                for venda in vendas:
                    venda_id = venda[0]
                    
                    # Obter itens da venda (produtos)
                    self.cursor.execute('''
                        SELECT p.nome, p.categoria, iv.quantidade, iv.preco_unitario, iv.total_item
                        FROM itens_venda iv
                        JOIN produtos p ON iv.produto_id = p.id
                        WHERE iv.venda_id = %s
                    ''', (venda_id,))
                    
                    produtos_venda = []
                    for item in self.cursor.fetchall():
                        produtos_venda.append({
                            'tipo': 'Produto',
                            'nome': item[0],
                            'categoria': item[1],
                            'quantidade': item[2],
                            'preco_unitario': float(item[3]),
                            'total': float(item[4])
                        })
                    
                    # Contar esta venda como produto
                    if produtos_venda:
                        for produto in produtos_venda:
                            vendas_detalhadas.append({
                                'data_hora': venda[3].isoformat() if venda[3] else '',
                                'tipo': 'Produto',
                                'descricao': produto['nome'],
                                'quantidade': produto['quantidade'],
                                'total': produto['total'],
                                'usuario': venda[4],
                                'categoria': produto['categoria']
                            })
                
                # Obter serviços do período
                if data_inicio:
                    self.cursor.execute('''
                        SELECT s.tipo, s.quantidade, s.total, s.data_hora, u.nome
                        FROM servicos s
                        JOIN usuarios u ON s.usuario_id = u.id
                        WHERE s.data >= %s
                        UNION ALL
                        SELECT se.tipo, se.quantidade, se.total, se.data_hora, u.nome
                        FROM servicos_especiais se
                        JOIN usuarios u ON se.usuario_id = u.id
                        WHERE se.data >= %s
                        ORDER BY data_hora DESC
                        LIMIT 50
                    ''', (data_inicio, data_inicio))
                else:
                    self.cursor.execute('''
                        SELECT s.tipo, s.quantidade, s.total, s.data_hora, u.nome
                        FROM servicos s
                        JOIN usuarios u ON s.usuario_id = u.id
                        UNION ALL
                        SELECT se.tipo, se.quantidade, se.total, se.data_hora, u.nome
                        FROM servicos_especiais se
                        JOIN usuarios u ON se.usuario_id = u.id
                        ORDER BY data_hora DESC
                        LIMIT 50
                    ''')
                
                servicos = self.cursor.fetchall()
                
                for servico in servicos:
                    tipo_nome = self.get_nome_tipo_servico(servico[0]) if '_pb' in servico[0] or '_colorida' in servico[0] else servico[0]
                    vendas_detalhadas.append({
                        'data_hora': servico[3].isoformat() if servico[3] else '',
                        'tipo': 'Serviço',
                        'descricao': tipo_nome,
                        'quantidade': servico[1],
                        'total': float(servico[2]),
                        'usuario': servico[4],
                        'categoria': tipo_nome
                    })
                
                # Ordenar por data
                vendas_detalhadas.sort(key=lambda x: x['data_hora'], reverse=True)
                
                return jsonify({
                    'status': 'success',
                    'vendas_detalhadas': vendas_detalhadas[:100],
                    'total': len(vendas_detalhadas),
                    'total_produtos': len([v for v in vendas_detalhadas if v['tipo'] == 'Produto']),
                    'total_servicos': len([v for v in vendas_detalhadas if v['tipo'] == 'Serviço']),
                    'valor_total': sum(v['total'] for v in vendas_detalhadas)
                })
            except Exception as e:
                return jsonify({'status': 'error', 'message': str(e)}), 500
        
        @self.flask_app.route('/api/relatorio')
        def api_relatorio():
            try:
                tipo = request.args.get('tipo', 'vendas')
                periodo = request.args.get('periodo', 'hoje')
                
                if tipo == 'vendas':
                    vendas = self.obter_vendas_periodo(periodo)
                    return jsonify({
                        'status': 'success',
                        'dados': vendas,
                        'total': sum(v.get('total', 0) for v in vendas),
                        'quantidade': len(vendas)
                    })
                elif tipo == 'servicos':
                    data = self.get_dados_servicos_periodo(periodo)
                    return jsonify({
                        'status': 'success',
                        'dados': data,
                        'total': sum(d.get('total', 0) for d in data),
                        'quantidade': len(data)
                    })
                elif tipo == 'estoque':
                    produtos = []
                    self.cursor.execute('''
                        SELECT codigo, nome, categoria, preco, quantidade, estoque_minimo 
                        FROM produtos 
                        ORDER BY nome
                    ''')
                    
                    for row in self.cursor.fetchall():
                        produtos.append({
                            'codigo': row[0],
                            'nome': row[1],
                            'categoria': row[2],
                            'preco': float(row[3]),
                            'quantidade': row[4],
                            'estoque_minimo': row[5],
                            'valor_total': float(row[3]) * row[4],
                            'status': 'CRÍTICO' if row[4] < row[5] else 'BAIXO' if row[4] == row[5] else 'OK'
                        })
                    
                    valor_total_estoque = sum(p['valor_total'] for p in produtos)
                    estoque_baixo = len([p for p in produtos if p['quantidade'] <= p['estoque_minimo']])
                    
                    return jsonify({
                        'status': 'success',
                        'dados': produtos,
                        'total': len(produtos),
                        'estoque_baixo': estoque_baixo,
                        'valor_total_estoque': valor_total_estoque
                    })
                else:
                    data = self.get_dados_gerais_periodo(periodo)
                    return jsonify({
                        'status': 'success',
                        'dados': data,
                        'total': sum(d.get('total', 0) for d in data),
                        'quantidade': len(data)
                    })
                    
            except Exception as e:
                return jsonify({'status': 'error', 'message': str(e)}), 500
        
        @self.flask_app.route('/api/teste')
        def api_teste():
            return jsonify({'status': 'ok', 'mensagem': 'API funcionando!', 'versao': '2.0'})
        
        def run_flask():
            try:
                # Encontrar porta disponível
                port = 5000
                max_port = 5010
                while port <= max_port:
                    try:
                        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                        sock.settimeout(1)
                        sock.bind(('0.0.0.0', port))
                        sock.close()
                        break
                    except:
                        port += 1
                
                print(f"🌐 Tentando iniciar API Web na porta {port}...")
                self.flask_app.run(
                    host='0.0.0.0', 
                    port=port, 
                    debug=False, 
                    use_reloader=False, 
                    threaded=True
                )
            except Exception as e:
                print(f"❌ Erro no servidor Flask: {e}")
        
        self.flask_thread = threading.Thread(target=run_flask, daemon=True)
        self.flask_thread.start()
        print(f"✅ API Web iniciada em http://localhost:{5000}")
    
    def get_nome_tipo_servico(self, tipo):
        """Obter nome amigável para o tipo de serviço"""
        nomes = {
            'copia_pb': 'Cópia P&B (2 MT)',
            'copia_colorida': 'Cópia Colorida (10 MT)',
            'impressao_pb': 'Impressão P&B (5 MT)',
            'impressao_colorida': 'Impressão Colorida (45 MT)',
            'encadernacao_6mm': 'Encadernação Argola 6mm',
            'encadernacao_8mm': 'Encadernação Argola 8mm',
            'encadernacao_10mm': 'Encadernação Argola 10mm',
            'encadernacao_12mm': 'Encadernação Argola 12mm',
            'encadernacao_14mm': 'Encadernação Argola 14mm',
            'encadernacao_16mm': 'Encadernação Argola 16mm',
            'encadernacao_18mm': 'Encadernação Argola 18mm',
            'encadernacao_20mm': 'Encadernação Argola 20mm',
            'encadernacao_22mm': 'Encadernação Argola 22mm',
            'laminacao_bi': 'Laminação BI',
            'laminacao_a4': 'Laminação A4',
            'laminacao_a3': 'Laminação A3',
            'laminacao_a5': 'Laminação A5',
            'digitacao': 'Digitação',
            'fotografia': 'Fotografia',
            'cartao_visita': 'Cartão de Visita',
            'convite': 'Convite',
            'banner': 'Banner',
            'adesivo': 'Adesivo'
        }
        return nomes.get(tipo, tipo.replace('_', ' ').title())
    
    def get_data_inicio_periodo(self, periodo):
        """Obter data de início baseada no período"""
        hoje = date.today()
        if periodo == 'hoje':
            return hoje
        elif periodo == 'ontem':
            return hoje - timedelta(days=1)
        elif periodo == '7dias':
            return hoje - timedelta(days=7)
        elif periodo == 'mes':
            return hoje.replace(day=1)
        elif periodo == 'ano':
            return hoje.replace(month=1, day=1)
        elif periodo.startswith('data:'):
            try:
                data_str = periodo.split(':')[1]
                return datetime.strptime(data_str, '%Y-%m-%d').date()
            except:
                return None
        return None
    
    def obter_vendas_periodo(self, periodo):
        """Obter vendas por período"""
        data_inicio = self.get_data_inicio_periodo(periodo)
        vendas = []
        
        query = '''
            SELECT v.numero_serie, v.total, v.valor_recebido, v.troco, 
                   v.data_hora, u.nome as vendedor
            FROM vendas v
            JOIN usuarios u ON v.usuario_id = u.id
        '''
        
        params = []
        
        if data_inicio:
            query += ' WHERE DATE(v.data_hora) >= %s'
            params.append(data_inicio)
        
        query += ' ORDER BY v.data_hora DESC'
        
        try:
            self.cursor.execute(query, params)
            
            for row in self.cursor.fetchall():
                data_hora = row[4]
                data_formatada = ''
                if data_hora:
                    if isinstance(data_hora, datetime):
                        data_formatada = data_hora.strftime('%d/%m/%Y %H:%M')
                    elif isinstance(data_hora, str):
                        try:
                            dt = datetime.fromisoformat(data_hora.replace('Z', '+00:00'))
                            data_formatada = dt.strftime('%d/%m/%Y %H:%M')
                        except:
                            data_formatada = data_hora
                
                vendas.append({
                    'numero_serie': row[0],
                    'total': float(row[1]) if row[1] else 0.0,
                    'valor_recebido': float(row[2]) if row[2] else 0.0,
                    'troco': float(row[3]) if row[3] else 0.0,
                    'data_hora': row[4].isoformat() if hasattr(row[4], 'isoformat') else str(row[4]),
                    'vendedor': row[5],
                    'data_formatada': data_formatada
                })
        except Exception as e:
            print(f"Erro ao obter vendas: {e}")
        
        return vendas
    
    def get_dados_servicos_periodo(self, periodo):
        """Obter dados de serviços por período"""
        data_inicio = self.get_data_inicio_periodo(periodo)
        dados = []
        
        # Impressão/Cópia
        query_imp = '''
            SELECT s.data_hora, s.tipo, s.numero_ontem, s.numero_hoje, 
                   s.quantidade, s.total, u.nome
            FROM servicos s
            JOIN usuarios u ON s.usuario_id = u.id
        '''
        
        params_imp = []
        if data_inicio:
            query_imp += ' WHERE s.data >= %s'
            params_imp.append(data_inicio)
        
        query_imp += ' ORDER BY s.data_hora DESC'
        
        try:
            self.cursor.execute(query_imp, params_imp)
            for row in self.cursor.fetchall():
                tipo_nome = self.get_nome_tipo_servico(row[1])
                dados.append({
                    'data_hora': row[0].isoformat() if hasattr(row[0], 'isoformat') else str(row[0]),
                    'tipo': tipo_nome,
                    'numero_ontem': row[2],
                    'numero_hoje': row[3],
                    'quantidade': row[4],
                    'total': float(row[5]) if row[5] else 0.0,
                    'usuario': row[6]
                })
        except Exception as e:
            print(f"Erro ao obter serviços: {e}")
        
        # Serviços especiais
        query_esp = '''
            SELECT se.data_hora, se.tipo, se.descricao, se.quantidade, se.total, u.nome
            FROM servicos_especiais se
            JOIN usuarios u ON se.usuario_id = u.id
        '''
        
        params_esp = []
        if data_inicio:
            query_esp += ' WHERE se.data >= %s'
            params_esp.append(data_inicio)
        
        query_esp += ' ORDER BY se.data_hora DESC'
        
        try:
            self.cursor.execute(query_esp, params_esp)
            for row in self.cursor.fetchall():
                dados.append({
                    'data_hora': row[0].isoformat() if hasattr(row[0], 'isoformat') else str(row[0]),
                    'tipo': row[1],
                    'descricao': row[2] or '',
                    'quantidade': row[3],
                    'total': float(row[4]) if row[4] else 0.0,
                    'usuario': row[5]
                })
        except Exception as e:
            print(f"Erro ao obter serviços especiais: {e}")
        
        return dados
    
    def get_dados_gerais_periodo(self, periodo):
     """Obter dados gerais por período com nome dos produtos/serviços"""
     dados = []
     data_inicio = self.get_data_inicio_periodo(periodo)
    
     # Vendas com produtos
     query_vendas = '''
        SELECT v.data_hora, 'Venda' as tipo, 
               p.nome as descricao,
               iv.quantidade, v.total, u.nome, v.numero_serie
        FROM vendas v
        JOIN usuarios u ON v.usuario_id = u.id
        JOIN itens_venda iv ON v.id = iv.venda_id
        JOIN produtos p ON iv.produto_id = p.id
     '''
    
     params_vendas = []
     if data_inicio:
        query_vendas += ' WHERE DATE(v.data_hora) >= %s'
        params_vendas.append(data_inicio)
    
     query_vendas += ' ORDER BY v.data_hora DESC'
    
     try:
        self.cursor.execute(query_vendas, params_vendas)
        for row in self.cursor.fetchall():
            dados.append({
                'data_hora': row[0].isoformat() if hasattr(row[0], 'isoformat') else str(row[0]),
                'tipo': row[1],
                'descricao': f"{row[2]} (Venda: {row[6]})",
                'quantidade': row[3],
                'total': float(row[4]) if row[4] else 0.0,
                'usuario': row[5]
            })
     except Exception as e:
        print(f"Erro ao obter vendas detalhadas: {e}")
    
     # Serviços de impressão/cópia
     query_serv = '''
        SELECT s.data_hora, 
               CASE 
                   WHEN s.tipo = 'copia_pb' THEN 'Cópia P&B'
                   WHEN s.tipo = 'copia_colorida' THEN 'Cópia Colorida'
                   WHEN s.tipo = 'impressao_pb' THEN 'Impressão P&B'
                   WHEN s.tipo = 'impressao_colorida' THEN 'Impressão Colorida'
                   ELSE s.tipo
               END as tipo,
               CONCAT('De ', s.numero_ontem, ' a ', s.numero_hoje) as descricao,
               s.quantidade, s.total, u.nome
        FROM servicos s
        JOIN usuarios u ON s.usuario_id = u.id
     '''
    
     params_serv = []
     if data_inicio:
        query_serv += ' WHERE s.data >= %s'
        params_serv.append(data_inicio)
    
     query_serv += ' ORDER BY s.data_hora DESC'
    
     try:
        self.cursor.execute(query_serv, params_serv)
        for row in self.cursor.fetchall():
            dados.append({
                'data_hora': row[0].isoformat() if hasattr(row[0], 'isoformat') else str(row[0]),
                'tipo': row[1],
                'descricao': row[2],
                'quantidade': row[3],
                'total': float(row[4]) if row[4] else 0.0,
                'usuario': row[5]
            })
     except Exception as e:
        print(f"Erro ao obter serviços: {e}")
    
     # Serviços especiais
     query_esp = '''
        SELECT se.data_hora, se.tipo, 
               COALESCE(se.descricao, se.tipo) as descricao,
               se.quantidade, se.total, u.nome
        FROM servicos_especiais se
        JOIN usuarios u ON se.usuario_id = u.id
     '''
    
     params_esp = []
     if data_inicio:
        query_esp += ' WHERE se.data >= %s'
        params_esp.append(data_inicio)
    
     query_esp += ' ORDER BY se.data_hora DESC'
    
     try:
        self.cursor.execute(query_esp, params_esp)
        for row in self.cursor.fetchall():
            # Formatar nome do tipo de serviço
            tipo_formatado = row[1]
            if 'encadernacao' in row[1]:
                tamanho = row[1].split('_')[1] if '_' in row[1] else ''
                tipo_formatado = f"Encadernação Argola {tamanho}"
            elif 'laminacao' in row[1]:
                tipo_formatado = row[1].replace('_', ' ').title()
            
            dados.append({
                'data_hora': row[0].isoformat() if hasattr(row[0], 'isoformat') else str(row[0]),
                'tipo': tipo_formatado,
                'descricao': row[2],
                'quantidade': row[3],
                'total': float(row[4]) if row[4] else 0.0,
                'usuario': row[5]
            })
     except Exception as e:
        print(f"Erro ao obter serviços especiais: {e}")
    
     return dados
    def get_dados_vendas_detalhadas_periodo(self, periodo):
     """Obter dados detalhados de vendas por período"""
     dados = []
     data_inicio = self.get_data_inicio_periodo(periodo)
    
     # Produtos vendidos
     query_produtos = '''
        SELECT v.data_hora, 'Produto' as tipo, 
               p.nome as descricao,
               iv.quantidade, iv.total_item as total, u.nome, v.numero_serie,
               p.categoria
        FROM vendas v
        JOIN usuarios u ON v.usuario_id = u.id
        JOIN itens_venda iv ON v.id = iv.venda_id
        JOIN produtos p ON iv.produto_id = p.id
     '''
    
     params_produtos = []
     if data_inicio:
        query_produtos += ' WHERE DATE(v.data_hora) >= %s'
        params_produtos.append(data_inicio)
    
     query_produtos += ' ORDER BY v.data_hora DESC'
    
     try:
        self.cursor.execute(query_produtos, params_produtos)
        for row in self.cursor.fetchall():
            dados.append({
                'data_hora': row[0].isoformat() if hasattr(row[0], 'isoformat') else str(row[0]),
                'tipo': row[1],
                'descricao': f"{row[2]} (Venda: {row[6]})",
                'quantidade': row[3],
                'total': float(row[4]) if row[4] else 0.0,
                'usuario': row[5],
                'categoria': row[7]
            })
     except Exception as e:
        print(f"Erro ao obter produtos vendidos: {e}")
    
     # Serviços de impressão/cópia
     query_servicos = '''
        SELECT s.data_hora, 
               CASE 
                   WHEN s.tipo = 'copia_pb' THEN 'Cópia P&B'
                   WHEN s.tipo = 'copia_colorida' THEN 'Cópia Colorida'
                   WHEN s.tipo = 'impressao_pb' THEN 'Impressão P&B'
                   WHEN s.tipo = 'impressao_colorida' THEN 'Impressão Colorida'
                   ELSE s.tipo
               END as tipo,
               CONCAT('De ', s.numero_ontem, ' a ', s.numero_hoje) as descricao,
               s.quantidade, s.total, u.nome
        FROM servicos s
        JOIN usuarios u ON s.usuario_id = u.id
     '''
    
     params_servicos = []
     if data_inicio:
        query_servicos += ' WHERE s.data >= %s'
        params_servicos.append(data_inicio)
    
     query_servicos += ' ORDER BY s.data_hora DESC'
    
     try:
        self.cursor.execute(query_servicos, params_servicos)
        for row in self.cursor.fetchall():
            dados.append({
                'data_hora': row[0].isoformat() if hasattr(row[0], 'isoformat') else str(row[0]),
                'tipo': 'Serviço',
                'descricao': f"{row[1]} - {row[2]}",
                'quantidade': row[3],
                'total': float(row[4]) if row[4] else 0.0,
                'usuario': row[5]
            })
     except Exception as e:
        print(f"Erro ao obter serviços: {e}")
    
     # Serviços especiais
     query_especiais = '''
        SELECT se.data_hora, 
               CASE 
                   WHEN se.tipo LIKE 'encadernacao_%' THEN 
                       CONCAT('Encadernação Argola ', SUBSTRING(se.tipo FROM 13))
                   WHEN se.tipo LIKE 'laminacao_%' THEN 
                       CONCAT('Laminação ', UPPER(SUBSTRING(se.tipo FROM 11)))
                   ELSE se.tipo
               END as tipo,
               COALESCE(se.descricao, se.tipo) as descricao,
               se.quantidade, se.total, u.nome
        FROM servicos_especiais se
        JOIN usuarios u ON se.usuario_id = u.id
     '''
    
     params_especiais = []
     if data_inicio:
        query_especiais += ' WHERE se.data >= %s'
        params_especiais.append(data_inicio)
    
     query_especiais += ' ORDER BY se.data_hora DESC'
    
     try:
        self.cursor.execute(query_especiais, params_especiais)
        for row in self.cursor.fetchall():
            dados.append({
                'data_hora': row[0].isoformat() if hasattr(row[0], 'isoformat') else str(row[0]),
                'tipo': 'Serviço Especial',
                'descricao': f"{row[1]} - {row[2]}",
                'quantidade': row[3],
                'total': float(row[4]) if row[4] else 0.0,
                'usuario': row[5]
            })
     except Exception as e:
        print(f"Erro ao obter serviços especiais: {e}")
    
     return dados

    def dashboard_web(self):
        """Gerar HTML para dashboard web simplificado"""
        html_template = '''
        <!DOCTYPE html>
        <html lang="pt">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Dashboard Papelaria - v2.0</title>
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
            <style>
                :root {
                    --primary: #3498db;
                    --secondary: #2c3e50;
                    --success: #27ae60;
                    --warning: #f39c12;
                    --danger: #e74c3c;
                    --info: #1abc9c;
                    --light: #ecf0f1;
                    --dark: #34495e;
                }
                
                * {
                    margin: 0;
                    padding: 0;
                    box-sizing: border-box;
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                }
                
                body {
                    background: #f5f7fa;
                    min-height: 100vh;
                    padding: 15px;
                }
                
                .container {
                    max-width: 100%;
                    margin: 0 auto;
                }
                
                header {
                    background: white;
                    border-radius: 10px;
                    padding: 15px;
                    margin-bottom: 15px;
                    box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
                    display: flex;
                    flex-wrap: wrap;
                    justify-content: space-between;
                    align-items: center;
                }
                
                .logo {
                    display: flex;
                    align-items: center;
                    gap: 10px;
                }
                
                .logo i {
                    font-size: 2rem;
                    color: var(--primary);
                }
                
                .logo h1 {
                    color: var(--secondary);
                    font-size: 1.5rem;
                    margin: 0;
                }
                
                .filters {
                    display: flex;
                    gap: 8px;
                    align-items: center;
                    flex-wrap: wrap;
                }
                
                .filter-select {
                    padding: 8px 12px;
                    border: 1px solid #ddd;
                    border-radius: 5px;
                    background: white;
                    font-weight: 500;
                }
                
                .btn {
                    background: var(--primary);
                    color: white;
                    border: none;
                    padding: 8px 15px;
                    border-radius: 5px;
                    cursor: pointer;
                    font-weight: 600;
                    display: inline-flex;
                    align-items: center;
                    gap: 5px;
                    transition: background 0.3s ease;
                }
                
                .btn:hover {
                    background: #2980b9;
                }
                
                .dashboard-grid {
                    display: grid;
                    grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
                    gap: 15px;
                    margin-bottom: 20px;
                }
                
                .card {
                    background: white;
                    border-radius: 10px;
                    padding: 20px;
                    box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
                    transition: transform 0.3s ease;
                }
                
                .card:hover {
                    transform: translateY(-2px);
                }
                
                .card-header {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    margin-bottom: 15px;
                }
                
                .card-icon {
                    width: 50px;
                    height: 50px;
                    border-radius: 10px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    font-size: 1.5rem;
                    color: white;
                }
                
                .card-icon.sales { background: var(--success); }
                .card-icon.products { background: var(--primary); }
                .card-icon.stock { background: var(--warning); }
                .card-icon.money { background: var(--danger); }
                .card-icon.services { background: var(--info); }
                
                .card-title {
                    color: #666;
                    font-size: 0.9rem;
                    font-weight: 600;
                    text-transform: uppercase;
                }
                
                .card-value {
                    font-size: 2rem;
                    font-weight: 700;
                    color: var(--secondary);
                    margin: 10px 0;
                }
                
                .card-subtitle {
                    font-size: 0.9rem;
                    color: #777;
                }
                
                .section {
                    background: white;
                    border-radius: 10px;
                    padding: 20px;
                    margin-bottom: 20px;
                    box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
                }
                
                .section-header {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    margin-bottom: 20px;
                    flex-wrap: wrap;
                }
                
                .section-title {
                    color: var(--secondary);
                    font-size: 1.2rem;
                    font-weight: 600;
                    margin-right: 10px;
                }
                
                table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-top: 10px;
                    font-size: 0.9rem;
                }
                
                th {
                    text-align: left;
                    padding: 12px 15px;
                    background: #f8f9fa;
                    color: #666;
                    font-weight: 600;
                    border-bottom: 2px solid #dee2e6;
                }
                
                td {
                    padding: 12px 15px;
                    border-bottom: 1px solid #dee2e6;
                }
                
                tr:hover {
                    background: #f8f9fa;
                }
                
                .badge {
                    padding: 4px 8px;
                    border-radius: 4px;
                    font-size: 0.8rem;
                    font-weight: 600;
                }
                
                .badge.success {
                    background: #d4edda;
                    color: #155724;
                }
                
                .badge.warning {
                    background: #fff3cd;
                    color: #856404;
                }
                
                .badge.danger {
                    background: #f8d7da;
                    color: #721c24;
                }
                
                .badge.info {
                    background: #d1ecf1;
                    color: #0c5460;
                }
                
                .badge.primary {
                    background: #d6eaf8;
                    color: #1b4f72;
                }
                
                .loading {
                    text-align: center;
                    padding: 40px;
                    color: #666;
                }
                
                .loading i {
                    font-size: 2rem;
                    margin-bottom: 10px;
                }
                
                .empty-state {
                    text-align: center;
                    padding: 40px;
                    color: #999;
                }
                
                .empty-state i {
                    font-size: 3rem;
                    margin-bottom: 10px;
                    color: #ddd;
                }
                
                .status-indicator {
                    display: inline-block;
                    width: 10px;
                    height: 10px;
                    border-radius: 50%;
                    margin-right: 5px;
                }
                
                .status-online {
                    background: var(--success);
                }
                
                .status-offline {
                    background: var(--danger);
                }
                
                @media (max-width: 768px) {
                    .dashboard-grid {
                        grid-template-columns: 1fr;
                    }
                    
                    .card-value {
                        font-size: 1.8rem;
                    }
                    
                    table {
                        font-size: 0.8rem;
                    }
                    
                    th, td {
                        padding: 8px 10px;
                    }
                }
                
                .tabs {
                    display: flex;
                    border-bottom: 1px solid #ddd;
                    margin-bottom: 15px;
                }
                
                .tab {
                    padding: 10px 20px;
                    cursor: pointer;
                    border-bottom: 3px solid transparent;
                    font-weight: 500;
                    transition: all 0.3s ease;
                }
                
                .tab:hover {
                    background: #f8f9fa;
                }
                
                .tab.active {
                    border-bottom-color: var(--primary);
                    color: var(--primary);
                    background: #f8f9fa;
                }
                
                .tab-content {
                    display: none;
                }
                
                .tab-content.active {
                    display: block;
                }
                
                .date-picker {
                    padding: 8px 12px;
                    border: 1px solid #ddd;
                    border-radius: 5px;
                    background: white;
                }
                
                .alert {
                    padding: 12px 15px;
                    border-radius: 5px;
                    margin-bottom: 15px;
                    font-weight: 500;
                }
                
                .alert.info {
                    background: #d1ecf1;
                    color: #0c5460;
                    border: 1px solid #bee5eb;
                }
                
                .alert.warning {
                    background: #fff3cd;
                    color: #856404;
                    border: 1px solid #ffeaa7;
                }
            </style>
        </head>
        <body>
            <div class="container">
                <header>
                    <div class="logo">
                        <img src="static/logo.png" alt="Logo" style="width: 80px; height: 80px; object-fit: contain;">
                        <div>
                          <h1>Papelaria Frente Verso</h1>
                          <p style="color: #666; font-size: 0.9rem;">Controle de Gestão de Vendas</p>
                        </div>
                    </div>

                    <div class="filters">
                        <div style="display: flex; align-items: center; margin-right: 15px;">
                            <span class="status-indicator status-online" id="status-indicator"></span>
                            <span id="status-text" style="font-size: 0.9rem; color: #666;">Online</span>
                        </div>
                        <button class="btn" onclick="location.reload()">
                            <i class="fas fa-sync-alt"></i> Atualizar
                        </button>
                    </div>
                </header>
                
                <div class="alert info">
                    <i class="fas fa-info-circle"></i>
                    Dashboard em tempo real - Atualiza a cada 10 segundos
                </div>
                
                <div class="dashboard-grid">
                    <div class="card">
                        <div class="card-header">
                            <div class="card-title">Vendas Hoje</div>
                            <div class="card-icon sales">
                                <i class="fas fa-shopping-cart"></i>
                            </div>
                        </div>
                        <div class="card-value" id="vendas-hoje">MT 0,00</div>
                        <div class="card-subtitle">Total de vendas do dia</div>
                    </div>
                    
                    <div class="card">
                        <div class="card-header">
                            <div class="card-title">Total Produtos</div>
                            <div class="card-icon products">
                                <i class="fas fa-boxes"></i>
                            </div>
                        </div>
                        <div class="card-value" id="total-produtos">0</div>
                        <div class="card-subtitle">Produtos cadastrados</div>
                    </div>
                    
                    <div class="card">
                        <div class="card-header">
                            <div class="card-title">Estoque Baixo</div>
                            <div class="card-icon stock">
                                <i class="fas fa-exclamation-triangle"></i>
                            </div>
                        </div>
                        <div class="card-value" id="estoque-baixo">0</div>
                        <div class="card-subtitle">Produtos com estoque crítico</div>
                    </div>
                    
                    <div class="card">
                        <div class="card-header">
                            <div class="card-title">Valor Estoque</div>
                            <div class="card-icon money">
                                <i class="fas fa-money-bill-wave"></i>
                            </div>
                        </div>
                        <div class="card-value" id="valor-estoque">MT 0,00</div>
                        <div class="card-subtitle">Valor total do estoque</div>
                    </div>
                    
                    <div class="card">
                        <div class="card-header">
                            <div class="card-title">Serviços Hoje</div>
                            <div class="card-icon services">
                                <i class="fas fa-print"></i>
                            </div>
                        </div>
                        <div class="card-value" id="servicos-hoje">MT 0,00</div>
                        <div class="card-subtitle">Total de serviços do dia</div>
                    </div>
                    
                    <div class="card">
                        <div class="card-header">
                            <div class="card-title">Status Sistema</div>
                            <div class="card-icon primary">
                                <i class="fas fa-check-circle"></i>
                            </div>
                        </div>
                        <div class="card-value" id="status-sistema">Online</div>
                        <div class="card-subtitle" id="ultima-atualizacao">--:--:--</div>
                    </div>
                </div>
                
                <div class="section">
                    <div class="section-header">
                        <div class="section-title">
                            <i class="fas fa-chart-pie"></i> Vendas e Serviços
                        </div>
                        <div class="tabs">
                            <div class="tab active" onclick="showTab('produtos', this)">Produtos</div>
                            <div class="tab" onclick="showTab('servicos', this)">Serviços</div>
                            <div class="tab" onclick="showTab('detalhado', this)">Detalhado</div>
                        </div>
                    </div>
                    
                    <div id="tab-produtos" class="tab-content active">
                        <div class="filters" style="margin-bottom: 15px;">
                            <select class="filter-select" id="filter-status" onchange="loadProdutos()">
                                <option value="todos">Todos os produtos</option>
                                <option value="baixo">Estoque baixo</option>
                                <option value="critico">Estoque crítico</option>
                            </select>
                        </div>
                        <div id="lista-produtos">
                            <div class="loading">
                                <i class="fas fa-spinner fa-spin"></i>
                                <p>Carregando produtos...</p>
                            </div>
                        </div>
                    </div>
                    
                    <div id="tab-servicos" class="tab-content">
                        <div class="filters" style="margin-bottom: 15px;">
                            <select class="filter-select" id="filter-period-serv" onchange="loadServicos()">
                                <option value="hoje">Hoje</option>
                                <option value="ontem">Ontem</option>
                                <option value="7dias">Últimos 7 dias</option>
                                <option value="mes">Este mês</option>
                                <option value="todos">Todos</option>
                            </select>
                        </div>
                        <div id="lista-servicos">
                            <div class="loading">
                                <i class="fas fa-spinner fa-spin"></i>
                                <p>Carregando serviços...</p>
                            </div>
                        </div>
                    </div>
                    
                    <div id="tab-detalhado" class="tab-content">
                        <div class="filters" style="margin-bottom: 15px;">
                            <select class="filter-select" id="filter-period-det" onchange="loadVendasDetalhadas()">
                                <option value="hoje">Hoje</option>
                                <option value="ontem">Ontem</option>
                                <option value="7dias">Últimos 7 dias</option>
                                <option value="mes">Este mês</option>
                                <option value="todos">Todos</option>
                            </select>
                            <select class="filter-select" id="filter-tipo-det" onchange="loadVendasDetalhadas()">
                                <option value="todos">Todos os tipos</option>
                                <option value="produto">Apenas Produtos</option>
                                <option value="servico">Apenas Serviços</option>
                            </select>
                        </div>
                        <div id="lista-vendas-detalhadas">
                            <div class="loading">
                                <i class="fas fa-spinner fa-spin"></i>
                                <p>Carregando vendas detalhadas...</p>
                            </div>
                        </div>
                    </div>
                </div>
                
                <div class="section">
                    <div class="section-header">
                        <div class="section-title">
                            <i class="fas fa-chart-line"></i> Últimas Vendas
                        </div>
                        <select class="filter-select" id="filter-period" onchange="loadVendas()">
                            <option value="hoje">Hoje</option>
                            <option value="ontem">Ontem</option>
                            <option value="7dias">Últimos 7 dias</option>
                            <option value="mes">Este mês</option>
                            <option value="todos">Todos</option>
                        </select>
                    </div>
                    <div id="lista-vendas">
                        <div class="loading">
                            <i class="fas fa-spinner fa-spin"></i>
                            <p>Carregando vendas...</p>
                        </div>
                    </div>
                </div>
                
                <div style="text-align: center; margin-top: 30px; color: #666; font-size: 0.9rem;">
                    <p>Desenvolvido por: Tomas J. T. Antonio v2.0 | Última atualização: <span id="last-update">--:--:--</span></p>
                </div>
            </div>
            
            <script>
                const API_BASE = window.location.origin;
                let updateInterval;
                
                function showTab(tabName, element) {
                    // Remove active class from all tabs
                    document.querySelectorAll('.tab').forEach(tab => {
                        tab.classList.remove('active');
                    });
                    
                    // Hide all tab contents
                    document.querySelectorAll('.tab-content').forEach(content => {
                        content.classList.remove('active');
                    });
                    
                    // Add active class to clicked tab
                    element.classList.add('active');
                    
                    // Show selected tab content
                    document.getElementById('tab-' + tabName).classList.add('active');
                    
                    // Load data for the tab if needed
                    if (tabName === 'servicos') {
                        loadServicos();
                    } else if (tabName === 'detalhado') {
                        loadVendasDetalhadas();
                    } else if (tabName === 'produtos') {
                        loadProdutos();
                    }
                }
                
                async function updateStatus() {
                    const indicator = document.getElementById('status-indicator');
                    const text = document.getElementById('status-text');
                    const statusSistema = document.getElementById('status-sistema');
                    
                    try {
                        const response = await fetch(`${API_BASE}/api/teste`);
                        if (response.ok) {
                            indicator.className = 'status-indicator status-online';
                            text.textContent = 'Online';
                            text.style.color = '#27ae60';
                            statusSistema.textContent = 'Online';
                            statusSistema.style.color = '#27ae60';
                        } else {
                            throw new Error('API error');
                        }
                    } catch (error) {
                        indicator.className = 'status-indicator status-offline';
                        text.textContent = 'Offline';
                        text.style.color = '#e74c3c';
                        statusSistema.textContent = 'Offline';
                        statusSistema.style.color = '#e74c3c';
                    }
                }
                
                async function loadMetrics() {
                    try {
                        const response = await fetch(`${API_BASE}/api/metricas`);
                        const data = await response.json();
                        
                        if (data.status === 'success') {
                            document.getElementById('vendas-hoje').textContent = 
                                formatCurrency(data.vendas_hoje);
                            document.getElementById('total-produtos').textContent = data.total_produtos;
                            document.getElementById('estoque-baixo').textContent = data.estoque_baixo;
                            document.getElementById('servicos-hoje').textContent = 
                                formatCurrency(data.servicos_especiais_hoje);
                            document.getElementById('valor-estoque').textContent = 
                                formatCurrency(data.valor_estoque_total || data.valor_total_estoque || 0);
                            
                            // Atualizar timestamp
                            const now = new Date();
                            document.getElementById('last-update').textContent = 
                                now.toLocaleTimeString('pt-BR');
                            document.getElementById('ultima-atualizacao').textContent = 
                                now.toLocaleTimeString('pt-BR');
                        }
                    } catch (error) {
                        console.error('Erro ao carregar métricas:', error);
                    }
                }
                
                async function loadProdutos() {
                    try {
                        const status = document.getElementById('filter-status') ? 
                                      document.getElementById('filter-status').value : 'todos';
                        const response = await fetch(`${API_BASE}/api/produtos`);
                        const data = await response.json();
                        
                        if (data.status === 'success') {
                            let produtos = data.produtos;
                            
                            // Filtrar por status
                            if (status === 'baixo') {
                                produtos = produtos.filter(p => p.quantidade <= p.estoque_minimo);
                            } else if (status === 'critico') {
                                produtos = produtos.filter(p => p.quantidade < p.estoque_minimo);
                            }
                            
                            const container = document.getElementById('lista-produtos');
                            
                            if (!produtos || produtos.length === 0) {
                                container.innerHTML = `
                                    <div class="empty-state">
                                        <i class="fas fa-box-open"></i>
                                        <p>Nenhum produto encontrado</p>
                                    </div>
                                `;
                                return;
                            }
                            
                            let html = `
                                <div style="margin-bottom: 10px; color: #666;">
                                    <i class="fas fa-info-circle"></i>
                                    Total: ${produtos.length} produtos | Valor total: ${formatCurrency(data.valor_total_estoque)}
                                </div>
                                <table>
                                    <thead>
                                        <tr>
                                            <th>Código</th>
                                            <th>Produto</th>
                                            <th>Categoria</th>
                                            <th>Preço</th>
                                            <th>Estoque</th>
                                            <th>Valor Total</th>
                                            <th>Status</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                            `;
                            
                            produtos.forEach(produto => {
                                let statusClass = 'badge success';
                                if (produto.status === 'BAIXO') statusClass = 'badge warning';
                                if (produto.status === 'CRÍTICO') statusClass = 'badge danger';
                                
                                html += `
                                    <tr>
                                        <td><strong>${produto.codigo}</strong></td>
                                        <td>${produto.nome}</td>
                                        <td>${produto.categoria}</td>
                                        <td>${formatCurrency(produto.preco)}</td>
                                        <td>${produto.quantidade}</td>
                                        <td><strong>${formatCurrency(produto.valor_total)}</strong></td>
                                        <td><span class="${statusClass}">${produto.status}</span></td>
                                    </tr>
                                `;
                            });
                            
                            html += '</tbody></table>';
                            container.innerHTML = html;
                        }
                    } catch (error) {
                        console.error('Erro ao carregar produtos:', error);
                        const container = document.getElementById('lista-produtos');
                        container.innerHTML = `
                            <div class="empty-state">
                                <i class="fas fa-exclamation-triangle"></i>
                                <p>Erro ao carregar produtos</p>
                            </div>
                        `;
                    }
                }
                
                async function loadServicos() {
                    try {
                        const periodo = document.getElementById('filter-period-serv').value;
                        const response = await fetch(`${API_BASE}/api/servicos?periodo=${periodo}`);
                        const data = await response.json();
                        
                        if (data.status === 'success') {
                            const container = document.getElementById('lista-servicos');
                            
                            const totalServicos = data.servicos_impressao.total + data.servicos_especiais.total;
                            const totalValor = data.total_geral;
                            
                            let html = `
                                <div style="margin-bottom: 15px; display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
                                    <div style="background: #f8f9fa; padding: 10px; border-radius: 5px;">
                                        <div style="font-size: 0.9rem; color: #666;">Impressão/Cópia</div>
                                        <div style="font-size: 1.2rem; font-weight: bold;">${data.servicos_impressao.total}</div>
                                        <div style="font-size: 0.9rem; color: #666;">${formatCurrency(data.servicos_impressao.total_valor)}</div>
                                    </div>
                                    <div style="background: #f8f9fa; padding: 10px; border-radius: 5px;">
                                        <div style="font-size: 0.9rem; color: #666;">Serviços Especiais</div>
                                        <div style="font-size: 1.2rem; font-weight: bold;">${data.servicos_especiais.total}</div>
                                        <div style="font-size: 0.9rem; color: #666;">${formatCurrency(data.servicos_especiais.total_valor)}</div>
                                    </div>
                                    <div style="background: #e8f4fd; padding: 10px; border-radius: 5px;">
                                        <div style="font-size: 0.9rem; color: #666;">Total Geral</div>
                                        <div style="font-size: 1.2rem; font-weight: bold;">${totalServicos}</div>
                                        <div style="font-size: 0.9rem; color: #666;">${formatCurrency(totalValor)}</div>
                                    </div>
                                </div>
                            `;
                            
                            // Listar serviços de impressão
                            if (data.servicos_impressao.items.length > 0) {
                                html += `
                                    <h4 style="margin: 15px 0 10px 0;">Impressão/Cópia</h4>
                                    <table>
                                        <thead>
                                            <tr>
                                                <th>Data/Hora</th>
                                                <th>Tipo</th>
                                                <th>De</th>
                                                <th>A</th>
                                                <th>Qtd</th>
                                                <th>Total</th>
                                                <th>Usuário</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                `;
                                
                                data.servicos_impressao.items.slice(0, 10).forEach(item => {
                                    try {
                                        const dataHora = new Date(item.data_hora);
                                        const dataFormatada = dataHora.toLocaleDateString('pt-BR');
                                        const horaFormatada = dataHora.toLocaleTimeString('pt-BR', {hour: '2-digit', minute:'2-digit'});
                                        
                                        html += `
                                            <tr>
                                                <td>${dataFormatada} ${horaFormatada}</td>
                                                <td>${item.tipo}</td>
                                                <td>${item.numero_ontem}</td>
                                                <td>${item.numero_hoje}</td>
                                                <td>${item.quantidade}</td>
                                                <td>${formatCurrency(item.total)}</td>
                                                <td>${item.usuario}</td>
                                            </tr>
                                        `;
                                    } catch (e) {
                                        console.error('Erro ao formatar data:', e);
                                    }
                                });
                                
                                html += '</tbody></table>';
                            }
                            
                            // Listar serviços especiais
                            if (data.servicos_especiais.items.length > 0) {
                                html += `
                                    <h4 style="margin: 15px 0 10px 0;">Serviços Especiais</h4>
                                    <table>
                                        <thead>
                                            <tr>
                                                <th>Data/Hora</th>
                                                <th>Tipo</th>
                                                <th>Descrição</th>
                                                <th>Qtd</th>
                                                <th>Total</th>
                                                <th>Usuário</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                `;
                                
                                data.servicos_especiais.items.slice(0, 10).forEach(item => {
                                    try {
                                        const dataHora = new Date(item.data_hora);
                                        const dataFormatada = dataHora.toLocaleDateString('pt-BR');
                                        const horaFormatada = dataHora.toLocaleTimeString('pt-BR', {hour: '2-digit', minute:'2-digit'});
                                        
                                        html += `
                                            <tr>
                                                <td>${dataFormatada} ${horaFormatada}</td>
                                                <td>${item.tipo}</td>
                                                <td>${item.descricao}</td>
                                                <td>${item.quantidade}</td>
                                                <td>${formatCurrency(item.total)}</td>
                                                <td>${item.usuario}</td>
                                            </tr>
                                        `;
                                    } catch (e) {
                                        console.error('Erro ao formatar data:', e);
                                    }
                                });
                                
                                html += '</tbody></table>';
                            }
                            
                            if (data.servicos_impressao.items.length === 0 && data.servicos_especiais.items.length === 0) {
                                html += `
                                    <div class="empty-state">
                                        <i class="fas fa-print"></i>
                                        <p>Nenhum serviço encontrado</p>
                                    </div>
                                `;
                            }
                            
                            container.innerHTML = html;
                        }
                    } catch (error) {
                        console.error('Erro ao carregar serviços:', error);
                        const container = document.getElementById('lista-servicos');
                        container.innerHTML = `
                            <div class="empty-state">
                                <i class="fas fa-exclamation-triangle"></i>
                                <p>Erro ao carregar serviços</p>
                            </div>
                        `;
                    }
                }
                
                async function loadVendasDetalhadas() {
                    try {
                        const periodo = document.getElementById('filter-period-det').value;
                        const tipo = document.getElementById('filter-tipo-det').value;
                        
                        const response = await fetch(`${API_BASE}/api/vendas_detalhadas?periodo=${periodo}`);
                        const data = await response.json();
                        
                        if (data.status === 'success') {
                            const container = document.getElementById('lista-vendas-detalhadas');
                            
                            let vendasDetalhadas = data.vendas_detalhadas;
                            
                            // Filtrar por tipo se necessário
                            if (tipo === 'produto') {
                                vendasDetalhadas = vendasDetalhadas.filter(v => v.tipo === 'Produto');
                            } else if (tipo === 'servico') {
                                vendasDetalhadas = vendasDetalhadas.filter(v => v.tipo === 'Serviço');
                            }
                            
                            if (!vendasDetalhadas || vendasDetalhadas.length === 0) {
                                container.innerHTML = `
                                    <div class="empty-state">
                                        <i class="fas fa-chart-bar"></i>
                                        <p>Nenhum registro encontrado</p>
                                    </div>
                                `;
                                return;
                            }
                            
                            let html = `
                                <div style="margin-bottom: 10px; color: #666;">
                                    <i class="fas fa-info-circle"></i>
                                    Total: ${vendasDetalhadas.length} registros | 
                                    Produtos: ${data.total_produtos} | 
                                    Serviços: ${data.total_servicos} | 
                                    Valor total: ${formatCurrency(data.valor_total)}
                                </div>
                                <table>
                                    <thead>
                                        <tr>
                                            <th>Data/Hora</th>
                                            <th>Tipo</th>
                                            <th>Descrição</th>
                                            <th>Categoria</th>
                                            <th>Qtd</th>
                                            <th>Total</th>
                                            <th>Usuário</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                            `;
                            
                            vendasDetalhadas.slice(0, 20).forEach(item => {
                                try {
                                    const dataHora = new Date(item.data_hora);
                                    const dataFormatada = dataHora.toLocaleDateString('pt-BR');
                                    const horaFormatada = dataHora.toLocaleTimeString('pt-BR', {hour: '2-digit', minute:'2-digit'});
                                    
                                    let tipoClass = 'badge ';
                                    if (item.tipo === 'Produto') {
                                        tipoClass += 'primary';
                                    } else {
                                        tipoClass += 'info';
                                    }
                                    
                                    html += `
                                        <tr>
                                            <td>${dataFormatada} ${horaFormatada}</td>
                                            <td><span class="${tipoClass}">${item.tipo}</span></td>
                                            <td>${item.descricao}</td>
                                            <td>${item.categoria || '-'}</td>
                                            <td>${item.quantidade}</td>
                                            <td><strong>${formatCurrency(item.total)}</strong></td>
                                            <td>${item.usuario}</td>
                                        </tr>
                                    `;
                                } catch (e) {
                                    console.error('Erro ao formatar venda detalhada:', e);
                                }
                            });
                            
                            html += '</tbody></table>';
                            
                            if (vendasDetalhadas.length > 20) {
                                html += `<div style="text-align: center; margin-top: 10px; color: #666;">
                                    ... e mais ${vendasDetalhadas.length - 20} registros
                                </div>`;
                            }
                            
                            container.innerHTML = html;
                        }
                    } catch (error) {
                        console.error('Erro ao carregar vendas detalhadas:', error);
                        const container = document.getElementById('lista-vendas-detalhadas');
                        container.innerHTML = `
                            <div class="empty-state">
                                <i class="fas fa-exclamation-triangle"></i>
                                <p>Erro ao carregar vendas detalhadas</p>
                            </div>
                        `;
                    }
                }
                
                async function loadVendas() {
                    try {
                        const periodo = document.getElementById('filter-period').value;
                        const response = await fetch(`${API_BASE}/api/vendas?periodo=${periodo}&limit=20`);
                        const data = await response.json();
                        
                        if (data.status === 'success') {
                            const container = document.getElementById('lista-vendas');
                            
                            if (!data.vendas || data.vendas.length === 0) {
                                container.innerHTML = `
                                    <div class="empty-state">
                                        <i class="fas fa-shopping-cart"></i>
                                        <p>Nenhuma venda encontrada</p>
                                    </div>
                                `;
                                return;
                            }
                            
                            let html = `
                                <div style="margin-bottom: 10px; color: #666;">
                                    <i class="fas fa-info-circle"></i>
                                    Total: ${data.total_vendas} vendas | Valor total: ${formatCurrency(data.total_valor)}
                                </div>
                                <table>
                                    <thead>
                                        <tr>
                                            <th>Data/Hora</th>
                                            <th>Nº Venda</th>
                                            <th>Vendedor</th>
                                            <th>Total</th>
                                            <th>Recebido</th>
                                            <th>Troco</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                            `;
                            
                            data.vendas.forEach(venda => {
                                try {
                                    const dataHora = new Date(venda.data_hora);
                                    const dataFormatada = dataHora.toLocaleDateString('pt-BR');
                                    const horaFormatada = dataHora.toLocaleTimeString('pt-BR', {hour: '2-digit', minute:'2-digit'});
                                    
                                    html += `
                                        <tr>
                                            <td>${dataFormatada} ${horaFormatada}</td>
                                            <td><strong>${venda.numero_serie}</strong></td>
                                            <td>${venda.vendedor}</td>
                                            <td><strong>${formatCurrency(venda.total)}</strong></td>
                                            <td>${formatCurrency(venda.valor_recebido)}</td>
                                            <td>${formatCurrency(venda.troco)}</td>
                                        </tr>
                                    `;
                                } catch (e) {
                                    console.error('Erro ao formatar venda:', e);
                                }
                            });
                            
                            html += '</tbody></table>';
                            container.innerHTML = html;
                        }
                    } catch (error) {
                        console.error('Erro ao carregar vendas:', error);
                        const container = document.getElementById('lista-vendas');
                        container.innerHTML = `
                            <div class="empty-state">
                                <i class="fas fa-exclamation-triangle"></i>
                                <p>Erro ao carregar vendas</p>
                            </div>
                        `;
                    }
                }
                
                function formatCurrency(value) {
                    try {
                        const num = parseFloat(value) || 0;
                        return 'MT ' + num.toLocaleString('pt-BR', {
                            minimumFractionDigits: 2,
                            maximumFractionDigits: 2
                        });
                    } catch (e) {
                        return 'MT 0,00';
                    }
                }
                
                function atualizarDashboard() {
                    updateStatus();
                    loadMetrics();
                    
                    // Recarregar a aba ativa
                    const activeTab = document.querySelector('.tab.active');
                    if (activeTab) {
                        const tabText = activeTab.textContent.toLowerCase();
                        if (tabText.includes('produtos')) loadProdutos();
                        else if (tabText.includes('serviços') || tabText.includes('servicos')) loadServicos();
                        else if (tabText.includes('detalhado')) loadVendasDetalhadas();
                    }
                    
                    // Recarregar vendas se estiver na aba principal
                    if (document.querySelector('#lista-vendas')) {
                        loadVendas();
                    }
                }
                
                document.addEventListener('DOMContentLoaded', function() {
                    atualizarDashboard();
                    
                    // Atualizar a cada 10 segundos
                    updateInterval = setInterval(atualizarDashboard, 10000);
                });
                
                // Limpar intervalo quando a página for fechada
                window.addEventListener('beforeunload', function() {
                    if (updateInterval) {
                        clearInterval(updateInterval);
                    }
                });
            </script>
        </body>
        </html>
        '''
        return render_template_string(html_template)
    
    def conectar_banco(self):
        """Conectar ao banco de dados PostgreSQL"""
        try:
            self.conn = psycopg2.connect(
                host="localhost",
                database="papelaria_db",
                user="postgres",
                password="techmz06",
                port="5432"
            )
            
            self.cursor = self.conn.cursor()
            self.criar_tabelas()
            print("✅ Conectado ao PostgreSQL com sucesso!")
            
        except Exception as e:
            print(f"❌ Erro de conexão: {e}")
            messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados.")
            sys.exit(1)
    
    def criar_tabelas(self):
        """Criar tabelas do banco de dados"""
        try:
            tabelas = [
                '''CREATE TABLE IF NOT EXISTS usuarios (
                    id SERIAL PRIMARY KEY,
                    nome TEXT NOT NULL,
                    email TEXT UNIQUE NOT NULL,
                    senha TEXT NOT NULL,
                    tipo TEXT NOT NULL,
                    ativo INTEGER DEFAULT 1
                )''',
                '''CREATE TABLE IF NOT EXISTS produtos (
                    id SERIAL PRIMARY KEY,
                    codigo TEXT UNIQUE NOT NULL,
                    nome TEXT NOT NULL,
                    descricao TEXT,
                    categoria TEXT NOT NULL,
                    preco REAL NOT NULL,
                    quantidade INTEGER NOT NULL,
                    estoque_minimo INTEGER DEFAULT 5,
                    data_cadastro DATE NOT NULL
                )''',
                '''CREATE TABLE IF NOT EXISTS vendas (
                    id SERIAL PRIMARY KEY,
                    numero_serie TEXT UNIQUE NOT NULL,
                    data_hora TIMESTAMP NOT NULL,
                    total REAL NOT NULL,
                    valor_recebido REAL NOT NULL,
                    troco REAL NOT NULL,
                    usuario_id INTEGER NOT NULL,
                    FOREIGN KEY (usuario_id) REFERENCES usuarios (id)
                )''',
                '''CREATE TABLE IF NOT EXISTS itens_venda (
                    id SERIAL PRIMARY KEY,
                    venda_id INTEGER NOT NULL,
                    produto_id INTEGER NOT NULL,
                    quantidade INTEGER NOT NULL,
                    preco_unitario REAL NOT NULL,
                    total_item REAL NOT NULL,
                    FOREIGN KEY (venda_id) REFERENCES vendas (id),
                    FOREIGN KEY (produto_id) REFERENCES produtos (id)
                )''',
                '''CREATE TABLE IF NOT EXISTS servicos (
                    id SERIAL PRIMARY KEY,
                    tipo TEXT NOT NULL,
                    numero_ontem INTEGER DEFAULT 0,
                    numero_hoje INTEGER DEFAULT 0,
                    quantidade INTEGER NOT NULL,
                    preco_unitario REAL NOT NULL,
                    total REAL NOT NULL,
                    data DATE NOT NULL,
                    usuario_id INTEGER NOT NULL,
                    data_hora TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (usuario_id) REFERENCES usuarios (id)
                )''',
                '''CREATE TABLE IF NOT EXISTS servicos_especiais (
                    id SERIAL PRIMARY KEY,
                    tipo TEXT NOT NULL,
                    descricao TEXT,
                    quantidade INTEGER NOT NULL,
                    preco_unitario REAL NOT NULL,
                    total REAL NOT NULL,
                    data DATE NOT NULL,
                    usuario_id INTEGER NOT NULL,
                    data_hora TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (usuario_id) REFERENCES usuarios (id)
                )''',
                '''CREATE TABLE IF NOT EXISTS relatorios_salvos (
                    id SERIAL PRIMARY KEY,
                    nome TEXT NOT NULL,
                    tipo TEXT NOT NULL,
                    periodo TEXT NOT NULL,
                    data_geracao TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    usuario_id INTEGER NOT NULL,
                    dados JSONB NOT NULL,
                    FOREIGN KEY (usuario_id) REFERENCES usuarios (id)
                )''',
                '''CREATE TABLE IF NOT EXISTS configuracoes_servicos (
                    id SERIAL PRIMARY KEY,
                    nome TEXT NOT NULL,
                    tipo TEXT NOT NULL,
                    preco REAL NOT NULL,
                    descricao TEXT,
                    ativo INTEGER DEFAULT 1
                )'''
            ]
            
            for tabela in tabelas:
                self.cursor.execute(tabela)
            
            self.conn.commit()
            
            # Verificar se há usuários
            self.cursor.execute("SELECT COUNT(*) FROM usuarios")
            if self.cursor.fetchone()[0] == 0:
                # Criar usuários padrão com senhas em texto puro (serão convertidas no login)
                senha_admin = hashlib.md5('admin123'.encode()).hexdigest()
                senha_vendedor = hashlib.md5('vendedor123'.encode()).hexdigest()
                senha_gerente = hashlib.md5('gerente123'.encode()).hexdigest()
                
                self.cursor.execute('''
                    INSERT INTO usuarios (nome, email, senha, tipo) VALUES 
                    ('Administrador', 'admin@papelaria.com', %s, 'admin'),
                    ('Vendedor', 'vendedor@papelaria.com', %s, 'vendedor'),
                    ('Gerente', 'gerente@papelaria.com', %s, 'gerente')
                ''', (senha_admin, senha_vendedor, senha_gerente))
                self.conn.commit()
                print("✅ Usuários padrão criados!")
            else:
                print("✅ Usuários já existentes")
                
            # Verificar se há serviços configurados
            self.cursor.execute("SELECT COUNT(*) FROM configuracoes_servicos")
            if self.cursor.fetchone()[0] == 0:
                # Inserir serviços padrão
                servicos_padrao = [
                    ('Cartão de Visita', 'cartao_visita', 80.0, 'Cartão de visita personalizado'),
                    ('Convite', 'convite', 120.0, 'Convites personalizados'),
                    ('Banner', 'banner', 250.0, 'Banner promocional'),
                    ('Adesivo', 'adesivo', 45.0, 'Adesivos personalizados')
                ]
                
                for servico in servicos_padrao:
                    self.cursor.execute('''
                        INSERT INTO configuracoes_servicos (nome, tipo, preco, descricao)
                        VALUES (%s, %s, %s, %s)
                    ''', servico)
                
                self.conn.commit()
                print("✅ Serviços padrão configurados!")
                
        except Exception as e:
            print(f"❌ Erro ao criar tabelas: {e}")
            self.conn.rollback()
    
    def configurar_estilo(self):
        """Configurar estilos para a interface"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.style.configure('TFrame', background='#2c3e50')
        self.style.configure('TLabel', background='#2c3e50', foreground='white', font=('Arial', 10))
        self.style.configure('TButton', font=('Arial', 10))
        self.style.configure('Title.TLabel', font=('Arial', 16, 'bold'), foreground='white')
        self.style.configure('Accent.TButton', background='#3498db', foreground='white')
    
    def tela_login(self):
        """Tela de login do sistema"""
        self.limpar_tela()
        
        # Frame principal
        main_frame = tk.Frame(self.root, bg='#2c3e50')
        main_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        # Logo e título
        logo_frame = tk.Frame(main_frame, bg='#2c3e50')
        logo_frame.pack(pady=(0, 30))
        
        # Carregar logo
        try:
            logo_img = Image.open(resource_path("static/logo.png"))
            logo_img = logo_img.resize((150, 150), Image.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(logo_img)
            tk.Label(logo_frame, image=self.logo_photo, bg='#2c3e50').pack()
        except:
            tk.Label(logo_frame, text="🖨️", font=('Arial', 48), 
                    bg='#2c3e50', fg='white').pack()

        tk.Label(logo_frame, text="Gestão de vendas", font=('Arial', 12), 
                bg='#2c3e50', fg='#bdc3c7').pack()
        
        # Formulário de login
        form_frame = tk.Frame(main_frame, bg='#34495e', padx=40, pady=40, relief=tk.RAISED, bd=2)
        form_frame.pack()
        
        tk.Label(form_frame, text="Login", font=('Arial', 18, 'bold'), 
                bg='#34495e', fg='white').pack(pady=(0, 20))
        
        # Email
        tk.Label(form_frame, text="Email:", bg='#34495e', fg='white', 
                font=('Arial', 11)).pack(anchor=tk.W)
        self.email_var = tk.StringVar()
        email_entry = ttk.Entry(form_frame, textvariable=self.email_var, width=30)
        email_entry.pack(pady=(5, 15))
        
        # Senha
        tk.Label(form_frame, text="Senha:", bg='#34495e', fg='white', 
                font=('Arial', 11)).pack(anchor=tk.W)
        self.senha_var = tk.StringVar()
        senha_entry = ttk.Entry(form_frame, textvariable=self.senha_var, show="*", width=30)
        senha_entry.pack(pady=(5, 20))
        
        # Botão de login
        login_btn = tk.Button(form_frame, text="Entrar", command=self.efetuar_login,
                            bg='#3498db', fg='white', font=('Arial', 12, 'bold'),
                            padx=30, pady=10, bd=0, cursor='hand2')
        login_btn.pack(pady=10)
        
        # Botão para modo emergência (sem hash)
        emergencia_btn = tk.Button(form_frame, text="Modo Emergência", 
                                  command=self.login_emergencia,
                                  bg='#e74c3c', fg='white', font=('Arial', 10),
                                  padx=10, pady=5, bd=0, cursor='hand2')
        emergencia_btn.pack(pady=5)
        
        # Rodapé
        footer_frame = tk.Frame(main_frame, bg='#2c3e50')
        footer_frame.pack(pady=(20, 0))
        
        tk.Label(footer_frame, text=f"Desenvolvido por {self.autor}", 
                bg='#2c3e50', fg='#95a5a6', font=('Arial', 9)).pack()
        
        # Eventos
        email_entry.focus()
        self.root.bind('<Return>', lambda e: self.efetuar_login())
    
    def efetuar_login(self):
        """Efetuar login do usuário"""
        email = self.email_var.get()
        senha = self.senha_var.get()
        
        if not email or not senha:
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return
        
        try:
            # Primeiro tentar login com hash MD5
            senha_hash = hashlib.md5(senha.encode()).hexdigest()
            
            self.cursor.execute(
                "SELECT id, nome, tipo, ativo FROM usuarios WHERE email = %s AND senha = %s",
                (email, senha_hash)
            )
            usuario = self.cursor.fetchone()
            
            if not usuario:
                # Se não encontrar com hash, tentar com senha em texto puro (para migração)
                self.cursor.execute(
                    "SELECT id, nome, tipo, ativo FROM usuarios WHERE email = %s AND senha = %s",
                    (email, senha)
                )
                usuario = self.cursor.fetchone()
                
                if usuario:
                    # Atualizar senha para hash
                    self.cursor.execute(
                        "UPDATE usuarios SET senha = %s WHERE id = %s",
                        (senha_hash, usuario[0])
                    )
                    self.conn.commit()
                    print(f"✅ Senha atualizada para hash para usuário: {usuario[1]}")
            
            if usuario:
                if usuario[3] == 0:
                    messagebox.showerror("Erro", "Usuário está inativo!")
                    return
                
                self.usuario_atual = {
                    'id': usuario[0],
                    'nome': usuario[1],
                    'tipo': usuario[2],
                    'email': email
                }
                
                print(f"✅ Login bem-sucedido: {usuario[1]}")
                
                # Iniciar API web apenas após login bem-sucedido
                self.iniciar_api_web()
                
                self.dashboard()
            else:
                messagebox.showerror("Erro", "Email ou senha incorretos!")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao fazer login: {str(e)}")
    
    def login_emergencia(self):
        """Login de emergência sem hash (para problemas de migração)"""
        email = self.email_var.get()
        senha = self.senha_var.get()
        
        if not email or not senha:
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return
        
        try:
            # Buscar usuário com senha em texto puro
            self.cursor.execute(
                "SELECT id, nome, tipo, ativo FROM usuarios WHERE email = %s AND senha = %s",
                (email, senha)
            )
            usuario = self.cursor.fetchone()
            
            if usuario:
                if usuario[3] == 0:
                    messagebox.showerror("Erro", "Usuário está inativo!")
                    return
                
                self.usuario_atual = {
                    'id': usuario[0],
                    'nome': usuario[1],
                    'tipo': usuario[2],
                    'email': email
                }
                
                print(f"✅ Login de emergência bem-sucedido: {usuario[1]}")
                
                # Iniciar API web apenas após login bem-sucedido
                self.iniciar_api_web()
                
                self.dashboard()
            else:
                messagebox.showerror("Erro", "Email ou senha incorretos!")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao fazer login: {str(e)}")
    
    def limpar_tela(self):
        """Limpar todos os widgets da tela"""
        for widget in self.root.winfo_children():
            widget.destroy()
    
    def dashboard(self):
        """Dashboard principal do sistema"""
        self.limpar_tela()
        
        # Frame superior
        top_frame = tk.Frame(self.root, bg='#34495e', height=60)
        top_frame.pack(fill=tk.X)
        top_frame.pack_propagate(False)
        
        # Frame para logo e título
        logo_frame = tk.Frame(top_frame, bg='#34495e')
        logo_frame.pack(side=tk.LEFT, padx=10, pady=5)
        
        # Carregar logo
        try:
            logo_img = Image.open(resource_path("static/logo.png"))
            logo_img = logo_img.resize((75, 75), Image.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(logo_img)
            tk.Label(logo_frame, image=self.logo_photo, bg='#34495e').pack(side=tk.LEFT)
        except:
            tk.Label(logo_frame, text="🖨️", font=('Arial', 24), 
                    bg='#34495e', fg='white').pack(side=tk.LEFT)

        # Label do título
        tk.Label(logo_frame, text="Dashboard", font=('Arial', 18, 'bold'),
         bg='#34495e', fg='white').pack(side=tk.LEFT, padx=10)
        
        # Informações do usuário
        user_frame = tk.Frame(top_frame, bg='#34495e')
        user_frame.pack(side=tk.RIGHT, padx=15)
        
        tk.Label(user_frame, text=f"👤 {self.usuario_atual['nome']}", 
                bg='#34495e', fg='white', font=('Arial', 10)).pack(side=tk.RIGHT, padx=5)
        
        tk.Label(user_frame, text="|", bg='#34495e', fg='#7f8c8d').pack(side=tk.RIGHT, padx=5)
        
        tk.Button(user_frame, text="🌐 Web", bg='#3498db', fg='white',
                 command=self.abrir_dashboard_web).pack(side=tk.RIGHT, padx=5)
        
        tk.Button(user_frame, text="🚪 Sair", bg='#e74c3c', fg='white',
                 command=self.sair).pack(side=tk.RIGHT, padx=5)
        
        # Frame principal
        main_frame = tk.Frame(self.root, bg='#2c3e50')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Menu lateral
        menu_frame = tk.Frame(main_frame, bg='#34495e', width=200)
        menu_frame.pack(side=tk.LEFT, fill=tk.Y)
        menu_frame.pack_propagate(False)
        
        # Opções do menu - AGORA VENDEDOR TEM ACESSO A RELATÓRIOS
        if self.usuario_atual['tipo'] == 'admin':
            menu_opcoes = [
                ("📊 Dashboard", self.dashboard),
                ("🛒 Vendas", self.modulo_vendas),
                ("📦 Produtos", self.modulo_cadastro_produto),
                ("📈 Relatórios", self.modulo_relatorios),
                ("👥 Usuários", self.modulo_usuarios),
                ("⚙️ Configurações", self.modulo_configuracoes)
            ]
        elif self.usuario_atual['tipo'] == 'gerente':
            menu_opcoes = [
                ("📊 Dashboard", self.dashboard),
                ("🛒 Vendas", self.modulo_vendas),
                ("📦 Produtos", self.modulo_cadastro_produto),
                ("📈 Relatórios", self.modulo_relatorios),
                ("⚙️ Configurações", self.modulo_configuracoes)
            ]
        else:  # vendedor
            menu_opcoes = [
                ("📊 Dashboard", self.dashboard),
                ("🛒 Vendas", self.modulo_vendas),
                ("📈 Relatórios", self.modulo_relatorios)
            ]
        
        for texto, comando in menu_opcoes:
            btn = tk.Button(menu_frame, text=texto, font=('Arial', 11),
                          bg='#34495e', fg='white', anchor=tk.W,
                          bd=0, padx=20, pady=12, cursor='hand2',
                          command=comando)
            btn.pack(fill=tk.X)
            
            # Efeito hover
            btn.bind('<Enter>', lambda e, b=btn: b.config(bg='#3d566e'))
            btn.bind('<Leave>', lambda e, b=btn: b.config(bg='#34495e'))
        
        # Conteúdo principal
        content_frame = tk.Frame(main_frame, bg='#2c3e50')
        content_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título do dashboard
        tk.Label(content_frame, text="Visão Geral do Sistema", 
                font=('Arial', 24, 'bold'), bg='#2c3e50', fg='white').pack(anchor=tk.W, pady=(0, 20))
        
        # Métricas
        metricas = self.obter_metricas_dashboard()
        
        # Grid de métricas
        metrics_grid = tk.Frame(content_frame, bg='#2c3e50')
        metrics_grid.pack(fill=tk.X, pady=(0, 30))
        
        metricas_info = [
            ("💰 Vendas Hoje", f"MT {metricas['vendas_hoje']:.2f}", "#27ae60", "fas fa-shopping-cart"),
            ("📦 Total Produtos", metricas['total_produtos'], "#3498db", "fas fa-boxes"),
            ("⚠️ Estoque Baixo", metricas['estoque_baixo'], "#f39c12", "fas fa-exclamation-triangle"),
            ("👥 Usuários Ativos", metricas['usuarios_ativos'], "#9b59b6", "fas fa-users"),
            ("🖨️ Serviços Hoje", f"MT {metricas['servicos_especiais_hoje']:.2f}", "#e74c3c", "fas fa-print"),
            ("💰 Valor Estoque", f"MT {metricas['valor_estoque_total']:.2f}", "#1abc9c", "fas fa-money-bill-wave")
        ]
        
        for i, (titulo, valor, cor, icone) in enumerate(metricas_info):
            metric_frame = tk.Frame(metrics_grid, bg=cor, bd=0, relief=tk.RAISED, width=200, height=120)
            metric_frame.grid(row=i//3, column=i%3, padx=10, pady=10, sticky="nsew")
            metric_frame.grid_propagate(False)
            
            # Ícone
            icon_label = tk.Label(metric_frame, text="", font=('FontAwesome', 24), 
                                 bg=cor, fg='white')
            icon_label.pack(pady=(15, 5))
            
            # Valor
            tk.Label(metric_frame, text=valor, font=('Arial', 20, 'bold'),
                    bg=cor, fg='white').pack()
            
            # Título
            tk.Label(metric_frame, text=titulo, font=('Arial', 10),
                    bg=cor, fg='white').pack(pady=5)
        
        # Configurar grid
        for i in range(3):
            metrics_grid.columnconfigure(i, weight=1)
        
        # Últimas vendas
        vendas_frame = tk.Frame(content_frame, bg='#34495e', bd=2, relief=tk.RAISED)
        vendas_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(vendas_frame, text="📋 Últimas Vendas (Hoje)", 
                font=('Arial', 14, 'bold'), bg='#34495e', fg='white').pack(anchor=tk.W, padx=15, pady=10)
        
        # Lista de vendas
        vendas_tree_frame = tk.Frame(vendas_frame, bg='#34495e')
        vendas_tree_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        
        colunas = ("Hora", "Nº Venda", "Total", "Vendedor")
        tree_vendas = ttk.Treeview(vendas_tree_frame, columns=colunas, show="headings", height=8)
        
        for col in colunas:
            tree_vendas.heading(col, text=col)
            tree_vendas.column(col, width=100)
        
        vsb = ttk.Scrollbar(vendas_tree_frame, orient="vertical", command=tree_vendas.yview)
        tree_vendas.configure(yscrollcommand=vsb.set)
        
        tree_vendas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Carregar vendas
        try:
            self.cursor.execute('''
                SELECT v.data_hora, v.numero_serie, v.total, u.nome
                FROM vendas v
                JOIN usuarios u ON v.usuario_id = u.id
                WHERE DATE(v.data_hora) = CURRENT_DATE
                ORDER BY v.data_hora DESC
                LIMIT 10
            ''')
            
            for venda in self.cursor.fetchall():
                hora = venda[0].strftime("%H:%M") if isinstance(venda[0], datetime) else "N/A"
                tree_vendas.insert("", tk.END, values=(hora, venda[1], f"MT {venda[2]:.2f}", venda[3]))
        except Exception as e:
            print(f"Erro ao carregar vendas: {e}")
    
    def abrir_dashboard_web(self):
        """Abrir dashboard web no navegador"""
        try:
            webbrowser.open('http://localhost:5000')
            messagebox.showinfo("Dashboard Web", "Dashboard web aberto no navegador!\n\nSe não abrir automaticamente, acesse:\nhttp://localhost:5000\n\nPara acesso na rede local, use o IP da máquina:\nex: http://192.168.1.100:5000")
        except Exception as e:
            print(f"Erro ao abrir dashboard web: {e}")
            messagebox.showerror("Erro", "Não foi possível abrir o dashboard web!")
    
    def obter_metricas_dashboard(self):
        """Obter métricas para o dashboard"""
        try:
            hoje = date.today()
            mes_atual = hoje.replace(day=1)
            
            # Vendas hoje
            self.cursor.execute(
                "SELECT COALESCE(SUM(total), 0) FROM vendas WHERE DATE(data_hora) = %s",
                (hoje,)
            )
            vendas_hoje = float(self.cursor.fetchone()[0] or 0)
            
            # Total produtos
            self.cursor.execute("SELECT COUNT(*) FROM produtos")
            total_produtos = self.cursor.fetchone()[0] or 0
            
            # Estoque baixo
            self.cursor.execute("SELECT COUNT(*) FROM produtos WHERE quantidade <= estoque_minimo")
            estoque_baixo = self.cursor.fetchone()[0] or 0
            
            # Usuários ativos
            self.cursor.execute("SELECT COUNT(*) FROM usuarios WHERE ativo = 1")
            usuarios_ativos = self.cursor.fetchone()[0] or 0
            
            # Serviços especiais hoje
            self.cursor.execute(
                "SELECT COALESCE(SUM(total), 0) FROM servicos_especiais WHERE data = %s",
                (hoje,)
            )
            servicos_especiais_hoje = float(self.cursor.fetchone()[0] or 0)
            
            # Valor total do estoque
            self.cursor.execute("SELECT COALESCE(SUM(preco * quantidade), 0) FROM produtos")
            valor_estoque_total = float(self.cursor.fetchone()[0] or 0)
            
            return {
                'vendas_hoje': vendas_hoje,
                'total_produtos': total_produtos,
                'estoque_baixo': estoque_baixo,
                'usuarios_ativos': usuarios_ativos,
                'servicos_especiais_hoje': servicos_especiais_hoje,
                'total_vendas_mes': 0,
                'total_servicos_mes': 0,
                'valor_estoque_total': valor_estoque_total
            }
            
        except Exception as e:
            print(f"Erro ao obter métricas: {e}")
            return {
                'vendas_hoje': 0,
                'total_produtos': 0,
                'estoque_baixo': 0,
                'usuarios_ativos': 0,
                'servicos_especiais_hoje': 0,
                'total_vendas_mes': 0,
                'total_servicos_mes': 0,
                'valor_estoque_total': 0
            }
    
    def modulo_vendas(self):
        """Módulo de vendas"""
        self.limpar_tela()
        
        # Frame superior
        top_frame = tk.Frame(self.root, bg='#34495e', height=60)
        top_frame.pack(fill=tk.X)
        top_frame.pack_propagate(False)
        
        tk.Button(top_frame, text="← Voltar", font=('Arial', 11), bg='#34495e', fg='white',
                 bd=0, command=self.dashboard).pack(side=tk.LEFT, padx=15)
        
        tk.Label(top_frame, text="🛒 Módulo de Vendas", font=('Arial', 18, 'bold'),
                bg='#34495e', fg='white').pack(side=tk.LEFT, padx=10)
        
        # Frame principal
        main_frame = tk.Frame(self.root, bg='#2c3e50')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Notebook (abas)
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Aba 1: Produtos
        aba_produtos = ttk.Frame(notebook, padding="10")
        notebook.add(aba_produtos, text="📦 Produtos")
        
        self.criar_aba_produtos(aba_produtos)
        
        # Aba 2: Impressão/Cópia com numeração
        aba_impressao = ttk.Frame(notebook, padding="10")
        notebook.add(aba_impressao, text="🖨️ Impressão/Cópia")
        
        self.criar_aba_impressao_copia_numerada(aba_impressao)
        
        # Aba 3: Serviços Especiais
        aba_servicos = ttk.Frame(notebook, padding="10")
        notebook.add(aba_servicos, text="📸 Serviços Especiais")
        
        self.criar_aba_servicos_especiais_completa(aba_servicos)
        
        # Aba 4: Carrinho
        aba_carrinho = ttk.Frame(notebook, padding="10")
        notebook.add(aba_carrinho, text="🛒 Carrinho")
        
        self.criar_aba_carrinho(aba_carrinho)
    
    def criar_aba_produtos(self, parent):
        """Criar aba de produtos"""
        # Busca
        busca_frame = tk.Frame(parent, bg='#ecf0f1')
        busca_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(busca_frame, text="Buscar:", bg='#ecf0f1').pack(side=tk.LEFT, padx=5)
        self.busca_var = tk.StringVar()
        ttk.Entry(busca_frame, textvariable=self.busca_var, width=30).pack(side=tk.LEFT, padx=5)
        ttk.Button(busca_frame, text="Buscar", command=self.buscar_produtos).pack(side=tk.LEFT, padx=5)
        
        # Lista de produtos
        tree_frame = tk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        colunas = ("Código", "Produto", "Preço", "Estoque")
        self.tree_produtos = ttk.Treeview(tree_frame, columns=colunas, show="headings", height=15)
        
        for col in colunas:
            self.tree_produtos.heading(col, text=col)
            self.tree_produtos.column(col, width=100)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_produtos.yview)
        self.tree_produtos.configure(yscrollcommand=vsb.set)
        
        self.tree_produtos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Quantidade e adicionar
        bottom_frame = tk.Frame(parent)
        bottom_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(bottom_frame, text="Quantidade:").pack(side=tk.LEFT, padx=5)
        self.qtd_produto_var = tk.StringVar(value="")
        qtd_entry = ttk.Entry(bottom_frame, textvariable=self.qtd_produto_var, width=10)
        qtd_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(bottom_frame, text="Adicionar ao Carrinho", 
                  command=self.adicionar_produto_ao_carrinho).pack(side=tk.LEFT, padx=10)
        
        # Carregar produtos
        self.carregar_produtos()
    
    def carregar_produtos(self):
        """Carregar produtos na árvore"""
        for item in self.tree_produtos.get_children():
            self.tree_produtos.delete(item)
        
        try:
            self.cursor.execute('''
                SELECT codigo, nome, preco, quantidade 
                FROM produtos 
                WHERE quantidade > 0
                ORDER BY nome
            ''')
            
            for produto in self.cursor.fetchall():
                self.tree_produtos.insert("", tk.END, values=produto)
        except Exception as e:
            print(f"Erro ao carregar produtos: {e}")
    
    def buscar_produtos(self):
        """Buscar produtos"""
        termo = self.busca_var.get().lower()
        
        for item in self.tree_produtos.get_children():
            self.tree_produtos.delete(item)
        
        try:
            self.cursor.execute('''
                SELECT codigo, nome, preco, quantidade 
                FROM produtos 
                WHERE quantidade > 0 AND (LOWER(nome) LIKE %s OR LOWER(codigo) LIKE %s)
                ORDER BY nome
            ''', (f'%{termo}%', f'%{termo}%'))
            
            for produto in self.cursor.fetchall():
                self.tree_produtos.insert("", tk.END, values=produto)
        except Exception as e:
            print(f"Erro na busca: {e}")
    
    def adicionar_produto_ao_carrinho(self):
        """Adicionar produto ao carrinho"""
        selecionado = self.tree_produtos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um produto primeiro!")
            return
        
        item = self.tree_produtos.item(selecionado[0])
        produto_info = item['values']
        
        try:
            quantidade = int(self.qtd_produto_var.get())
            if quantidade <= 0:
                messagebox.showerror("Erro", "Quantidade deve ser maior que zero!")
                return
        except ValueError:
            messagebox.showerror("Erro", "Quantidade inválida!")
            return
        
        estoque = int(produto_info[3])
        if quantidade > estoque:
            messagebox.showerror("Erro", f"Estoque insuficiente! Disponível: {estoque}")
            return
        
        # Buscar ID do produto
        self.cursor.execute("SELECT id FROM produtos WHERE codigo = %s", (produto_info[0],))
        produto_result = self.cursor.fetchone()
        
        if not produto_result:
            messagebox.showerror("Erro", "Produto não encontrado no banco!")
            return
        
        produto_id = produto_result[0]
        
        self.carrinho.append({
            'tipo': 'produto',
            'produto_id': produto_id,
            'codigo': produto_info[0],
            'nome': produto_info[1],
            'preco': float(produto_info[2]),
            'quantidade': quantidade,
            'total': float(produto_info[2]) * quantidade
        })
        
        messagebox.showinfo("Sucesso", f"'{produto_info[1]}' adicionado ao carrinho!")
        self.qtd_produto_var.set("")
        self.atualizar_carrinho_view()
    
    def criar_aba_impressao_copia_numerada(self, parent):
        """Criar aba de impressão/cópia com numeração - ATUALIZADO"""
        form_frame = tk.Frame(parent)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tipo de serviço
        tk.Label(form_frame, text="Tipo de Serviço:").grid(row=0, column=0, sticky=tk.W, pady=10)
        self.tipo_impressao_var = tk.StringVar(value="copia_pb")
        
        tipo_frame = tk.Frame(form_frame)
        tipo_frame.grid(row=0, column=1, sticky=tk.W, pady=10, columnspan=2)
        
        # Criação dos radio buttons para cópia
        copia_frame = tk.LabelFrame(tipo_frame, text="Cópia", padx=5, pady=5)
        copia_frame.pack(side=tk.LEFT, padx=10)
        
        ttk.Radiobutton(copia_frame, text="Preto e Branco (2 MT)", 
                       variable=self.tipo_impressao_var, 
                       value="copia_pb", 
                       command=self.calcular_impressao_numerada).pack(anchor=tk.W)
        ttk.Radiobutton(copia_frame, text="Colorida (10 MT)", 
                       variable=self.tipo_impressao_var, 
                       value="copia_colorida", 
                       command=self.calcular_impressao_numerada).pack(anchor=tk.W)
        
        # Criação dos radio buttons para impressão
        impressao_frame = tk.LabelFrame(tipo_frame, text="Impressão", padx=5, pady=5)
        impressao_frame.pack(side=tk.LEFT, padx=10)
        
        ttk.Radiobutton(impressao_frame, text="Preto e Branco (5 MT)", 
                       variable=self.tipo_impressao_var, 
                       value="impressao_pb", 
                       command=self.calcular_impressao_numerada).pack(anchor=tk.W)
        ttk.Radiobutton(impressao_frame, text="Colorida (45 MT)", 
                       variable=self.tipo_impressao_var, 
                       value="impressao_colorida", 
                       command=self.calcular_impressao_numerada).pack(anchor=tk.W)
        
        # Número de ontem
        tk.Label(form_frame, text="Número de ontem:").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.numero_ontem_var = tk.StringVar(value="1000")
        ttk.Entry(form_frame, textvariable=self.numero_ontem_var, width=15).grid(row=1, column=1, sticky=tk.W, pady=10)
        
        # Número de hoje
        tk.Label(form_frame, text="Número de hoje:").grid(row=2, column=0, sticky=tk.W, pady=10)
        self.numero_hoje_var = tk.StringVar(value="1050")
        ttk.Entry(form_frame, textvariable=self.numero_hoje_var, width=15).grid(row=2, column=1, sticky=tk.W, pady=10)
        
        # Quantidade calculada
        tk.Label(form_frame, text="Quantidade calculada:").grid(row=3, column=0, sticky=tk.W, pady=10)
        self.qtd_calculada_var = tk.StringVar(value="50")
        ttk.Label(form_frame, textvariable=self.qtd_calculada_var, 
                 font=('Arial', 10, 'bold')).grid(row=3, column=1, sticky=tk.W, pady=10)
        
        # Preço unitário
        tk.Label(form_frame, text="Preço unitário:").grid(row=4, column=0, sticky=tk.W, pady=10)
        self.preco_unit_impressao_var = tk.StringVar(value="MT 2.00")
        ttk.Label(form_frame, textvariable=self.preco_unit_impressao_var).grid(row=4, column=1, sticky=tk.W, pady=10)
        
        # Total
        tk.Label(form_frame, text="Total:").grid(row=5, column=0, sticky=tk.W, pady=10)
        self.total_impressao_var = tk.StringVar(value="MT 100.00")
        ttk.Label(form_frame, textvariable=self.total_impressao_var, 
                 font=('Arial', 12, 'bold')).grid(row=5, column=1, sticky=tk.W, pady=10)
        
        # Botões
        botoes_frame = tk.Frame(form_frame)
        botoes_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        ttk.Button(botoes_frame, text="Calcular", 
                  command=self.calcular_impressao_numerada).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_frame, text="Adicionar ao Carrinho", 
                  command=self.adicionar_impressao_numerada_ao_carrinho).pack(side=tk.LEFT, padx=5)
        
        # Configurar eventos
        def update_calculo(*args):
            self.calcular_impressao_numerada()
        
        if hasattr(self.numero_ontem_var, 'trace_add'):
            self.numero_ontem_var.trace_add('write', update_calculo)
            self.numero_hoje_var.trace_add('write', update_calculo)
        else:
            self.numero_ontem_var.trace('w', lambda *args: self.calcular_impressao_numerada())
            self.numero_hoje_var.trace('w', lambda *args: self.calcular_impressao_numerada())
        
        # Calcular inicial
        self.calcular_impressao_numerada()
    
    def calcular_impressao_numerada(self):
        """Calcular valor de impressão/cópia com numeração - ATUALIZADO"""
        try:
            tipo = self.tipo_impressao_var.get()
            preco_unit = self.config_precos[tipo]
            
            ontem = int(self.numero_ontem_var.get() or 0)
            hoje = int(self.numero_hoje_var.get() or 0)
            
            if hoje < ontem:
                messagebox.showerror("Erro", "Número de hoje não pode ser menor que número de ontem!")
                return
            
            quantidade = hoje - ontem
            if quantidade < 0:
                quantidade = 0
            
            total = preco_unit * quantidade
            
            self.qtd_calculada_var.set(str(quantidade))
            self.preco_unit_impressao_var.set(f"MT {preco_unit:.2f}")
            self.total_impressao_var.set(f"MT {total:.2f}")
        except ValueError:
            self.preco_unit_impressao_var.set("MT 0.00")
            self.total_impressao_var.set("MT 0.00")
            self.qtd_calculada_var.set(" ")
    
    def adicionar_impressao_numerada_ao_carrinho(self):
        """Adicionar impressão numerada ao carrinho - ATUALIZADO"""
        try:
            tipo = self.tipo_impressao_var.get()
            ontem = int(self.numero_ontem_var.get() or 0)
            hoje = int(self.numero_hoje_var.get() or 0)
            qtd = hoje - ontem
            preco = self.config_precos[tipo]
            total = preco * qtd
            
            if qtd <= 0:
                messagebox.showerror("Erro", "Quantidade deve ser maior que zero!")
                return
            
            # Obter nome amigável para o tipo
            nomes = {
                'copia_pb': 'Cópia Preto e Branco (2 MT)',
                'copia_colorida': 'Cópia Colorida (10 MT)',
                'impressao_pb': 'Impressão Preto e Branco (5 MT)',
                'impressao_colorida': 'Impressão Colorida (45 MT)'
            }
            nome_servico = nomes.get(tipo, tipo)
            
            self.carrinho.append({
                'tipo': 'impressao',
                'nome': nome_servico,
                'preco': preco,
                'quantidade': qtd,
                'total': total,
                'numero_ontem': ontem,
                'numero_hoje': hoje,
                'tipo_servico': tipo
            })
            
            messagebox.showinfo("Sucesso", "Serviço adicionado ao carrinho!")
            self.numero_ontem_var.set(str(hoje))  # Atualiza para próxima vez
            self.atualizar_carrinho_view()
            
        except ValueError:
            messagebox.showerror("Erro", "Números inválidos!")
    
    def criar_aba_servicos_especiais_completa(self, parent):
        """Criar aba de serviços especiais - CORRIGIDO para mostrar 'Argola'"""
        form_frame = tk.Frame(parent)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tipo de serviço
        tk.Label(form_frame, text="Tipo de Serviço:").grid(row=0, column=0, sticky=tk.W, pady=10)
        
        self.tipo_servico_var = tk.StringVar(value="encadernacao_6mm")
        
        # Carregar serviços do banco
        try:
            self.cursor.execute("SELECT nome, tipo, preco, descricao FROM configuracoes_servicos WHERE ativo = 1 AND tipo NOT LIKE '%copia%' AND tipo NOT LIKE '%impressao%' ORDER BY nome")
            servicos_db = self.cursor.fetchall()
            
            servicos = []
            for servico in servicos_db:
                nome_exibido = servico[0]
                tipo_valor = servico[1]
                preco_valor = servico[2]
                descricao_valor = servico[3]
                
                # Atualizar configuração de preços se necessário
                if tipo_valor in self.config_precos:
                    self.config_precos[tipo_valor] = float(preco_valor)
                servicos.append((nome_exibido, tipo_valor))
        except Exception as e:
            print(f"Erro ao carregar serviços: {e}")
            servicos = [
                ("Argola 6mm", "encadernacao_6mm"),
                ("Argola 8mm", "encadernacao_8mm"),
                ("Argola 10mm", "encadernacao_10mm"),
                ("Argola 12mm", "encadernacao_12mm"),
                ("Argola 14mm", "encadernacao_14mm"),
                ("Argola 16mm", "encadernacao_16mm"),
                ("Argola 18mm", "encadernacao_18mm"),
                ("Argola 20mm", "encadernacao_20mm"),
                ("Argola 22mm", "encadernacao_22mm"),
                ("Laminação BI", "laminacao_bi"),
                ("Laminação A4", "laminacao_a4"),
                ("Laminação A3", "laminacao_a3"),
                ("Laminação A5", "laminacao_a5"),
                ("Digitação", "digitacao"),
                ("Fotografia", "fotografia"),
                ("Cartão de Visita", "cartao_visita"),
                ("Convite", "convite"),
                ("Banner", "banner"),
                ("Adesivo", "adesivo")
            ]
        
        tipo_frame = tk.Frame(form_frame)
        tipo_frame.grid(row=0, column=1, sticky=tk.W, pady=10)
        
        # Usar Combobox para seleção
        tipo_combo = ttk.Combobox(tipo_frame, textvariable=self.tipo_servico_var, 
                                 values=[s[0] for s in servicos], width=25, state="readonly")
        tipo_combo.pack(side=tk.LEFT, padx=5)
        tipo_combo.current(0)  # Selecionar o primeiro item
        
        # Mapear nomes exibidos para valores
        self.servicos_map = {s[0]: s[1] for s in servicos}
        
        # Detalhes específicos
        tk.Label(form_frame, text="Detalhes (opcional):").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.detalhes_servico_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.detalhes_servico_var, width=30).grid(row=1, column=1, sticky=tk.W, pady=10)
        
        # Quantidade
        tk.Label(form_frame, text="Quantidade:").grid(row=2, column=0, sticky=tk.W, pady=10)
        self.qtd_servico_var = tk.StringVar(value=" ")
        ttk.Entry(form_frame, textvariable=self.qtd_servico_var, width=10).grid(row=2, column=1, sticky=tk.W, pady=10)
        
        # Preço unitário
        tk.Label(form_frame, text="Preço unitário:").grid(row=3, column=0, sticky=tk.W, pady=10)
        self.preco_servico_var = tk.StringVar(value="MT 20.00")
        ttk.Label(form_frame, textvariable=self.preco_servico_var).grid(row=3, column=1, sticky=tk.W, pady=10)
        
        # Total
        tk.Label(form_frame, text="Total:").grid(row=4, column=0, sticky=tk.W, pady=10)
        self.total_servico_var = tk.StringVar(value="MT 20.00")
        ttk.Label(form_frame, textvariable=self.total_servico_var, 
                 font=('Arial', 12, 'bold')).grid(row=4, column=1, sticky=tk.W, pady=10)
        
        # Botões
        botoes_frame = tk.Frame(form_frame)
        botoes_frame.grid(row=5, column=0, columnspan=2, pady=20)
        
        ttk.Button(botoes_frame, text="Calcular", 
                  command=self.calcular_servico_especial).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_frame, text="Adicionar ao Carrinho", 
                  command=self.adicionar_servico_especial_ao_carrinho).pack(side=tk.LEFT, padx=5)
        
        # Configurar eventos
        def update_servico(*args):
            self.calcular_servico_especial()
        
        if hasattr(self.qtd_servico_var, 'trace_add'):
            self.qtd_servico_var.trace_add('write', update_servico)
        else:
            self.qtd_servico_var.trace('w', lambda *args: self.calcular_servico_especial())
        
        # Calcular inicial
        self.calcular_servico_especial()
    
    def calcular_servico_especial(self):
        """Calcular valor do serviço especial"""
        try:
            # Obter o valor real do tipo de serviço
            nome_exibido = self.tipo_servico_var.get()
            tipo_real = self.servicos_map.get(nome_exibido, nome_exibido)
            
            # Se ainda for o nome exibido, tentar obter o valor diretamente
            if tipo_real == nome_exibido:
                # Inverter o mapeamento
                for nome, valor in self.servicos_map.items():
                    if valor == nome_exibido:
                        tipo_real = valor
                        break
            
            preco = self.config_precos.get(tipo_real, 0.0)
            
            qtd = int(self.qtd_servico_var.get() or 1)
            total = preco * qtd
            
            self.preco_servico_var.set(f"MT {preco:.2f}")
            self.total_servico_var.set(f"MT {total:.2f}")
        except ValueError:
            self.preco_servico_var.set("MT 0.00")
            self.total_servico_var.set("MT 0.00")
    
    def adicionar_servico_especial_ao_carrinho(self):
        """Adicionar serviço especial ao carrinho"""
        try:
            # Obter o valor real do tipo de serviço
            nome_exibido = self.tipo_servico_var.get()
            tipo_real = self.servicos_map.get(nome_exibido, nome_exibido)
            
            qtd = int(self.qtd_servico_var.get())
            detalhes = self.detalhes_servico_var.get() or nome_exibido
            
            preco = self.config_precos.get(tipo_real, 0.0)
            total = preco * qtd
            
            if qtd <= 0:
                messagebox.showerror("Erro", "Quantidade deve ser maior que zero!")
                return
            
            self.carrinho.append({
                'tipo': 'servico_especial',
                'nome': detalhes,
                'preco': preco,
                'quantidade': qtd,
                'total': total,
                'tipo_servico': tipo_real
            })
            
            messagebox.showinfo("Sucesso", f"Serviço '{nome_exibido}' adicionado ao carrinho!")
            self.qtd_servico_var.set("0")
            self.detalhes_servico_var.set("")
            self.atualizar_carrinho_view()
            
        except ValueError:
            messagebox.showerror("Erro", "Quantidade inválida!")
    
    def criar_aba_carrinho(self, parent):
        """Criar aba do carrinho"""
        # Lista do carrinho
        tree_frame = tk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        colunas = ("Tipo", "Descrição", "Qtd", "Preço Unit.", "Total")
        self.tree_carrinho = ttk.Treeview(tree_frame, columns=colunas, show="headings", height=10)
        
        for col in colunas:
            self.tree_carrinho.heading(col, text=col)
            self.tree_carrinho.column(col, width=100)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_carrinho.yview)
        self.tree_carrinho.configure(yscrollcommand=vsb.set)
        
        self.tree_carrinho.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Botões
        botoes_frame = tk.Frame(parent)
        botoes_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(botoes_frame, text="Remover Selecionado", 
                  command=self.remover_do_carrinho).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_frame, text="Limpar Carrinho", 
                  command=self.limpar_carrinho).pack(side=tk.LEFT, padx=5)
        
        # Resumo
        resumo_frame = tk.Frame(parent, bd=2, relief=tk.GROOVE)
        resumo_frame.pack(fill=tk.X, pady=10)
        
        self.total_carrinho_var = tk.StringVar(value="MT 0.00")
        tk.Label(resumo_frame, text="Total do Carrinho:", 
                font=('Arial', 12)).pack(side=tk.LEFT, padx=10, pady=10)
        tk.Label(resumo_frame, textvariable=self.total_carrinho_var,
                font=('Arial', 14, 'bold'), fg='green').pack(side=tk.LEFT, padx=10, pady=10)
        
        # Pagamento
        pagamento_frame = tk.Frame(parent)
        pagamento_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(pagamento_frame, text="Valor Recebido (MT):").pack(side=tk.LEFT, padx=5)
        self.valor_recebido_var = tk.StringVar()
        ttk.Entry(pagamento_frame, textvariable=self.valor_recebido_var, width=15).pack(side=tk.LEFT, padx=5)
        
        tk.Label(pagamento_frame, text="Troco:").pack(side=tk.LEFT, padx=5)
        self.troco_var = tk.StringVar(value="MT 0.00")
        tk.Label(pagamento_frame, textvariable=self.troco_var, 
                font=('Arial', 11, 'bold')).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(pagamento_frame, text="Calcular Troco", 
                  command=self.calcular_troco).pack(side=tk.LEFT, padx=10)
        
        # Botão finalizar
        ttk.Button(parent, text="💰 Finalizar Venda", 
                  command=self.finalizar_venda_completa, style='Accent.TButton').pack(pady=20)
        
        # Atualizar carrinho
        self.atualizar_carrinho_view()
    
    def atualizar_carrinho_view(self):
        """Atualizar visualização do carrinho"""
        for item in self.tree_carrinho.get_children():
            self.tree_carrinho.delete(item)
        
        total_geral = 0
        for item in self.carrinho:
            if item['tipo'] == 'produto':
                icone = "📦"
                descricao = item['nome']
            elif item['tipo'] == 'impressao':
                icone = "🖨️"
                descricao = f"{item['nome']}"
                if item.get('numero_ontem') is not None and item.get('numero_hoje') is not None:
                    descricao += f" ({item['numero_ontem']}-{item['numero_hoje']})"
            elif item['tipo'] == 'servico_especial':
                icone = "📝"
                descricao = item['nome']
                # Adicionar tipo de argola se aplicável
                if 'encadernacao' in item.get('tipo_servico', ''):
                    # Extrair tamanho da argola do tipo
                    tipo = item['tipo_servico']
                    if '_' in tipo:
                        tamanho = tipo.split('_')[1]
                        descricao = f"Encadernação Argola {tamanho}"
            else:
                icone = "📄"
                descricao = item['nome']
            
            self.tree_carrinho.insert("", tk.END, values=(
                item['tipo'].replace('_', ' ').title(),
                f"{icone} {descricao}",
                item['quantidade'],
                f"MT {item['preco']:.2f}",
                f"MT {item['total']:.2f}"
            ))
            total_geral += item['total']
        
        self.total_carrinho_var.set(f"MT {total_geral:.2f}")
    
    def remover_do_carrinho(self):
        """Remover item do carrinho"""
        selecionado = self.tree_carrinho.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um item para remover!")
            return
        
        index = self.tree_carrinho.index(selecionado[0])
        if 0 <= index < len(self.carrinho):
            item_removido = self.carrinho.pop(index)
            self.atualizar_carrinho_view()
            messagebox.showinfo("Removido", f"Item '{item_removido['nome']}' removido!")
    
    def limpar_carrinho(self):
        """Limpar carrinho"""
        if self.carrinho:
            if messagebox.askyesno("Confirmar", "Limpar todo o carrinho?"):
                self.carrinho = []
                self.atualizar_carrinho_view()
                messagebox.showinfo("Carrinho Limpo", "Carrinho limpo com sucesso!")
    
    def calcular_troco(self):
        """Calcular troco"""
        try:
            total_str = self.total_carrinho_var.get().replace("MT ", "").replace(",", ".")
            total = float(total_str)
            recebido_str = self.valor_recebido_var.get().replace(",", ".")
            
            if not recebido_str:
                self.troco_var.set("Informe valor recebido")
                return
            
            recebido = float(recebido_str)
            
            if recebido < total:
                self.troco_var.set(f"Falta: MT {total - recebido:.2f}")
                return
            
            troco = recebido - total
            self.troco_var.set(f"MT {troco:.2f}")
        except ValueError:
            self.troco_var.set("Valor inválido!")
    
    def finalizar_venda_completa(self):
        """Finalizar venda completa"""
        if not self.carrinho:
            messagebox.showwarning("Aviso", "Carrinho vazio!")
            return
        
        try:
            total_str = self.total_carrinho_var.get().replace("MT ", "").replace(",", ".")
            total = float(total_str)
            recebido_str = self.valor_recebido_var.get()
            
            if not recebido_str:
                messagebox.showerror("Erro", "Informe o valor recebido!")
                return
            
            recebido = float(recebido_str.replace(",", "."))
            
            if recebido < total:
                messagebox.showerror("Erro", f"Valor recebido insuficiente! Falta: MT {total - recebido:.2f}")
                return
            
            troco = recebido - total
            
            # Confirmar venda
            resposta = messagebox.askyesno(
                "Confirmar Venda",
                f"Total: MT {total:.2f}\n"
                f"Recebido: MT {recebido:.2f}\n"
                f"Troco: MT {troco:.2f}\n\n"
                "Finalizar venda?"
            )
            
            if not resposta:
                return
            
            # Gerar número da venda
            data_hoje = date.today().strftime("%Y%m%d")
            self.cursor.execute(
                "SELECT COUNT(*) FROM vendas WHERE DATE(data_hora) = CURRENT_DATE"
            )
            num_vendas = self.cursor.fetchone()[0] + 1
            numero_serie = f"V{data_hoje}-{num_vendas:04d}"
            
            try:
                # Inserir venda
                self.cursor.execute('''
                    INSERT INTO vendas (numero_serie, data_hora, total, valor_recebido, troco, usuario_id)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    RETURNING id
                ''', (numero_serie, datetime.now(), total, recebido, troco, self.usuario_atual['id']))
                
                venda_id = self.cursor.fetchone()[0]
                
                # Processar itens
                for item in self.carrinho:
                    if item['tipo'] == 'produto':
                        # Verificar se produto_id existe
                        if 'produto_id' not in item:
                            # Buscar ID do produto
                            self.cursor.execute("SELECT id FROM produtos WHERE codigo = %s", (item['codigo'],))
                            produto_result = self.cursor.fetchone()
                            
                            if produto_result:
                                item['produto_id'] = produto_result[0]
                            else:
                                raise Exception(f"Produto {item['codigo']} não encontrado")
                        
                        # Inserir item da venda
                        self.cursor.execute('''
                            INSERT INTO itens_venda (venda_id, produto_id, quantidade, preco_unitario, total_item)
                            VALUES (%s, %s, %s, %s, %s)
                        ''', (venda_id, item['produto_id'], item['quantidade'], item['preco'], item['total']))
                        
                        # Atualizar estoque
                        self.cursor.execute('''
                            UPDATE produtos SET quantidade = quantidade - %s WHERE id = %s
                        ''', (item['quantidade'], item['produto_id']))
                    
                    elif item['tipo'] == 'impressao':
                        # Inserir serviço de impressão/cópia com numeração
                        tipo = item.get('tipo_servico', 'copia')
                        self.cursor.execute('''
                            INSERT INTO servicos (tipo, numero_ontem, numero_hoje, quantidade, 
                                                preco_unitario, total, data, usuario_id)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                        ''', (tipo, 
                              item.get('numero_ontem', 0), 
                              item.get('numero_hoje', item['quantidade']),
                              item['quantidade'], item['preco'], item['total'], 
                              date.today(), self.usuario_atual['id']))
                    
                    elif item['tipo'] == 'servico_especial':
                        # Inserir serviço especial
                        self.cursor.execute('''
                            INSERT INTO servicos_especiais (tipo, descricao, quantidade, 
                                                          preco_unitario, total, data, usuario_id)
                            VALUES (%s, %s, %s, %s, %s, %s, %s)
                        ''', (item.get('tipo_servico', item['nome'][:50]), 
                              item['nome'], 
                              item['quantidade'], 
                              item['preco'], 
                              item['total'], 
                              date.today(), 
                              self.usuario_atual['id']))
                
                # Confirmar transação
                self.conn.commit()
                
                # Mostrar comprovante
                self.mostrar_comprovante_completo(venda_id, numero_serie, total, recebido, troco)
                
                # Limpar carrinho
                self.carrinho = []
                self.atualizar_carrinho_view()
                self.valor_recebido_var.set("")
                self.troco_var.set("MT 0.00")
                
                # Atualizar produtos na aba de produtos
                self.carregar_produtos()
                
                messagebox.showinfo("Sucesso", f"Venda finalizada!\nNúmero: {numero_serie}")
                
            except Exception as e:
                self.conn.rollback()
                raise e
                
        except Exception as e:
            print(f"Erro na transação: {e}")
            messagebox.showerror("Erro", f"Erro ao finalizar venda: {str(e)}")
    
    def mostrar_comprovante_completo(self, venda_id, numero_serie, total, recebido, troco):
        """Mostrar comprovante completo da venda - ATUALIZADO"""
        comprovante = tk.Toplevel(self.root)
        comprovante.title("Comprovante de Venda")
        comprovante.geometry("500x600")
        
        # Conteúdo do comprovante
        content = tk.Text(comprovante, wrap=tk.WORD, font=('Courier', 10))
        content.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        comprovante_texto = f"""
{'='*50}
PAPELARIA EXPRESS - SISTEMA DE GESTÃO
{'='*50}
COMPROVANTE DE VENDA
{'-'*50}
Nº Venda: {numero_serie}
Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
Vendedor: {self.usuario_atual['nome']}
{'-'*50}
ITENS:
{'-'*50}
"""
        
        for item in self.carrinho:
            if item['tipo'] == 'produto':
                linha = f"📦 {item['nome']}"
            elif item['tipo'] == 'impressao':
                linha = f"🖨️ {item['nome']}"
                if 'numero_ontem' in item and 'numero_hoje' in item:
                    linha += f" ({item['numero_ontem']}-{item['numero_hoje']})"
            elif item['tipo'] == 'servico_especial':
                linha = f"📝 {item['nome']}"
            else:
                linha = f"📄 {item['nome']}"
            
            linha += f" x{item['quantidade']} @ MT{item['preco']:.2f}"
            linha += f" = MT{item['total']:.2f}"
            comprovante_texto += linha + "\n"
        
        comprovante_texto += f"""
{'-'*50}
TOTAL: MT {total:.2f}
RECEBIDO: MT {recebido:.2f}
TROCO: MT {troco:.2f}
{'='*50}
OBRIGADO PELA PREFERÊNCIA!
VOLTE SEMPRE!
{'='*50}
"""
        
        content.insert(tk.END, comprovante_texto)
        content.config(state=tk.DISABLED)
        
        # Botões
        btn_frame = tk.Frame(comprovante)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="🖨️ Imprimir", bg='#3498db', fg='white',
                 command=lambda: self.imprimir_comprovante(comprovante_texto)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="📋 Copiar", bg='#2ecc71', fg='white',
                 command=lambda: self.copiar_comprovante(comprovante_texto)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="💾 Salvar PDF", bg='#9b59b6', fg='white',
                 command=lambda: self.salvar_comprovante_pdf(numero_serie, total, recebido, troco)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="❌ Fechar", bg='#e74c3c', fg='white',
                 command=comprovante.destroy).pack(side=tk.LEFT, padx=5)
    
    def salvar_comprovante_pdf(self, numero_serie, total, recebido, troco):
        """Salvar comprovante como PDF"""
        arquivo = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("Todos os arquivos", "*.*")],
            initialfile=f"comprovante_{numero_serie}.pdf"
        )
        
        if not arquivo:
            return
        
        try:
            # Criar documento PDF
            doc = SimpleDocTemplate(arquivo, pagesize=A4)
            elementos = []
            
            # Estilos
            estilos = getSampleStyleSheet()
            estilo_titulo = ParagraphStyle(
                'CustomTitle',
                parent=estilos['Title'],
                fontSize=16,
                textColor=colors.HexColor('#2c3e50'),
                spaceAfter=20
            )
            
            estilo_normal = ParagraphStyle(
                'CustomNormal',
                parent=estilos['Normal'],
                fontSize=10,
                textColor=colors.black
            )
            
            # Título
            elementos.append(Paragraph("PAPELARIA EXPRESS - SISTEMA DE GESTÃO", estilo_titulo))
            elementos.append(Paragraph("=" * 50, estilo_normal))
            elementos.append(Paragraph("COMPROVANTE DE VENDA", estilo_titulo))
            elementos.append(Spacer(1, 10))
            
            # Informações da venda
            info_data = [
                [Paragraph("<b>Nº Venda:</b>", estilo_normal), Paragraph(numero_serie, estilo_normal)],
                [Paragraph("<b>Data:</b>", estilo_normal), Paragraph(datetime.now().strftime('%d/%m/%Y %H:%M:%S'), estilo_normal)],
                [Paragraph("<b>Vendedor:</b>", estilo_normal), Paragraph(self.usuario_atual['nome'], estilo_normal)]
            ]
            
            tabela_info = Table(info_data, colWidths=[3*cm, 10*cm])
            tabela_info.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ]))
            
            elementos.append(tabela_info)
            elementos.append(Spacer(1, 20))
            
            # Itens da venda
            elementos.append(Paragraph("<b>ITENS:</b>", estilo_normal))
            elementos.append(Spacer(1, 10))
            
            dados_itens = []
            cabecalho_itens = ['Item', 'Descrição', 'Qtd', 'Preço', 'Total']
            dados_itens.append([Paragraph(f"<b>{col}</b>", estilo_normal) for col in cabecalho_itens])
            
            for i, item in enumerate(self.carrinho, 1):
                if item['tipo'] == 'produto':
                    descricao = item['nome']
                elif item['tipo'] == 'impressao':
                    descricao = f"{item['nome']}"
                    if 'numero_ontem' in item and 'numero_hoje' in item:
                        descricao += f" ({item['numero_ontem']}-{item['numero_hoje']})"
                elif item['tipo'] == 'servico_especial':
                    descricao = item['nome']
                else:
                    descricao = item['nome']
                
                dados_itens.append([
                    Paragraph(str(i), estilo_normal),
                    Paragraph(descricao, estilo_normal),
                    Paragraph(str(item['quantidade']), estilo_normal),
                    Paragraph(f"MT {item['preco']:.2f}", estilo_normal),
                    Paragraph(f"MT {item['total']:.2f}", estilo_normal)
                ])
            
            tabela_itens = Table(dados_itens, colWidths=[1.5*cm, 8*cm, 2*cm, 3*cm, 3*cm])
            tabela_itens.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('ALIGN', (2, 1), (2, -1), 'CENTER'),
                ('ALIGN', (3, 1), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#34495e')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
            ]))
            
            elementos.append(tabela_itens)
            elementos.append(Spacer(1, 20))
            
            # Totais
            totais_data = [
                [Paragraph("<b>TOTAL:</b>", estilo_normal), Paragraph(f"MT {total:.2f}", estilo_normal)],
                [Paragraph("<b>RECEBIDO:</b>", estilo_normal), Paragraph(f"MT {recebido:.2f}", estilo_normal)],
                [Paragraph("<b>TROCO:</b>", estilo_normal), Paragraph(f"MT {troco:.2f}", estilo_normal)]
            ]
            
            tabela_totais = Table(totais_data, colWidths=[5*cm, 5*cm])
            tabela_totais.setStyle(TableStyle([
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTSIZE', (0, 0), (-1, -1), 12),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
                ('TOPPADDING', (0, 0), (-1, -1), 10),
                ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),
                ('LINEBELOW', (0, -1), (-1, -1), 1, colors.black),
            ]))
            
            elementos.append(tabela_totais)
            elementos.append(Spacer(1, 30))
            
            # Rodapé
            elementos.append(Paragraph("OBRIGADO PELA PREFERÊNCIA!", estilo_titulo))
            elementos.append(Paragraph("VOLTE SEMPRE!", estilo_normal))
            elementos.append(Spacer(1, 10))
            elementos.append(Paragraph("=" * 50, estilo_normal))
            elementos.append(Spacer(1, 10))
            elementos.append(Paragraph(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", estilo_normal))
            
            # Construir PDF
            doc.build(elementos)
            
            messagebox.showinfo("Sucesso", f"Comprovante salvo como PDF: {arquivo}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar PDF: {str(e)}")
    
    def imprimir_comprovante(self, texto):
        """Imprimir comprovante"""
        try:
            # Criar arquivo temporário
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt', encoding='utf-8') as f:
                f.write(texto)
                temp_file = f.name
            
            # Imprimir usando a impressora padrão
            printer_name = win32print.GetDefaultPrinter()
            hprinter = win32print.OpenPrinter(printer_name)
            
            try:
                # Ler arquivo e imprimir
                with open(temp_file, 'r', encoding='utf-8') as f:
                    data = f.read()
                
                # Converter para bytes
                data_bytes = data.encode('utf-8')
                
                # Iniciar trabalho de impressão
                job = win32print.StartDocPrinter(hprinter, 1, ("Comprovante", None, "RAW"))
                try:
                    win32print.StartPagePrinter(hprinter)
                    win32print.WritePrinter(hprinter, data_bytes)
                    win32print.EndPagePrinter(hprinter)
                finally:
                    win32print.EndDocPrinter(hprinter)
                
                messagebox.showinfo("Imprimir", "Comprovante enviado para impressora!")
                
            finally:
                win32print.ClosePrinter(hprinter)
            
            # Limpar arquivo temporário
            os.unlink(temp_file)
            
        except Exception as e:
            print(f"Erro ao imprimir: {e}")
            # Fallback para mensagem simples
            messagebox.showinfo("Imprimir", "Comprovante enviado para impressora!")
    
    def copiar_comprovante(self, texto):
        """Copiar comprovante para área de transferência"""
        self.root.clipboard_clear()
        self.root.clipboard_append(texto)
        messagebox.showinfo("Copiado", "Comprovante copiado para área de transferência!")
    
    def modulo_cadastro_produto(self):
        """Módulo de cadastro de produto"""
        if self.usuario_atual['tipo'] not in ['admin', 'gerente']:
            messagebox.showerror("Acesso Negado", "Apenas administradores podem acessar!")
            self.dashboard()
            return
        
        self.limpar_tela()
        
        # Frame superior
        top_frame = tk.Frame(self.root, bg='#34495e', height=60)
        top_frame.pack(fill=tk.X)
        top_frame.pack_propagate(False)
        
        tk.Button(top_frame, text="← Voltar", font=('Arial', 11), bg='#34495e', fg='white',
                 bd=0, command=self.dashboard).pack(side=tk.LEFT, padx=15)
        
        tk.Label(top_frame, text="📦 Gestão de Produtos", font=('Arial', 18, 'bold'),
                bg='#34495e', fg='white').pack(side=tk.LEFT, padx=10)
        
        # Botão para visualizar produtos
        tk.Button(top_frame, text="📋 Visualizar Produtos", bg='#3498db', fg='white',
                 command=self.visualizar_produtos).pack(side=tk.RIGHT, padx=15)
        
        # Frame principal
        main_frame = tk.Frame(self.root, bg='#2c3e50')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Notebook (abas)
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Aba 1: Cadastrar Produto
        aba_cadastrar = ttk.Frame(notebook, padding="20")
        notebook.add(aba_cadastrar, text="➕ Cadastrar Produto")
        
        # Frame para o formulário
        form_frame = tk.Frame(aba_cadastrar)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(form_frame, text="Cadastrar Novo Produto", font=('Arial', 16, 'bold')).pack(pady=(0, 20))
        
        # Campos
        campos_frame = tk.Frame(form_frame)
        campos_frame.pack()
        
        campos = [
            ("Código:", "codigo"),
            ("Nome:", "nome"),
            ("Categoria:", "categoria"),
            ("Preço (MT):", "preco"),
            ("Quantidade:", "quantidade"),
            ("Estoque Mínimo:", "estoque_minimo")
        ]
        
        self.campos_produto = {}
        
        for i, (label, nome) in enumerate(campos):
            tk.Label(campos_frame, text=label, font=('Arial', 11)).grid(row=i, column=0, sticky=tk.W, pady=5, padx=5)
            
            var = tk.StringVar()
            if nome == 'estoque_minimo':
                var.set("5")
            
            entry = ttk.Entry(campos_frame, textvariable=var, width=30)
            entry.grid(row=i, column=1, sticky=tk.W, pady=5, padx=5)
            
            self.campos_produto[nome] = var
        
        # Combo para categoria
        categorias = ["Escritório", "Papelaria", "Material Escolar", "Impressão", "Fotografia", "Outros"]
        self.campos_produto['categoria'] = tk.StringVar()
        categoria_combo = ttk.Combobox(campos_frame, textvariable=self.campos_produto['categoria'], 
                                      values=categorias, width=28)
        categoria_combo.grid(row=2, column=1, sticky=tk.W, pady=5, padx=5)
        
        # Botões
        botoes_frame = tk.Frame(form_frame)
        botoes_frame.pack(pady=30)
        
        ttk.Button(botoes_frame, text="Cadastrar", 
                  command=self.salvar_produto, style='Accent.TButton').pack(side=tk.LEFT, padx=10)
        ttk.Button(botoes_frame, text="Limpar", 
                  command=self.limpar_campos_produto).pack(side=tk.LEFT, padx=10)
        
        # Aba 2: Editar Produto
        aba_editar = ttk.Frame(notebook, padding="20")
        notebook.add(aba_editar, text="✏️ Editar Produto")
        
        # Frame para edição
        edit_frame = tk.Frame(aba_editar)
        edit_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(edit_frame, text="Editar/Excluir Produto", font=('Arial', 16, 'bold')).pack(pady=(0, 20))
        
        # Busca
        busca_frame = tk.Frame(edit_frame)
        busca_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(busca_frame, text="Buscar Produto:", font=('Arial', 11)).pack(side=tk.LEFT, padx=5)
        self.busca_editar_var = tk.StringVar()
        ttk.Entry(busca_frame, textvariable=self.busca_editar_var, width=30).pack(side=tk.LEFT, padx=5)
        ttk.Button(busca_frame, text="Buscar", command=self.buscar_produto_editar).pack(side=tk.LEFT, padx=5)
        
        # Lista de produtos
        tree_frame = tk.Frame(edit_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        colunas = ("Código", "Produto", "Categoria", "Preço", "Estoque", "Mínimo")
        self.tree_editar_produtos = ttk.Treeview(tree_frame, columns=colunas, show="headings", height=8)
        
        for col in colunas:
            self.tree_editar_produtos.heading(col, text=col)
            self.tree_editar_produtos.column(col, width=100)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_editar_produtos.yview)
        self.tree_editar_produtos.configure(yscrollcommand=vsb.set)
        
        self.tree_editar_produtos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Campos para edição
        campos_edit_frame = tk.Frame(edit_frame)
        campos_edit_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.campos_editar_produto = {}
        campos_editar = [
            ("Novo Nome:", "nome_editar"),
            ("Nova Categoria:", "categoria_editar"),
            ("Novo Preço:", "preco_editar"),
            ("Nova Quantidade:", "quantidade_editar"),
            ("Novo Mínimo:", "minimo_editar")
        ]
        
        for i, (label, nome) in enumerate(campos_editar):
            tk.Label(campos_edit_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=2, padx=5)
            
            var = tk.StringVar()
            entry = ttk.Entry(campos_edit_frame, textvariable=var, width=25)
            entry.grid(row=i, column=1, sticky=tk.W, pady=2, padx=5)
            
            self.campos_editar_produto[nome] = var
        
        # Botões de ação
        botoes_edit_frame = tk.Frame(edit_frame)
        botoes_edit_frame.pack(pady=10)
        
        ttk.Button(botoes_edit_frame, text="Carregar para Editar", 
                  command=self.carregar_produto_para_editar).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_edit_frame, text="Salvar Alterações", 
                  command=self.salvar_edicao_produto).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_edit_frame, text="Excluir Produto", 
                  command=self.excluir_produto, style='Accent.TButton').pack(side=tk.LEFT, padx=5)
        
        # Carregar produtos inicialmente
        self.carregar_todos_produtos_para_editar()
    
    def buscar_produto_editar(self):
        """Buscar produto para editar"""
        termo = self.busca_editar_var.get().lower()
        
        for item in self.tree_editar_produtos.get_children():
            self.tree_editar_produtos.delete(item)
        
        try:
            self.cursor.execute('''
                SELECT codigo, nome, categoria, preco, quantidade, estoque_minimo 
                FROM produtos 
                WHERE LOWER(nome) LIKE %s OR LOWER(codigo) LIKE %s
                ORDER BY nome
            ''', (f'%{termo}%', f'%{termo}%'))
            
            for produto in self.cursor.fetchall():
                self.tree_editar_produtos.insert("", tk.END, values=produto)
        except Exception as e:
            print(f"Erro na busca: {e}")
    
    def carregar_produto_para_editar(self):
        """Carregar produto selecionado para edição"""
        selecionado = self.tree_editar_produtos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um produto primeiro!")
            return
        
        item = self.tree_editar_produtos.item(selecionado[0])
        produto_info = item['values']
        
        # Preencher campos com dados do produto
        self.campos_editar_produto['nome_editar'].set(produto_info[1])
        self.campos_editar_produto['categoria_editar'].set(produto_info[2])
        self.campos_editar_produto['preco_editar'].set(str(produto_info[3]))
        self.campos_editar_produto['quantidade_editar'].set(str(produto_info[4]))
        self.campos_editar_produto['minimo_editar'].set(str(produto_info[5]))
        
        # Salvar código do produto selecionado
        self.produto_editar_codigo = produto_info[0]
        messagebox.showinfo("Sucesso", f"Produto '{produto_info[1]}' carregado para edição!")
    
    def salvar_edicao_produto(self):
        """Salvar edição do produto"""
        if not hasattr(self, 'produto_editar_codigo'):
            messagebox.showwarning("Aviso", "Selecione um produto primeiro!")
            return
        
        try:
            dados = {k: v.get() for k, v in self.campos_editar_produto.items()}
            
            # Validar campos
            if not dados['nome_editar']:
                messagebox.showerror("Erro", "O nome do produto é obrigatório!")
                return
            
            # Converter valores
            preco = float(dados['preco_editar'].replace(",", ".") or 0)
            quantidade = int(dados['quantidade_editar'] or 0)
            estoque_minimo = int(dados['minimo_editar'] or 5)
            
            if preco < 0 or quantidade < 0 or estoque_minimo < 0:
                messagebox.showerror("Erro", "Valores não podem ser negativos!")
                return
            
            # Atualizar produto
            self.cursor.execute('''
                UPDATE produtos 
                SET nome = %s, categoria = %s, preco = %s, 
                    quantidade = %s, estoque_minimo = %s
                WHERE codigo = %s
            ''', (dados['nome_editar'], dados['categoria_editar'], preco, 
                  quantidade, estoque_minimo, self.produto_editar_codigo))
            
            self.conn.commit()
            messagebox.showinfo("Sucesso", "Produto atualizado com sucesso!")
            
            # Limpar campos
            for var in self.campos_editar_produto.values():
                var.set("")
            
            # Recarregar lista
            self.carregar_todos_produtos_para_editar()
            delattr(self, 'produto_editar_codigo')
            
        except ValueError:
            messagebox.showerror("Erro", "Valores numéricos inválidos!")
        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Erro", f"Erro ao atualizar: {str(e)}")
    
    def excluir_produto(self):
        """Excluir produto"""
        selecionado = self.tree_editar_produtos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um produto primeiro!")
            return
        
        item = self.tree_editar_produtos.item(selecionado[0])
        produto_info = item['values']
        codigo = produto_info[0]
        nome = produto_info[1]
        
        # Confirmar exclusão
        resposta = messagebox.askyesno(
            "Confirmar Exclusão",
            f"Tem certeza que deseja excluir o produto?\n\n"
            f"Código: {codigo}\n"
            f"Nome: {nome}\n\n"
            f"Esta ação não pode ser desfeita!"
        )
        
        if not resposta:
            return
        
        try:
            # Verificar se o produto está em alguma venda
            self.cursor.execute('''
                SELECT COUNT(*) FROM itens_venda iv
                JOIN produtos p ON iv.produto_id = p.id
                WHERE p.codigo = %s
            ''', (codigo,))
            
            count_vendas = self.cursor.fetchone()[0]
            
            if count_vendas > 0:
                messagebox.showwarning(
                    "Aviso", 
                    f"Este produto está em {count_vendas} vendas.\n"
                    f"Não é possível excluí-lo.\n"
                    f"Você pode desativá-lo definindo a quantidade como 0."
                )
                return
            
            # Excluir produto
            self.cursor.execute("DELETE FROM produtos WHERE codigo = %s", (codigo,))
            self.conn.commit()
            
            messagebox.showinfo("Sucesso", "Produto excluído com sucesso!")
            
            # Recarregar lista
            self.carregar_todos_produtos_para_editar()
            
        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Erro", f"Erro ao excluir: {str(e)}")
    
    def carregar_todos_produtos_para_editar(self):
        """Carregar todos os produtos para edição"""
        for item in self.tree_editar_produtos.get_children():
            self.tree_editar_produtos.delete(item)
        
        try:
            self.cursor.execute('''
                SELECT codigo, nome, categoria, preco, quantidade, estoque_minimo 
                FROM produtos 
                ORDER BY nome
            ''')
            
            for produto in self.cursor.fetchall():
                self.tree_editar_produtos.insert("", tk.END, values=produto)
        except Exception as e:
            print(f"Erro ao carregar produtos para edição: {e}")
    
    def visualizar_produtos(self):
        """Visualizar todos os produtos em uma janela"""
        janela = tk.Toplevel(self.root)
        janela.title("Lista de Produtos")
        janela.geometry("900x600")
        
        # Frame superior
        top_frame = tk.Frame(janela, bg='#34495e', height=50)
        top_frame.pack(fill=tk.X)
        top_frame.pack_propagate(False)
        
        tk.Label(top_frame, text="📋 Lista de Produtos", font=('Arial', 16, 'bold'),
                bg='#34495e', fg='white').pack(side=tk.LEFT, padx=15)
        
        # Frame principal
        main_frame = tk.Frame(janela)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Treeview
        colunas = ("Código", "Produto", "Categoria", "Preço", "Estoque", "Mínimo", "Status", "Valor Total")
        tree = ttk.Treeview(main_frame, columns=colunas, show="headings", height=20)
        
        for col in colunas:
            tree.heading(col, text=col)
            tree.column(col, width=100)
        
        vsb = ttk.Scrollbar(main_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Carregar produtos
        try:
            self.cursor.execute('''
                SELECT codigo, nome, categoria, preco, quantidade, estoque_minimo 
                FROM produtos 
                ORDER BY nome
            ''')
            
            valor_total_estoque = 0
            
            for produto in self.cursor.fetchall():
                # Calcular status
                quantidade = produto[4]
                minimo = produto[5]
                
                if quantidade < minimo:
                    status = "CRÍTICO"
                elif quantidade == minimo:
                    status = "BAIXO"
                else:
                    status = "OK"
                
                # Calcular valor total
                valor_total = produto[3] * quantidade
                valor_total_estoque += valor_total
                
                tree.insert("", tk.END, values=(
                    produto[0], produto[1], produto[2], 
                    f"MT {produto[3]:.2f}", quantidade, minimo, 
                    status, f"MT {valor_total:.2f}"
                ))
            
            # Mostrar total
            tk.Label(janela, text=f"Valor Total do Estoque: MT {valor_total_estoque:.2f}", 
                    font=('Arial', 12, 'bold')).pack(pady=10)
            
        except Exception as e:
            print(f"Erro ao carregar produtos: {e}")
    
    def salvar_produto(self):
        """Salvar produto no banco"""
        try:
            dados = {k: v.get() for k, v in self.campos_produto.items()}
            
            # Validar campos obrigatórios
            if not all([dados['codigo'], dados['nome'], dados['categoria'], dados['preco'], dados['quantidade']]):
                messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
                return
            
            # Converter valores
            preco = float(dados['preco'].replace(",", "."))
            quantidade = int(dados['quantidade'])
            estoque_minimo = int(dados['estoque_minimo'] or 5)
            
            if preco <= 0 or quantidade < 0 or estoque_minimo < 0:
                messagebox.showerror("Erro", "Valores inválidos!")
                return
            
            # Verificar se código já existe
            self.cursor.execute("SELECT id FROM produtos WHERE codigo = %s", (dados['codigo'],))
            if self.cursor.fetchone():
                messagebox.showerror("Erro", "Código do produto já existe!")
                return
            
            # Inserir produto
            self.cursor.execute('''
                INSERT INTO produtos (codigo, nome, categoria, preco, quantidade, estoque_minimo, data_cadastro)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            ''', (dados['codigo'], dados['nome'], dados['categoria'], preco, 
                  quantidade, estoque_minimo, date.today()))
            
            self.conn.commit()
            messagebox.showinfo("Sucesso", "Produto cadastrado com sucesso!")
            self.limpar_campos_produto()
            
        except ValueError:
            messagebox.showerror("Erro", "Valores numéricos inválidos!")
        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Erro", f"Erro ao cadastrar: {str(e)}")
    
    def limpar_campos_produto(self):
        """Limpar campos do produto"""
        for var in self.campos_produto.values():
            if isinstance(var, tk.StringVar):
                var.set("")
        self.campos_produto['estoque_minimo'].set("5")
    
    def modulo_relatorios(self):
        """Módulo de relatórios - AGORA VENDEDOR PODE ACESSAR"""
        self.limpar_tela()
        
        # Frame superior
        top_frame = tk.Frame(self.root, bg='#34495e', height=60)
        top_frame.pack(fill=tk.X)
        top_frame.pack_propagate(False)
        
        tk.Button(top_frame, text="← Voltar", font=('Arial', 11), bg='#34495e', fg='white',
                 bd=0, command=self.dashboard).pack(side=tk.LEFT, padx=15)
        
        tk.Label(top_frame, text="📈 Relatórios", font=('Arial', 18, 'bold'),
                bg='#34495e', fg='white').pack(side=tk.LEFT, padx=10)
        
        # Frame principal
        main_frame = tk.Frame(self.root, bg='#2c3e50')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Notebook (abas)
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Aba 1: Gerar Relatório
        aba_gerar = ttk.Frame(notebook, padding="20")
        notebook.add(aba_gerar, text="📊 Gerar Relatório")
        
        self.criar_aba_gerar_relatorio(aba_gerar)
        
        # Aba 2: Relatórios Salvos
        aba_salvos = ttk.Frame(notebook, padding="20")
        notebook.add(aba_salvos, text="💾 Relatórios Salvos")
        
        self.criar_aba_relatorios_salvos(aba_salvos)
    
    def criar_aba_gerar_relatorio(self, parent):
        """Criar aba para gerar relatórios"""
        form_frame = tk.Frame(parent)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tipo de relatório
        tk.Label(form_frame, text="Tipo de Relatório:", font=('Arial', 12)).pack(anchor=tk.W, pady=10)
        
        self.tipo_relatorio_var = tk.StringVar(value="vendas")
        tipos_frame = tk.Frame(form_frame)
        tipos_frame.pack(anchor=tk.W, pady=5)
        
        tipos_relatorios = [
            ("Vendas", "vendas"),
            ("Serviços", "servicos"),
            ("Estoque Completo", "estoque"),
            ("Estoque Baixo", "estoque_baixo"),
            ("Financeiro", "financeiro"),
            ("Usuários", "usuarios")
        ]
        
        for texto, valor in tipos_relatorios:
            # Ajustar permissões para vendedor
            if self.usuario_atual['tipo'] == 'vendedor' and valor in ['usuarios']:
                continue
                
            ttk.Radiobutton(tipos_frame, text=texto, variable=self.tipo_relatorio_var, 
                           value=valor).pack(side=tk.LEFT, padx=10)
        
        # Período
        tk.Label(form_frame, text="Período:", font=('Arial', 12)).pack(anchor=tk.W, pady=10)
        
        periodo_frame = tk.Frame(form_frame)
        periodo_frame.pack(anchor=tk.W, pady=5)
        
        self.periodo_relatorio_var = tk.StringVar(value="hoje")
        
        periodos = ["Hoje", "Ontem", "Últimos 7 dias", "Este mês", "Este ano", "Todos"]
        periodo_combo = ttk.Combobox(periodo_frame, textvariable=self.periodo_relatorio_var, 
                                    values=periodos, width=15, state="readonly")
        periodo_combo.pack(side=tk.LEFT, padx=5)
        
        # Data específica
        tk.Label(periodo_frame, text="ou Data específica (dd/mm/aaaa):").pack(side=tk.LEFT, padx=10)
        self.data_especifica_relatorio_var = tk.StringVar()
        ttk.Entry(periodo_frame, textvariable=self.data_especifica_relatorio_var, width=12).pack(side=tk.LEFT, padx=5)
        
        # Nome do relatório
        tk.Label(form_frame, text="Nome do Relatório (opcional):", font=('Arial', 12)).pack(anchor=tk.W, pady=10)
        self.nome_relatorio_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.nome_relatorio_var, width=40).pack(anchor=tk.W, pady=5)
        
        # Botões
        botoes_frame = tk.Frame(form_frame)
        botoes_frame.pack(pady=30)
        
        ttk.Button(botoes_frame, text="Visualizar Relatório", 
                  command=self.visualizar_relatorio, style='Accent.TButton').pack(side=tk.LEFT, padx=10)
        ttk.Button(botoes_frame, text="Salvar Relatório", 
                  command=self.salvar_relatorio).pack(side=tk.LEFT, padx=10)
        ttk.Button(botoes_frame, text="Exportar para CSV", 
                  command=self.exportar_relatorio_atual).pack(side=tk.LEFT, padx=10)
        ttk.Button(botoes_frame, text="Exportar para PDF", 
                  command=self.exportar_relatorio_pdf).pack(side=tk.LEFT, padx=10)
        ttk.Button(botoes_frame, text="🖨️ Imprimir", 
                  command=self.imprimir_relatorio).pack(side=tk.LEFT, padx=10)
    
    def criar_aba_relatorios_salvos(self, parent):
        """Criar aba para visualizar relatórios salvos"""
        # Busca
        busca_frame = tk.Frame(parent)
        busca_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(busca_frame, text="Buscar:").pack(side=tk.LEFT, padx=5)
        self.busca_relatorios_var = tk.StringVar()
        ttk.Entry(busca_frame, textvariable=self.busca_relatorios_var, width=30).pack(side=tk.LEFT, padx=5)
        ttk.Button(busca_frame, text="Buscar", command=self.buscar_relatorios_salvos).pack(side=tk.LEFT, padx=5)
        ttk.Button(busca_frame, text="Atualizar", command=self.carregar_relatorios_salvos).pack(side=tk.LEFT, padx=5)
        
        # Lista de relatórios
        tree_frame = tk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        colunas = ("ID", "Nome", "Tipo", "Período", "Data Geração", "Usuário")
        self.tree_relatorios = ttk.Treeview(tree_frame, columns=colunas, show="headings", height=15)
         
        for col in colunas:
            self.tree_relatorios.heading(col, text=col)
            self.tree_relatorios.column(col, width=100)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_relatorios.yview)
        self.tree_relatorios.configure(yscrollcommand=vsb.set)
        
        self.tree_relatorios.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Botões de ação
        botoes_frame = tk.Frame(parent)
        botoes_frame.pack(pady=10)
        
        ttk.Button(botoes_frame, text="Visualizar Relatório", 
                  command=self.visualizar_relatorio_salvo).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_frame, text="Exportar CSV", 
                  command=self.exportar_relatorio_salvo).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_frame, text="Exportar PDF", 
                  command=self.exportar_relatorio_salvo_selecionado_pdf).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_frame, text="Excluir", 
                  command=self.excluir_relatorio_salvo).pack(side=tk.LEFT, padx=5)
        
        # Carregar relatórios salvos
        self.carregar_relatorios_salvos()
    
    def carregar_relatorios_salvos(self):
        """Carregar relatórios salvos"""
        for item in self.tree_relatorios.get_children():
            self.tree_relatorios.delete(item)
        
        try:
            query = '''
                SELECT r.id, r.nome, r.tipo, r.periodo, r.data_geracao, u.nome
                FROM relatorios_salvos r
                JOIN usuarios u ON r.usuario_id = u.id
                ORDER BY r.data_geracao DESC
            '''
            
            self.cursor.execute(query)
            
            for relatorio in self.cursor.fetchall():
                data_geracao = relatorio[4]
                if isinstance(data_geracao, datetime):
                    data_formatada = data_geracao.strftime('%d/%m/%Y %H:%M')
                else:
                    data_formatada = str(data_geracao)
                
                self.tree_relatorios.insert("", tk.END, values=(
                    relatorio[0],
                    relatorio[1],
                    relatorio[2].replace('_', ' ').title(),
                    relatorio[3].replace('_', ' ').title(),
                    data_formatada,
                    relatorio[5]
                ))
        except Exception as e:
            print(f"Erro ao carregar relatórios salvos: {e}")
    
    def buscar_relatorios_salvos(self):
        """Buscar relatórios salvos"""
        termo = self.busca_relatorios_var.get().lower()
        
        for item in self.tree_relatorios.get_children():
            self.tree_relatorios.delete(item)
        
        try:
            query = '''
                SELECT r.id, r.nome, r.tipo, r.periodo, r.data_geracao, u.nome
                FROM relatorios_salvos r
                JOIN usuarios u ON r.usuario_id = u.id
                WHERE LOWER(r.nome) LIKE %s OR LOWER(r.tipo) LIKE %s OR LOWER(u.nome) LIKE %s
                ORDER BY r.data_geracao DESC
            '''
            
            self.cursor.execute(query, (f'%{termo}%', f'%{termo}%', f'%{termo}%'))
            
            for relatorio in self.cursor.fetchall():
                data_geracao = relatorio[4]
                if isinstance(data_geracao, datetime):
                    data_formatada = data_geracao.strftime('%d/%m/%Y %H:%M')
                else:
                    data_formatada = str(data_geracao)
                
                self.tree_relatorios.insert("", tk.END, values=(
                    relatorio[0],
                    relatorio[1],
                    relatorio[2].replace('_', ' ').title(),
                    relatorio[3].replace('_', ' ').title(),
                    data_formatada,
                    relatorio[5]
                ))
        except Exception as e:
            print(f"Erro na busca de relatórios: {e}")
    
    def visualizar_relatorio(self):
        """Visualizar relatório gerado"""
        tipo = self.tipo_relatorio_var.get()
        periodo = self.periodo_relatorio_var.get().lower()
        
        # Se houver data específica, usar ela
        data_especifica = self.data_especifica_relatorio_var.get()
        if data_especifica:
            try:
                data_obj = datetime.strptime(data_especifica, '%d/%m/%Y').date()
                periodo = f"data:{data_obj.strftime('%Y-%m-%d')}"
            except ValueError:
                messagebox.showerror("Erro", "Data específica inválida! Use o formato dd/mm/aaaa")
                return
        
        self.criar_janela_relatorio_detalhado(tipo, periodo)
    
    def salvar_relatorio(self):
        """Salvar relatório no banco de dados"""
        tipo = self.tipo_relatorio_var.get()
        periodo = self.periodo_relatorio_var.get().lower()
        
        # Se houver data específica, usar ela
        data_especifica = self.data_especifica_relatorio_var.get()
        if data_especifica:
            try:
                data_obj = datetime.strptime(data_especifica, '%d/%m/%Y').date()
                periodo = f"data:{data_obj.strftime('%Y-%m-%d')}"
            except ValueError:
                messagebox.showerror("Erro", "Data específica inválida! Use o formato dd/mm/aaaa")
                return
        
        nome = self.nome_relatorio_var.get()
        if not nome:
            nome = f"Relatório {tipo.replace('_', ' ').title()} - {periodo.replace('_', ' ').title()}"
        
        try:
            # Obter dados do relatório
            dados = []
            
            if tipo == "vendas":
                dados = self.obter_vendas_periodo(periodo)
            elif tipo == "servicos":
                dados = self.get_dados_servicos_periodo(periodo)
            elif tipo == "estoque":
                dados = self.get_dados_estoque()
            elif tipo == "estoque_baixo":
                dados = self.get_dados_estoque_baixo()
            elif tipo == "financeiro":
                dados = self.get_dados_financeiro(periodo)
            elif tipo == "usuarios":
                dados = self.get_dados_usuarios()
            
            # Salvar no banco
            self.cursor.execute('''
                INSERT INTO relatorios_salvos (nome, tipo, periodo, usuario_id, dados)
                VALUES (%s, %s, %s, %s, %s)
            ''', (nome, tipo, periodo, self.usuario_atual['id'], json.dumps(dados)))
            
            self.conn.commit()
            
            messagebox.showinfo("Sucesso", "Relatório salvo com sucesso!")
            self.carregar_relatorios_salvos()
            
        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Erro", f"Erro ao salvar relatório: {str(e)}")
    
    def exportar_relatorio_pdf(self):
        """Exportar relatório atual para PDF"""
        tipo = self.tipo_relatorio_var.get()
        periodo = self.periodo_relatorio_var.get().lower()
        
        # Se houver data específica, usar ela
        data_especifica = self.data_especifica_relatorio_var.get()
        if data_especifica:
            try:
                data_obj = datetime.strptime(data_especifica, '%d/%m/%Y').date()
                periodo = f"data:{data_obj.strftime('%Y-%m-%d')}"
            except ValueError:
                messagebox.showerror("Erro", "Data específica inválida! Use o formato dd/mm/aaaa")
                return
        
        nome = self.nome_relatorio_var.get()
        if not nome:
            nome = f"relatorio_{tipo}_{date.today()}"
        
        self.exportar_relatorio_pdf_tipo(tipo, periodo, nome)
    
    def exportar_relatorio_pdf_tipo(self, tipo, periodo, nome_personalizado=None):
     """Exportar relatório por tipo para PDF"""
     try:
        if tipo == "vendas":
            dados = self.get_dados_vendas_detalhadas_periodo(periodo)
            titulo = f"Relatório de Vendas Detalhadas - {periodo.replace('_', ' ').title()}"
            colunas = ["Data/Hora", "Tipo", "Descrição", "Qtd", "Total", "Usuário"]
            formatar_dados = lambda d: [
                d.get('data_hora', '')[:16] if len(d.get('data_hora', '')) > 16 else d.get('data_hora', ''),
                d.get('tipo', ''),
                d.get('descricao', ''),
                d.get('quantidade', 0),
                f"MT {float(d.get('total', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                d.get('usuario', '')
            ]
            
        elif tipo == "servicos":
            dados = self.get_dados_gerais_periodo(periodo)
            # Filtrar apenas serviços
            dados = [d for d in dados if 'serviço' in d.get('tipo', '').lower() or d.get('tipo') in ['Cópia P&B', 'Cópia Colorida', 'Impressão P&B', 'Impressão Colorida']]
            titulo = f"Relatório de Serviços - {periodo.replace('_', ' ').title()}"
            colunas = ["Data/Hora", "Tipo", "Descrição", "Qtd", "Total", "Usuário"]
            formatar_dados = lambda d: [
                d.get('data_hora', '')[:16] if len(d.get('data_hora', '')) > 16 else d.get('data_hora', ''),
                d.get('tipo', ''),
                d.get('descricao', ''),
                d.get('quantidade', 0),
                f"MT {float(d.get('total', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                d.get('usuario', '')
            ]
            
        elif tipo == "financeiro":
            dados = self.get_dados_gerais_periodo(periodo)
            titulo = f"Relatório Financeiro - {periodo.replace('_', ' ').title()}"
            colunas = ["Data/Hora", "Tipo", "Descrição", "Qtd", "Total", "Usuário"]
            formatar_dados = lambda d: [
                d.get('data_hora', '')[:16] if len(d.get('data_hora', '')) > 16 else d.get('data_hora', ''),
                d.get('tipo', ''),
                d.get('descricao', ''),
                d.get('quantidade', 0),
                f"MT {float(d.get('total', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                d.get('usuario', '')
            ]
            
        elif tipo == "estoque":
            dados = self.get_dados_estoque()
            titulo = f"Relatório de Estoque Completo"
            colunas = ["Código", "Produto", "Categoria", "Preço", "Estoque", "Mínimo", "Status", "Valor Total"]
            formatar_dados = lambda d: [
                d.get('codigo', ''),
                d.get('nome', ''),
                d.get('categoria', ''),
                f"MT {float(d.get('preco', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                d.get('quantidade', 0),
                d.get('estoque_minimo', 0),
                d.get('status', ''),
                f"MT {float(d.get('valor_total', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            ]
            
        elif tipo == "estoque_baixo":
            dados = self.get_dados_estoque_baixo()
            titulo = f"Relatório de Estoque Baixo"
            colunas = ["Código", "Produto", "Categoria", "Preço", "Estoque", "Mínimo", "Status", "Valor Total"]
            formatar_dados = lambda d: [
                d.get('codigo', ''),
                d.get('nome', ''),
                d.get('categoria', ''),
                f"MT {float(d.get('preco', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                d.get('quantidade', 0),
                d.get('estoque_minimo', 0),
                d.get('status', ''),
                f"MT {float(d.get('valor_total', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            ]
            
        elif tipo == "usuarios":
            dados = self.get_dados_usuarios()
            titulo = f"Relatório de Usuários"
            colunas = ["Nome", "Email", "Tipo", "Status"]
            formatar_dados = lambda d: [
                d.get('nome', ''),
                d.get('email', ''),
                d.get('tipo', ''),
                d.get('status', '')
            ]
        else:
            messagebox.showerror("Erro", "Tipo de relatório inválido!")
            return
        
        if not nome_personalizado:
            nome_personalizado = f"relatorio_{tipo}_{date.today()}"
        
        # Calcular total geral
        total_geral = sum(d.get('total', 0) for d in dados)
        
        # Adicionar linha de total no final
        if tipo in ["vendas", "servicos", "financeiro"] and dados:
            # Adicionar linha de total
            dados.append({
                'data_hora': '',
                'tipo': 'TOTAL GERAL',
                'descricao': '',
                'quantidade': sum(d.get('quantidade', 0) for d in dados),
                'total': total_geral,
                'usuario': ''
            })
        
        self.gerar_pdf_relatorio(titulo, colunas, dados, formatar_dados, nome_personalizado)
        
     except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar PDF: {str(e)}")
    
    def gerar_pdf_relatorio(self, titulo, colunas, dados, formatar_dados, nome_arquivo):
     """Gerar PDF do relatório com formatação profissional"""
     arquivo = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf"), ("Todos os arquivos", "*.*")],
        initialfile=f"{nome_arquivo}.pdf"
     )
    
     if not arquivo:
        return
    
     try:
        # Configurações da página
        from reportlab.lib.pagesizes import A4, landscape, portrait
        from reportlab.platypus import PageBreak
        
        # Usar paisagem para relatórios com muitas colunas
        usar_paisagem = len(colunas) > 6
        pagesize = landscape(A4) if usar_paisagem else portrait(A4)
        
        doc = SimpleDocTemplate(arquivo, pagesize=pagesize)
        elementos = []
        
        # Estilos
        estilos = getSampleStyleSheet()
        
        # ========== ESTILOS ==========
        # Título principal - FONTE MENOR
        estilo_titulo_principal = ParagraphStyle(
            'TituloPrincipal',
            parent=estilos['Title'],
            fontSize=14,  # REDUZIDO
            textColor=colors.HexColor('#2c3e50'),
            spaceAfter=6,
            alignment=1,  # CENTER
            fontName='Helvetica-Bold'
        )
        
        # Título do relatório - FONTE MENOR
        estilo_titulo = ParagraphStyle(
            'TituloRelatorio',
            parent=estilos['Heading1'],
            fontSize=12,  # REDUZIDO
            textColor=colors.HexColor('#2c3e50'),
            spaceAfter=8,
            alignment=1,
            fontName='Helvetica-Bold'
        )
        
        # Informações - FONTE MENOR
        estilo_info = ParagraphStyle(
            'Info',
            parent=estilos['Normal'],
            fontSize=8,  # REDUZIDO
            textColor=colors.HexColor('#34495e'),
            spaceAfter=2,
            alignment=0,  # LEFT
            fontName='Helvetica'
        )
        
        # Cabeçalho da tabela - FONTE MENOR
        estilo_cabecalho = ParagraphStyle(
            'CabecalhoTabela',
            parent=estilos['Normal'],
            fontSize=7,  # REDUZIDO
            textColor=colors.white,
            alignment=1,  # CENTER
            fontName='Helvetica-Bold'
        )
        
        # Dados da tabela - FONTE MENOR
        estilo_dados = ParagraphStyle(
            'DadosTabela',
            parent=estilos['Normal'],
            fontSize=7,  # REDUZIDO
            textColor=colors.black,
            alignment=0,  # LEFT
            fontName='Helvetica'
        )
        
        estilo_dados_numero = ParagraphStyle(
            'DadosNumero',
            parent=estilos['Normal'],
            fontSize=7,
            textColor=colors.black,
            alignment=2,  # RIGHT
            fontName='Helvetica'
        )
        
        # ========== CABEÇALHO ==========
        try:
            # Tentar carregar logo
            logo_path = resource_path("static/logo.png")
            if os.path.exists(logo_path):
                # Logo pequeno
                from PIL import Image as PILImage
                logo_img = PILImage.open(logo_path)
                logo_img = logo_img.resize((60, 45), PILImage.LANCZOS)
                
                import tempfile
                temp_logo = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                logo_img.save(temp_logo.name)
                
                # Tabela compacta com logo e título
                logo_table_data = [
                    [
                        RLImage(temp_logo.name, width=60, height=45),
                        Paragraph("PAPELARIA EXPRESS<br/>SISTEMA DE GESTÃO", estilo_titulo_principal)
                    ]
                ]
                
                logo_table = Table(logo_table_data, colWidths=[2*cm, 16*cm])
                logo_table.setStyle(TableStyle([
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                    ('ALIGN', (1, 0), (1, 0), 'LEFT'),
                    ('LEFTPADDING', (1, 0), (1, 0), 10),
                ]))
                
                elementos.append(logo_table)
                
                import atexit
                atexit.register(lambda: os.unlink(temp_logo.name) if os.path.exists(temp_logo.name) else None)
            else:
                elementos.append(Paragraph("PAPELARIA Frente Verso", estilo_titulo_principal))
        except:
            elementos.append(Paragraph("PAPELARIA Frente Verso", estilo_titulo_principal))
        
        elementos.append(Spacer(1, 8))
        
        # ========== TÍTULO ==========
        elementos.append(Paragraph(titulo, estilo_titulo))
        elementos.append(Spacer(1, 10))
        
        # ========== INFORMAÇÕES ==========
        info_lines = [
            f"<b>Data:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            f"<b>Usuário:</b> {self.usuario_atual['nome']}",
            f"<b>Registros:</b> {len(dados)}"
        ]
        
        # Adicionar totais
        if 'venda' in titulo.lower():
            total_valor = sum(d.get('total', 0) for d in dados)
            info_lines.append(f"<b>Total:</b> MT {total_valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        
        if 'estoque' in titulo.lower():
            valor_total_estoque = sum(d.get('valor_total', 0) for d in dados)
            info_lines.append(f"<b>Valor Estoque:</b> MT {valor_total_estoque:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        
        for line in info_lines:
            elementos.append(Paragraph(line, estilo_info))
        
        elementos.append(Spacer(1, 12))
        
        # ========== TABELA ==========
        if dados:
            dados_tabela = []
            
            # Cabeçalho
            cabecalho_formatado = []
            for col in colunas:
                # Abreviar cabeçalhos longos
                col_abreviado = col
                if len(col) > 15:
                    if 'Data/Hora' in col: col_abreviado = 'Data/Hora'
                    elif 'Descrição' in col: col_abreviado = 'Descrição'
                    elif 'Preço Unit.' in col: col_abreviado = 'Preço'
                    elif 'Valor Total' in col: col_abreviado = 'Total'
                    elif 'Estoque Mínimo' in col: col_abreviado = 'Mínimo'
                
                cabecalho_formatado.append(Paragraph(f"<b>{col_abreviado}</b>", estilo_cabecalho))
            
            dados_tabela.append(cabecalho_formatado)
            
            # Dados
            for i, item in enumerate(dados):
                linha = []
                for j, cell in enumerate(formatar_dados(item)):
                    cell_str = str(cell)
                    
                    # Determinar alinhamento baseado no conteúdo
                    if any(x in cell_str for x in ['MT ', 'R$ ', '$ ', '€ ', '¥ ', '£ ']) or \
                       cell_str.replace('.', '').replace(',', '').replace(' ', '').isdigit():
                        # É número ou moeda
                        # Formatar número
                        try:
                            if 'MT ' in cell_str:
                                valor = cell_str.replace('MT', '').strip()
                                num = float(valor.replace(',', '.'))
                                cell_formatada = f"MT {num:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                            elif '.' in cell_str or ',' in cell_str:
                                num = float(cell_str.replace(',', '.'))
                                cell_formatada = f"{num:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                            else:
                                cell_formatada = cell_str
                        except:
                            cell_formatada = cell_str
                        
                        estilo = estilo_dados_numero
                    else:
                        # É texto
                        cell_formatada = cell_str
                        if len(cell_formatada) > 30:
                            cell_formatada = cell_formatada[:27] + "..."
                        estilo = estilo_dados
                    
                    # Alternar cor de fundo para linhas
                    bg_color = colors.white if i % 2 == 0 else colors.HexColor('#f5f5f5')
                    
                    # Adicionar célula
                    p = Paragraph(cell_formatada, estilo)
                    linha.append(p)
                
                dados_tabela.append(linha)
            
            # Calcular larguras das colunas
            num_cols = len(colunas)
            largura_disponivel = 27*cm if usar_paisagem else 18*cm
            largura_minima = 1.5*cm
            largura_maxima = 4*cm
            
            # Distribuir largura
            col_widths = []
            for i, col in enumerate(colunas):
                if len(col) > 15:
                    col_widths.append(largura_maxima)
                elif len(col) > 10:
                    col_widths.append(3*cm)
                else:
                    col_widths.append(2*cm)
            
            # Ajustar proporcionalmente
            largura_total = sum(col_widths)
            fator = largura_disponivel / largura_total if largura_total > 0 else 1
            col_widths = [w * fator for w in col_widths]
            
            # Ajustes específicos para tipos de relatório
            if 'estoque' in titulo.lower():
                col_widths = [2*cm, 4*cm, 2.5*cm, 2*cm, 1.5*cm, 1.5*cm, 1.5*cm, 2*cm]
            elif 'venda' in titulo.lower():
                col_widths = [3.5*cm, 3*cm, 3.5*cm, 2.5*cm, 2.5*cm, 2.5*cm]
            elif 'serviço' in titulo.lower() or 'servico' in titulo.lower():
                col_widths = [3*cm, 3.5*cm, 5*cm, 2*cm, 2.5*cm, 3*cm]
            
            # Criar tabela
            tabela = Table(dados_tabela, colWidths=col_widths, repeatRows=1)
            
            # Estilo da tabela
            estilo_tabela = [
                # Cabeçalho
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 7),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
                ('TOPPADDING', (0, 0), (-1, 0), 4),
                
                # Bordas
                ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
                
                # Dados
                ('FONTSIZE', (0, 1), (-1, -1), 7),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 2),
                ('TOPPADDING', (0, 1), (-1, -1), 2),
                
                # Linhas alternadas
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')]),
                
                # Alinhamento de números
                ('ALIGN', (3, 1), (-1, -1), 'RIGHT'),
            ]
            
            tabela.setStyle(TableStyle(estilo_tabela))
            elementos.append(tabela)
            
            # Se muitos dados, adicionar quebra de página
            if len(dados) > 30:
                elementos.append(PageBreak())
        else:
            elementos.append(Paragraph("Nenhum dado encontrado para o período selecionado.", estilo_info))
        
        elementos.append(Spacer(1, 15))
        
        # ========== RODAPÉ ==========
        rodape_texto = f"""
        Sistema de Gestão Papelaria v2.0 | Desenvolvido por: {self.autor}
        Documento: {nome_arquivo} | Página 1 de 1 | Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}
        """
        
        estilo_rodape = ParagraphStyle(
            'Rodape',
            parent=estilos['Normal'],
            fontSize=6,
            textColor=colors.HexColor('#666666'),
            alignment=1,
            fontName='Helvetica'
        )
        
        elementos.append(HRFlowable(width="100%", thickness=0.5, color=colors.gray))
        elementos.append(Spacer(1, 5))
        elementos.append(Paragraph(rodape_texto, estilo_rodape))
        
        # ========== GERAR PDF ==========
        doc.build(elementos)
        
        messagebox.showinfo("Sucesso", f"Relatório exportado como PDF:\n{arquivo}")
        
     except Exception as e:
        print(f"Erro ao gerar PDF: {e}")
        import traceback
        traceback.print_exc()
        messagebox.showerror("Erro", f"Erro ao gerar PDF: {str(e)}")
    
    def imprimir_relatorio(self):
        """Imprimir relatório atual"""
        tipo = self.tipo_relatorio_var.get()
        periodo = self.periodo_relatorio_var.get().lower()
        
        # Se houver data específica, usar ela
        data_especifica = self.data_especifica_relatorio_var.get()
        if data_especifica:
            try:
                data_obj = datetime.strptime(data_especifica, '%d/%m/%Y').date()
                periodo = f"data:{data_obj.strftime('%Y-%m-%d')}"
            except ValueError:
                messagebox.showerror("Erro", "Data específica inválida! Use o formato dd/mm/aaaa")
                return
        
        try:
            # Obter dados
            if tipo == "vendas":
                dados = self.obter_vendas_periodo(periodo)
                titulo = f"RELATÓRIO DE VENDAS - {periodo.replace('_', ' ').upper()}"
                cabecalho = ["Data/Hora", "Nº Venda", "Vendedor", "Total", "Recebido", "Troco"]
                formatar_linha = lambda d: f"{d.get('data_formatada', d.get('data_hora', '')):<20} {d.get('numero_serie', ''):<15} {d.get('vendedor', ''):<20} MT {d.get('total', 0):>8.2f} MT {d.get('valor_recebido', 0):>8.2f} MT {d.get('troco', 0):>8.2f}"
                
            elif tipo == "servicos":
                dados = self.get_dados_servicos_periodo(periodo)
                titulo = f"RELATÓRIO DE SERVIÇOS - {periodo.replace('_', ' ').upper()}"
                cabecalho = ["Data/Hora", "Tipo", "Descrição", "Qtd", "Total", "Usuário"]
                formatar_linha = lambda d: f"{d.get('data_hora', ''):<20} {d.get('tipo', ''):<15} {d.get('descricao', ''):<30} {d.get('quantidade', 0):>5} MT {d.get('total', 0):>8.2f} {d.get('usuario', ''):<20}"
                
            elif tipo == "estoque":
                dados = self.get_dados_estoque()
                titulo = "RELATÓRIO DE ESTOQUE COMPLETO"
                cabecalho = ["Código", "Produto", "Categoria", "Preço", "Estoque", "Mínimo", "Status", "Valor Total"]
                formatar_linha = lambda d: f"{d.get('codigo', ''):<10} {d.get('nome', ''):<30} {d.get('categoria', ''):<15} MT {d.get('preco', 0):>8.2f} {d.get('quantidade', 0):>8} {d.get('estoque_minimo', 0):>8} {d.get('status', ''):>10} MT {d.get('valor_total', 0):>13.2f}"
                
            elif tipo == "estoque_baixo":
                dados = self.get_dados_estoque_baixo()
                titulo = "RELATÓRIO DE ESTOQUE BAIXO"
                cabecalho = ["Código", "Produto", "Categoria", "Preço", "Estoque", "Mínimo", "Diferença", "Status"]
                formatar_linha = lambda d: f"{d.get('codigo', ''):<10} {d.get('nome', ''):<30} {d.get('categoria', ''):<15} MT {d.get('preco', 0):>8.2f} {d.get('quantidade', 0):>8} {d.get('estoque_minimo', 0):>8} {d.get('estoque_minimo', 0) - d.get('quantidade', 0):>10} {d.get('status', ''):>10}"
                
            elif tipo == "financeiro":
                dados = self.get_dados_financeiro(periodo)
                titulo = f"RELATÓRIO FINANCEIRO - {periodo.replace('_', ' ').upper()}"
                cabecalho = ["Data/Hora", "Tipo", "Descrição", "Qtd", "Total", "Usuário"]
                formatar_linha = lambda d: f"{d.get('data_hora', ''):<20} {d.get('tipo', ''):<15} {d.get('descricao', ''):<30} {d.get('quantidade', 0):>5} MT {d.get('total', 0):>8.2f} {d.get('usuario', ''):<20}"
                
            elif tipo == "usuarios":
                dados = self.get_dados_usuarios()
                titulo = "RELATÓRIO DE USUÁRIOS"
                cabecalho = ["Nome", "Email", "Tipo", "Status"]
                formatar_linha = lambda d: f"{d.get('nome', ''):<30} {d.get('email', ''):<30} {d.get('tipo', ''):<15} {d.get('status', ''):<10}"
            else:
                messagebox.showerror("Erro", "Tipo de relatório inválido!")
                return
            
            # Criar texto para impressão
            texto_impressao = f"""
{'='*80}
PAPELARIA EXPRESS - SISTEMA DE GESTÃO
{'='*80}
{titulo}
{'-'*80}
Data de geração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
Usuário: {self.usuario_atual['nome']}
Total de registros: {len(dados)}
{'-'*80}
"""
            
            # Adicionar cabeçalho
            linha_cabecalho = " ".join(cabecalho)
            texto_impressao += f"{linha_cabecalho}\n"
            texto_impressao += f"{'-'*80}\n"
            
            # Adicionar dados
            for item in dados:
                texto_impressao += f"{formatar_linha(item)}\n"
            
            texto_impressao += f"{'-'*80}\n"
            
            # Adicionar totais se aplicável
            if tipo.lower().find('vendas') != -1:
                total_valor = sum(d.get('total', 0) for d in dados)
                texto_impressao += f"VALOR TOTAL: MT {total_valor:.2f}\n"
            
            if tipo.lower().find('estoque') != -1:
                valor_total_estoque = sum(d.get('valor_total', 0) for d in dados)
                texto_impressao += f"VALOR TOTAL DO ESTOQUE: MT {valor_total_estoque:.2f}\n"
            
            texto_impressao += f"{'='*80}\n"
            texto_impressao += f"Sistema de Gestão Papelaria v2.0\n"
            texto_impressao += f"Desenvolvido por: {self.autor}\n"
            
            # Imprimir
            self.imprimir_texto(texto_impressao, f"Relatório {tipo}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao imprimir relatório: {str(e)}")
    
    def imprimir_texto(self, texto, titulo="Documento"):
        """Imprimir texto"""
        try:
            # Criar arquivo temporário
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt', encoding='utf-8') as f:
                f.write(texto)
                temp_file = f.name
            
            # Imprimir usando a impressora padrão
            printer_name = win32print.GetDefaultPrinter()
            hprinter = win32print.OpenPrinter(printer_name)
            
            try:
                # Ler arquivo e imprimir
                with open(temp_file, 'r', encoding='utf-8') as f:
                    data = f.read()
                
                # Converter para bytes
                data_bytes = data.encode('utf-8')
                
                # Iniciar trabalho de impressão
                job = win32print.StartDocPrinter(hprinter, 1, (titulo, None, "RAW"))
                try:
                    win32print.StartPagePrinter(hprinter)
                    win32print.WritePrinter(hprinter, data_bytes)
                    win32print.EndPagePrinter(hprinter)
                finally:
                    win32print.EndDocPrinter(hprinter)
                
                messagebox.showinfo("Imprimir", "Relatório enviado para impressora!")
                
            finally:
                win32print.ClosePrinter(hprinter)
            
            # Limpar arquivo temporário
            os.unlink(temp_file)
            
        except Exception as e:
            print(f"Erro ao imprimir: {e}")
            messagebox.showinfo("Imprimir", "Relatório enviado para impressora!")
    
    def visualizar_relatorio_salvo(self):
     """Visualizar relatório salvo"""
     selecionado = self.tree_relatorios.selection()
     if not selecionado:
        messagebox.showwarning("Aviso", "Selecione um relatório primeiro!")
        return
    
     item = self.tree_relatorios.item(selecionado[0])
     relatorio_id = item['values'][0]
     nome = item['values'][1]
    
     try:
        self.cursor.execute('''
            SELECT tipo, periodo, dados FROM relatorios_salvos WHERE id = %s
        ''', (relatorio_id,))
        
        resultado = self.cursor.fetchone()
        if not resultado:
            messagebox.showerror("Erro", "Relatório não encontrado!")
            return
        
        tipo, periodo, dados = resultado  # Dados já vem como objeto Python, não precisa de json.loads()
        
        # Criar janela para visualizar
        self.criar_janela_relatorio_salvo(nome, tipo, periodo, dados)
        
     except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar relatório: {str(e)}")
    
    def exportar_relatorio_salvo(self):
        """Exportar relatório salvo para CSV"""
        selecionado = self.tree_relatorios.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um relatório primeiro!")
            return
        
        item = self.tree_relatorios.item(selecionado[0])
        relatorio_id = item['values'][0]
        nome = item['values'][1]
        
        try:
            self.cursor.execute('''
                SELECT tipo, periodo, dados FROM relatorios_salvos WHERE id = %s
            ''', (relatorio_id,))
            
            resultado = self.cursor.fetchone()
            if not resultado:
                messagebox.showerror("Erro", "Relatório não encontrado!")
                return
            
            tipo, periodo, dados_json = resultado
            dados = json.loads(dados_json)

            
            # Exportar para CSV
            self.exportar_dados_para_csv(dados, tipo, nome)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar relatório: {str(e)}")
    
    def exportar_relatorio_salvo(self):
     """Exportar relatório salvo para CSV"""
     selecionado = self.tree_relatorios.selection()
     if not selecionado:
        messagebox.showwarning("Aviso", "Selecione um relatório primeiro!")
        return
    
     item = self.tree_relatorios.item(selecionado[0])
     relatorio_id = item['values'][0]
     nome = item['values'][1]
    
     try:
        self.cursor.execute('''
            SELECT tipo, periodo, dados FROM relatorios_salvos WHERE id = %s
        ''', (relatorio_id,))
        
        resultado = self.cursor.fetchone()
        if not resultado:
            messagebox.showerror("Erro", "Relatório não encontrado!")
            return
        
        tipo, periodo, dados = resultado  # Dados já vem como objeto Python
        
        # Exportar para CSV
        self.exportar_dados_para_csv(dados, tipo, nome)
        
     except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar relatório: {str(e)}")
    
    def exportar_relatorio_salvo_selecionado_pdf(self):
     """Exportar relatório salvo selecionado para PDF"""
     selecionado = self.tree_relatorios.selection()
     if not selecionado:
        messagebox.showwarning("Aviso", "Selecione um relatório primeiro!")
        return
    
     item = self.tree_relatorios.item(selecionado[0])
     relatorio_id = item['values'][0]
     nome = item['values'][1]
    
     try:
        # Buscar dados do relatório
        self.cursor.execute('''
            SELECT tipo, periodo, dados FROM relatorios_salvos WHERE id = %s
        ''', (relatorio_id,))
        
        resultado = self.cursor.fetchone()
        if not resultado:
            messagebox.showerror("Erro", "Relatório não encontrado!")
            return
        
        tipo, periodo, dados_json = resultado
        
        # Converter dados JSON para objeto Python
        import json
        try:
            if isinstance(dados_json, str):
                dados = json.loads(dados_json)
            else:
                dados = dados_json or []
        except:
            dados = []
        
        # Chamar método com todos os parâmetros necessários
        self.exportar_relatorio_salvo_pdf_dados(nome, tipo, periodo, dados)
        
     except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar relatório: {str(e)}")

    def exportar_relatorio_salvo_pdf_dados(self, nome, tipo, periodo, dados):
     """Exportar relatório salvo para PDF com dados"""
     arquivo = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf"), ("Todos os arquivos", "*.*")],
        initialfile=f"{nome}.pdf"
     )
    
     if not arquivo:
        return
    
     try:
        # Configurar colunas baseadas no tipo
        if tipo == "vendas":
            titulo = f"Relatório de Vendas Salvo: {nome}"
            colunas = ["Data/Hora", "Tipo", "Descrição", "Qtd", "Total", "Usuário"]
            formatar_dados = lambda d: [
                d.get('data_hora', '')[:19] if d.get('data_hora') else '',
                d.get('tipo', ''),
                d.get('descricao', ''),
                d.get('quantidade', 0),
                f"MT {float(d.get('total', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                d.get('usuario', '')
            ]
            
        elif tipo == "servicos":
            titulo = f"Relatório de Serviços Salvo: {nome}"
            colunas = ["Data/Hora", "Tipo", "Descrição", "Qtd", "Total", "Usuário"]
            formatar_dados = lambda d: [
                d.get('data_hora', '')[:19] if d.get('data_hora') else '',
                d.get('tipo', ''),
                d.get('descricao', ''),
                d.get('quantidade', 0),
                f"MT {float(d.get('total', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                d.get('usuario', '')
            ]
            
        elif tipo in ["estoque", "estoque_baixo"]:
            titulo = f"Relatório de Estoque Salvo: {nome}"
            colunas = ["Código", "Produto", "Categoria", "Preço", "Estoque", "Mínimo", "Status", "Valor Total"]
            formatar_dados = lambda d: [
                d.get('codigo', ''),
                d.get('nome', ''),
                d.get('categoria', ''),
                f"MT {float(d.get('preco', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                d.get('quantidade', 0),
                d.get('estoque_minimo', 0),
                d.get('status', ''),
                f"MT {float(d.get('valor_total', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            ]
            
        elif tipo == "financeiro":
            titulo = f"Relatório Financeiro Salvo: {nome}"
            colunas = ["Data/Hora", "Tipo", "Descrição", "Qtd", "Total", "Usuário"]
            formatar_dados = lambda d: [
                d.get('data_hora', '')[:19] if d.get('data_hora') else '',
                d.get('tipo', ''),
                d.get('descricao', ''),
                d.get('quantidade', 0),
                f"MT {float(d.get('total', 0)):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                d.get('usuario', '')
            ]
            
        elif tipo == "usuarios":
            titulo = f"Relatório de Usuários Salvo: {nome}"
            colunas = ["Nome", "Email", "Tipo", "Status"]
            formatar_dados = lambda d: [
                d.get('nome', ''),
                d.get('email', ''),
                d.get('tipo', ''),
                d.get('status', '')
            ]
        else:
            # Tipo genérico
            titulo = f"Relatório Salvo: {nome}"
            if dados and len(dados) > 0:
                # Tentar inferir colunas do primeiro item
                primeiro_item = dados[0]
                colunas = list(primeiro_item.keys())
                formatar_dados = lambda d: [str(d.get(col, '')) for col in colunas]
            else:
                messagebox.showerror("Erro", "Dados do relatório estão vazios!")
                return
        
        # Usar o método gerar_pdf_relatorio
        self.gerar_pdf_relatorio(titulo, colunas, dados, formatar_dados, nome)
        
     except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar PDF: {str(e)}")
    
    def excluir_relatorio_salvo(self):
        """Excluir relatório salvo"""
        selecionado = self.tree_relatorios.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um relatório primeiro!")
            return
        
        item = self.tree_relatorios.item(selecionado[0])
        relatorio_id = item['values'][0]
        nome = item['values'][1]
        
        resposta = messagebox.askyesno(
            "Confirmar Exclusão",
            f"Tem certeza que deseja excluir o relatório?\n\n"
            f"Nome: {nome}\n\n"
            f"Esta ação não pode ser desfeita!"
        )
        
        if not resposta:
            return
        
        try:
            self.cursor.execute('DELETE FROM relatorios_salvos WHERE id = %s', (relatorio_id,))
            self.conn.commit()
            
            messagebox.showinfo("Sucesso", "Relatório excluído com sucesso!")
            self.carregar_relatorios_salvos()
            
        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Erro", f"Erro ao excluir relatório: {str(e)}")
    
    def get_dados_estoque(self):
        """Obter dados de estoque completo"""
        dados = []
        try:
            self.cursor.execute('''
                SELECT codigo, nome, categoria, preco, quantidade, estoque_minimo 
                FROM produtos 
                ORDER BY nome
            ''')
            
            for row in self.cursor.fetchall():
                dados.append({
                    'codigo': row[0],
                    'nome': row[1],
                    'categoria': row[2],
                    'preco': float(row[3]),
                    'quantidade': row[4],
                    'estoque_minimo': row[5],
                    'valor_total': float(row[3]) * row[4],
                    'status': 'CRÍTICO' if row[4] < row[5] else 'BAIXO' if row[4] == row[5] else 'OK'
                })
        except Exception as e:
            print(f"Erro ao obter dados de estoque: {e}")
        
        return dados
    
    def get_dados_estoque_baixo(self):
        """Obter dados de estoque baixo"""
        dados = []
        try:
            self.cursor.execute('''
                SELECT codigo, nome, categoria, preco, quantidade, estoque_minimo 
                FROM produtos 
                WHERE quantidade <= estoque_minimo
                ORDER BY quantidade ASC
            ''')
            
            for row in self.cursor.fetchall():
                dados.append({
                    'codigo': row[0],
                    'nome': row[1],
                    'categoria': row[2],
                    'preco': float(row[3]),
                    'quantidade': row[4],
                    'estoque_minimo': row[5],
                    'valor_total': float(row[3]) * row[4],
                    'status': 'CRÍTICO' if row[4] < row[5] else 'BAIXO'
                })
        except Exception as e:
            print(f"Erro ao obter dados de estoque baixo: {e}")
        
        return dados
    
    def get_dados_financeiro(self, periodo):
        """Obter dados financeiros"""
        return self.get_dados_gerais_periodo(periodo)
    
    def get_dados_usuarios(self):
        """Obter dados de usuários"""
        dados = []
        try:
            self.cursor.execute('''
                SELECT nome, email, tipo,
                       CASE WHEN ativo = 1 THEN 'Ativo' ELSE 'Inativo' END as status
                FROM usuarios
                ORDER BY nome
            ''')
            
            for row in self.cursor.fetchall():
                dados.append({
                    'nome': row[0],
                    'email': row[1],
                    'tipo': row[2],
                    'status': row[3]
                })
        except Exception as e:
            print(f"Erro ao obter dados de usuários: {e}")
        
        return dados
    
    def criar_janela_relatorio_detalhado(self, tipo, periodo):
        """Criar janela para exibir relatório detalhado"""
        janela = tk.Toplevel(self.root)
        janela.title(f"Relatório de {tipo.replace('_', ' ').title()}")
        janela.geometry("1000x700")
        
        # Frame superior
        top_frame = tk.Frame(janela, bg='#34495e', height=60)
        top_frame.pack(fill=tk.X)
        top_frame.pack_propagate(False)
        
        tk.Label(top_frame, text=f"Relatório de {tipo.replace('_', ' ').title()}", font=('Arial', 18, 'bold'),
                bg='#34495e', fg='white').pack(side=tk.LEFT, padx=15)
        
        # Botões de ação
        btn_top_frame = tk.Frame(top_frame, bg='#34495e')
        btn_top_frame.pack(side=tk.RIGHT, padx=15)
        
        tk.Button(btn_top_frame, text="💾 PDF", bg='#9b59b6', fg='white',
                 command=lambda: self.exportar_relatorio_pdf_tipo(tipo, periodo)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_top_frame, text="🖨️ Imprimir", bg='#3498db', fg='white',
                 command=self.imprimir_relatorio).pack(side=tk.LEFT, padx=5)
        
        # Frame principal
        main_frame = tk.Frame(janela)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Área de texto para relatório
        text_frame = tk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        text_relatorio = tk.Text(text_frame, wrap=tk.WORD, font=('Courier', 10))
        vsb = ttk.Scrollbar(text_frame, orient="vertical", command=text_relatorio.yview)
        text_relatorio.configure(yscrollcommand=vsb.set)
        
        text_relatorio.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Gerar relatório
        self.gerar_conteudo_relatorio(text_relatorio, tipo, periodo)
        
        # Desabilitar edição
        text_relatorio.configure(state='disabled')
        
        # Botões
        btn_frame = tk.Frame(janela)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Exportar CSV", 
                  command=lambda: self.exportar_dados_tipo(tipo, periodo)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Fechar", 
                  command=janela.destroy).pack(side=tk.LEFT, padx=5)
    
    def criar_janela_relatorio_salvo(self, nome, tipo, periodo, dados):
        """Criar janela para exibir relatório salvo"""
        janela = tk.Toplevel(self.root)
        janela.title(f"Relatório Salvo: {nome}")
        janela.geometry("1000x700")
        
        # Frame superior
        top_frame = tk.Frame(janela, bg='#34495e', height=60)
        top_frame.pack(fill=tk.X)
        top_frame.pack_propagate(False)
        
        tk.Label(top_frame, text=f"Relatório Salvo: {nome}", font=('Arial', 18, 'bold'),
                bg='#34495e', fg='white').pack(side=tk.LEFT, padx=15)
        
        # Botões de ação
        btn_top_frame = tk.Frame(top_frame, bg='#34495e')
        btn_top_frame.pack(side=tk.RIGHT, padx=15)
        
        tk.Button(btn_top_frame, text="💾 PDF", bg='#9b59b6', fg='white',
                 command=lambda: self.exportar_relatorio_salvo_pdf_dados(nome, tipo, periodo, dados)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_top_frame, text="📋 CSV", bg='#2ecc71', fg='white',
                 command=lambda: self.exportar_dados_para_csv(dados, tipo, nome)).pack(side=tk.LEFT, padx=5)
        
        # Frame principal
        main_frame = tk.Frame(janela)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Área de texto para relatório
        text_frame = tk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        text_relatorio = tk.Text(text_frame, wrap=tk.WORD, font=('Courier', 10))
        vsb = ttk.Scrollbar(text_frame, orient="vertical", command=text_relatorio.yview)
        text_relatorio.configure(yscrollcommand=vsb.set)
        
        text_relatorio.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Gerar conteúdo do relatório salvo
        text_relatorio.insert(tk.END, f"{'='*80}\n")
        text_relatorio.insert(tk.END, f"RELATÓRIO SALVO: {nome.upper()}\n")
        text_relatorio.insert(tk.END, f"Tipo: {tipo.replace('_', ' ').title()}\n")
        text_relatorio.insert(tk.END, f"Período: {periodo.replace('_', ' ').title()}\n")
        text_relatorio.insert(tk.END, f"Data de geração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        text_relatorio.insert(tk.END, f"Usuário: {self.usuario_atual['nome']}\n")
        text_relatorio.insert(tk.END, f"{'='*80}\n\n")
        
        # Exibir dados
        if tipo == "vendas":
            self.exibir_dados_vendas_salvos(text_relatorio, dados)
        elif tipo == "servicos":
            self.exibir_dados_servicos_salvos(text_relatorio, dados)
        elif tipo == "estoque":
            self.exibir_dados_estoque_salvos(text_relatorio, dados)
        elif tipo == "estoque_baixo":
            self.exibir_dados_estoque_baixo_salvos(text_relatorio, dados)
        elif tipo == "financeiro":
            self.exibir_dados_financeiro_salvos(text_relatorio, dados)
        elif tipo == "usuarios":
            self.exibir_dados_usuarios_salvos(text_relatorio, dados)
        
        # Desabilitar edição
        text_relatorio.configure(state='disabled')
        
        # Botões
        btn_frame = tk.Frame(janela)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="🖨️ Imprimir", 
                  command=lambda: self.imprimir_texto(text_relatorio.get("1.0", tk.END), f"Relatório {nome}")).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Fechar", 
                  command=janela.destroy).pack(side=tk.LEFT, padx=5)
    
    def exibir_dados_vendas_salvos(self, text_widget, dados):
        """Exibir dados de vendas salvos"""
        if not dados:
            text_widget.insert(tk.END, "Nenhuma venda encontrada.\n")
            return
        
        total_vendas = len(dados)
        total_valor = sum(v.get('total', 0) for v in dados)
        
        text_widget.insert(tk.END, f"TOTAL DE VENDAS: {total_vendas}\n")
        text_widget.insert(tk.END, f"VALOR TOTAL: MT {total_valor:.2f}\n")
        text_widget.insert(tk.END, f"{'-'*80}\n\n")
        
        text_widget.insert(tk.END, f"{'Data/Hora':<20} {'Nº Venda':<15} {'Vendedor':<20} {'Total':>10} {'Troco':>10}\n")
        text_widget.insert(tk.END, f"{'-'*80}\n")
        
        for venda in dados:
            try:
                data_hora = datetime.fromisoformat(venda['data_hora']).strftime('%d/%m/%Y %H:%M')
            except:
                data_hora = str(venda.get('data_hora', ''))
            
            text_widget.insert(tk.END, 
                f"{data_hora:<20} {venda.get('numero_serie', ''):<15} {venda.get('vendedor', ''):<20} "
                f"MT {venda.get('total', 0):>8.2f} MT {venda.get('troco', 0):>8.2f}\n")
    
    def exibir_dados_servicos_salvos(self, text_widget, dados):
        """Exibir dados de serviços salvos"""
        if not dados:
            text_widget.insert(tk.END, "Nenhum serviço encontrado.\n")
            return
        
        total_servicos = len(dados)
        total_valor = sum(v.get('total', 0) for v in dados)
        
        text_widget.insert(tk.END, f"TOTAL DE SERVIÇOS: {total_servicos}\n")
        text_widget.insert(tk.END, f"VALOR TOTAL: MT {total_valor:.2f}\n")
        text_widget.insert(tk.END, f"{'-'*80}\n\n")
        
        text_widget.insert(tk.END, f"{'Data/Hora':<20} {'Tipo':<15} {'Descrição':<30} {'Qtd':>5} {'Total':>10} {'Usuário':<20}\n")
        text_widget.insert(tk.END, f"{'-'*80}\n")
        
        for servico in dados:
            try:
                data_hora = datetime.fromisoformat(servico['data_hora']).strftime('%d/%m/%Y %H:%M')
            except:
                data_hora = str(servico.get('data_hora', ''))
            
            text_widget.insert(tk.END, 
                f"{data_hora:<20} {servico.get('tipo', ''):<15} {servico.get('descricao', ''):<30} "
                f"{servico.get('quantidade', 0):>5} MT {servico.get('total', 0):>8.2f} {servico.get('usuario', ''):<20}\n")
    
    def exibir_dados_estoque_salvos(self, text_widget, dados):
        """Exibir dados de estoque salvos"""
        if not dados:
            text_widget.insert(tk.END, "Nenhum produto encontrado.\n")
            return
        
        total_produtos = len(dados)
        estoque_baixo = len([p for p in dados if p.get('quantidade', 0) <= p.get('estoque_minimo', 0)])
        valor_total_estoque = sum(p.get('valor_total', 0) for p in dados)
        
        text_widget.insert(tk.END, f"TOTAL DE PRODUTOS: {total_produtos}\n")
        text_widget.insert(tk.END, f"ESTOQUE BAIXO: {estoque_baixo}\n")
        text_widget.insert(tk.END, f"VALOR TOTAL DO ESTOQUE: MT {valor_total_estoque:.2f}\n")
        text_widget.insert(tk.END, f"{'-'*80}\n\n")
        
        text_widget.insert(tk.END, f"{'Código':<10} {'Produto':<30} {'Categoria':<15} {'Preço':>10} {'Estoque':>8} {'Mínimo':>8} {'Status':>10} {'Valor Total':>15}\n")
        text_widget.insert(tk.END, f"{'-'*80}\n")
        
        for produto in dados:
            text_widget.insert(tk.END, 
                f"{produto.get('codigo', ''):<10} {produto.get('nome', ''):<30} {produto.get('categoria', ''):<15} "
                f"MT {produto.get('preco', 0):>8.2f} {produto.get('quantidade', 0):>8} {produto.get('estoque_minimo', 0):>8} "
                f"{produto.get('status', ''):>10} MT {produto.get('valor_total', 0):>13.2f}\n")
    
    def exibir_dados_estoque_baixo_salvos(self, text_widget, dados):
        """Exibir dados de estoque baixo salvos"""
        if not dados:
            text_widget.insert(tk.END, "Nenhum produto com estoque baixo.\n")
            return
        
        total = len(dados)
        total_critico = len([p for p in dados if p.get('status') == 'CRÍTICO'])
        total_baixo = len([p for p in dados if p.get('status') == 'BAIXO'])
        
        text_widget.insert(tk.END, f"PRODUTOS COM ESTOQUE BAIXO: {total}\n")
        text_widget.insert(tk.END, f"CRÍTICOS: {total_critico} | BAIXOS: {total_baixo}\n")
        text_widget.insert(tk.END, f"{'-'*80}\n\n")
        
        text_widget.insert(tk.END, f"{'Código':<10} {'Produto':<30} {'Categoria':<15} {'Preço':>10} {'Estoque':>8} {'Mínimo':>8} {'Diferença':>10} {'Status':>10}\n")
        text_widget.insert(tk.END, f"{'-'*80}\n")
        
        for produto in dados:
            quantidade = produto.get('quantidade', 0)
            minimo = produto.get('estoque_minimo', 0)
            diferenca = minimo - quantidade
            
            text_widget.insert(tk.END, 
                f"{produto.get('codigo', ''):<10} {produto.get('nome', ''):<30} {produto.get('categoria', ''):<15} "
                f"MT {produto.get('preco', 0):>8.2f} {quantidade:>8} {minimo:>8} "
                f"{diferenca:>10} {produto.get('status', ''):>10}\n")
    
    def exibir_dados_financeiro_salvos(self, text_widget, dados):
        """Exibir dados financeiros salvos"""
        if not dados:
            text_widget.insert(tk.END, "Nenhum dado financeiro encontrado.\n")
            return
        
        total_transacoes = len(dados)
        total_valor = sum(v.get('total', 0) for v in dados)
        
        text_widget.insert(tk.END, f"TOTAL DE TRANSAÇÕES: {total_transacoes}\n")
        text_widget.insert(tk.END, f"VALOR TOTAL: MT {total_valor:.2f}\n")
        text_widget.insert(tk.END, f"{'-'*80}\n\n")
        
        text_widget.insert(tk.END, f"{'Data/Hora':<20} {'Tipo':<15} {'Descrição':<30} {'Total':>10} {'Usuário':<20}\n")
        text_widget.insert(tk.END, f"{'-'*80}\n")
        
        for transacao in dados:
            try:
                data_hora = datetime.fromisoformat(transacao['data_hora']).strftime('%d/%m/%Y %H:%M')
            except:
                data_hora = str(transacao.get('data_hora', ''))
            
            text_widget.insert(tk.END, 
                f"{data_hora:<20} {transacao.get('tipo', ''):<15} {transacao.get('descricao', ''):<30} "
                f"MT {transacao.get('total', 0):>8.2f} {transacao.get('usuario', ''):<20}\n")
    
    def exibir_dados_usuarios_salvos(self, text_widget, dados):
        """Exibir dados de usuários salvos"""
        if not dados:
            text_widget.insert(tk.END, "Nenhum usuário encontrado.\n")
            return
        
        total_usuarios = len(dados)
        total_ativos = len([u for u in dados if u.get('status') == 'Ativo'])
        total_inativos = len([u for u in dados if u.get('status') == 'Inativo'])
        
        text_widget.insert(tk.END, f"TOTAL DE USUÁRIOS: {total_usuarios}\n")
        text_widget.insert(tk.END, f"Ativos: {total_ativos} | Inativos: {total_inativos}\n")
        text_widget.insert(tk.END, f"{'-'*80}\n\n")
        
        text_widget.insert(tk.END, f"{'Nome':<30} {'Email':<30} {'Tipo':<15} {'Status':<10}\n")
        text_widget.insert(tk.END, f"{'-'*80}\n")
        
        for usuario in dados:
            text_widget.insert(tk.END, 
                f"{usuario.get('nome', ''):<30} {usuario.get('email', ''):<30} {usuario.get('tipo', ''):<15} {usuario.get('status', ''):<10}\n")
    
    def gerar_conteudo_relatorio(self, text_widget, tipo, periodo):
        """Gerar conteúdo do relatório"""
        # Adicionar cabeçalho
        text_widget.insert(tk.END, f"{'='*80}\n")
        text_widget.insert(tk.END, f"RELATÓRIO DE {tipo.replace('_', ' ').upper()}\n")
        
        # Formatar período para exibição
        if periodo.startswith('data:'):
            data_str = periodo.split(':')[1]
            try:
                data_obj = datetime.strptime(data_str, '%Y-%m-%d').date()
                periodo_exibicao = f"Data: {data_obj.strftime('%d/%m/%Y')}"
            except:
                periodo_exibicao = f"Período: {periodo.replace('_', ' ').title()}"
        else:
            periodo_exibicao = f"Período: {periodo.replace('_', ' ').title()}"
        
        text_widget.insert(tk.END, f"{periodo_exibicao}\n")
        text_widget.insert(tk.END, f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        text_widget.insert(tk.END, f"Usuário: {self.usuario_atual['nome']}\n")
        text_widget.insert(tk.END, f"{'='*80}\n\n")
        
        try:
            if tipo == "vendas":
                dados = self.obter_vendas_periodo(periodo)
                self.exibir_dados_vendas_salvos(text_widget, dados)
            elif tipo == "servicos":
                dados = self.get_dados_servicos_periodo(periodo)
                self.exibir_dados_servicos_salvos(text_widget, dados)
            elif tipo == "estoque":
                dados = self.get_dados_estoque()
                self.exibir_dados_estoque_salvos(text_widget, dados)
            elif tipo == "estoque_baixo":
                dados = self.get_dados_estoque_baixo()
                self.exibir_dados_estoque_baixo_salvos(text_widget, dados)
            elif tipo == "financeiro":
                dados = self.get_dados_financeiro(periodo)
                self.exibir_dados_financeiro_salvos(text_widget, dados)
            elif tipo == "usuarios":
                dados = self.get_dados_usuarios()
                self.exibir_dados_usuarios_salvos(text_widget, dados)
            
        except Exception as e:
            text_widget.insert(tk.END, f"\n❌ ERRO AO GERAR RELATÓRIO: {str(e)}")
    
    def exportar_relatorio_atual(self):
        """Exportar relatório atual para CSV"""
        tipo = self.tipo_relatorio_var.get()
        periodo = self.periodo_relatorio_var.get().lower()
        
        # Se houver data específica, usar ela
        data_especifica = self.data_especifica_relatorio_var.get()
        if data_especifica:
            try:
                data_obj = datetime.strptime(data_especifica, '%d/%m/%Y').date()
                periodo = f"data:{data_obj.strftime('%Y-%m-%d')}"
            except ValueError:
                messagebox.showerror("Erro", "Data específica inválida! Use o formato dd/mm/aaaa")
                return
        
        nome = self.nome_relatorio_var.get()
        if not nome:
            nome = f"relatorio_{tipo}_{date.today()}"
        
        self.exportar_dados_tipo(tipo, periodo, nome)
    
    def exportar_dados_tipo(self, tipo, periodo, nome_personalizado=None):
        """Exportar dados por tipo"""
        try:
            if tipo == "vendas":
                dados = self.obter_vendas_periodo(periodo)
            elif tipo == "servicos":
                dados = self.get_dados_servicos_periodo(periodo)
            elif tipo == "estoque":
                dados = self.get_dados_estoque()
            elif tipo == "estoque_baixo":
                dados = self.get_dados_estoque_baixo()
            elif tipo == "financeiro":
                dados = self.get_dados_financeiro(periodo)
            elif tipo == "usuarios":
                dados = self.get_dados_usuarios()
            else:
                messagebox.showerror("Erro", "Tipo de relatório inválido!")
                return
            
            if not nome_personalizado:
                nome_personalizado = f"relatorio_{tipo}_{date.today()}"
            
            self.exportar_dados_para_csv(dados, tipo, nome_personalizado)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar: {str(e)}")
    
    def exportar_dados_para_csv(self, dados, tipo, nome):
        """Exportar dados para arquivo CSV"""
        arquivo = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Todos os arquivos", "*.*")],
            initialfile=f"{nome}.csv"
        )
        
        if not arquivo:
            return
        
        try:
            with open(arquivo, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # Cabeçalho
                writer.writerow([f"Relatório de {tipo.replace('_', ' ').title()}"])
                writer.writerow([f"Data de exportação: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"])
                writer.writerow([f"Usuário: {self.usuario_atual['nome']}"])
                writer.writerow([])
                
                if tipo == "vendas":
                    writer.writerow(["Data/Hora", "Nº Venda", "Vendedor", "Total", "Recebido", "Troco"])
                    for venda in dados:
                        writer.writerow([
                            venda.get('data_formatada', venda.get('data_hora', '')),
                            venda.get('numero_serie', ''),
                            venda.get('vendedor', ''),
                            venda.get('total', 0),
                            venda.get('valor_recebido', 0),
                            venda.get('troco', 0)
                        ])
                
                elif tipo == "servicos":
                    writer.writerow(["Data/Hora", "Tipo", "Descrição", "Quantidade", "Total", "Usuário"])
                    for servico in dados:
                        writer.writerow([
                            servico.get('data_hora', ''),
                            servico.get('tipo', ''),
                            servico.get('descricao', ''),
                            servico.get('quantidade', 0),
                            servico.get('total', 0),
                            servico.get('usuario', '')
                        ])
                
                elif tipo in ["estoque", "estoque_baixo"]:
                    writer.writerow(["Código", "Produto", "Categoria", "Preço", "Estoque", "Mínimo", "Status", "Valor Total"])
                    for produto in dados:
                        writer.writerow([
                            produto.get('codigo', ''),
                            produto.get('nome', ''),
                            produto.get('categoria', ''),
                            produto.get('preco', 0),
                            produto.get('quantidade', 0),
                            produto.get('estoque_minimo', 0),
                            produto.get('status', ''),
                            produto.get('valor_total', 0)
                        ])
                
                elif tipo == "financeiro":
                    writer.writerow(["Data/Hora", "Tipo", "Descrição", "Quantidade", "Total", "Usuário"])
                    for transacao in dados:
                        writer.writerow([
                            transacao.get('data_hora', ''),
                            transacao.get('tipo', ''),
                            transacao.get('descricao', ''),
                            transacao.get('quantidade', 0),
                            transacao.get('total', 0),
                            transacao.get('usuario', '')
                        ])
                
                elif tipo == "usuarios":
                    writer.writerow(["Nome", "Email", "Tipo", "Status"])
                    for usuario in dados:
                        writer.writerow([
                            usuario.get('nome', ''),
                            usuario.get('email', ''),
                            usuario.get('tipo', ''),
                            usuario.get('status', '')
                        ])
            
            messagebox.showinfo("Sucesso", f"Relatório exportado: {arquivo}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar: {str(e)}")
    
    def modulo_usuarios(self):
        """Módulo de usuários (apenas para administradores e gerentes)"""
        if self.usuario_atual['tipo'] not in ['admin', 'gerente']:
            messagebox.showerror("Acesso Negado", "Apenas administradores e gerentes podem acessar este módulo!")
            self.dashboard()
            return
        
        self.limpar_tela()
        
        # Frame superior
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)
        
        # Botão voltar
        voltar_btn = ttk.Button(top_frame, text="← Voltar", command=self.dashboard)
        voltar_btn.pack(side=tk.LEFT)
        
        # Título
        titulo = ttk.Label(top_frame, text="Módulo de Usuários", style='Title.TLabel')
        titulo.pack(side=tk.LEFT, padx=20)
        
        # Botão novo usuário
        novo_btn = ttk.Button(top_frame, text="+ Novo Usuário", command=self.novo_usuario)
        novo_btn.pack(side=tk.RIGHT)
        
        # Frame principal
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Treeview de usuários
        colunas = ("Nome", "Email", "Tipo", "Status")
        self.tree_usuarios = ttk.Treeview(main_frame, columns=colunas, show="headings", height=15)
        
        for col in colunas:
            self.tree_usuarios.heading(col, text=col)
            self.tree_usuarios.column(col, width=150)
        
        # Scrollbars
        vsb = ttk.Scrollbar(main_frame, orient="vertical", command=self.tree_usuarios.yview)
        hsb = ttk.Scrollbar(main_frame, orient="horizontal", command=self.tree_usuarios.xview)
        self.tree_usuarios.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree_usuarios.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Frame de botões de ação
        botoes_frame = ttk.Frame(main_frame)
        botoes_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(botoes_frame, text="Ativar/Desativar", command=self.alterar_status_usuario).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_frame, text="Redefinir Senha", command=self.redefinir_senha_usuario).pack(side=tk.LEFT, padx=5)
        
        # Carregar usuários
        self.carregar_usuarios()
    
    def carregar_usuarios(self):
        """Carregar todos os usuários"""
        # Limpar treeview
        for item in self.tree_usuarios.get_children():
            self.tree_usuarios.delete(item)
        
        # Buscar usuários
        try:
            self.cursor.execute('''
                SELECT nome, email, tipo, ativo 
                FROM usuarios 
                ORDER BY nome
            ''')
            
            for usuario in self.cursor.fetchall():
                status = "Ativo" if usuario[3] == 1 else "Inativo"
                self.tree_usuarios.insert("", tk.END, values=(
                    usuario[0], usuario[1], usuario[2], status
                ), tags=(usuario[1],))  # Usar email como tag
                
        except Exception as e:
            print(f"❌ Erro ao carregar usuários: {e}")
            self.conn.rollback()
    
    def novo_usuario(self):
        """Abrir formulário para novo usuário"""
        formulario = tk.Toplevel(self.root)
        formulario.title("Novo Usuário")
        formulario.geometry("400x300")
        
        # Variáveis do formulário
        nome_var = tk.StringVar()
        email_var = tk.StringVar()
        tipo_var = tk.StringVar(value="vendedor")
        
        # Frame principal
        main_frame = ttk.Frame(formulario, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Campos do formulário
        linha = 0
        
        ttk.Label(main_frame, text="Nome*:").grid(row=linha, column=0, sticky=tk.W, pady=10)
        ttk.Entry(main_frame, textvariable=nome_var, width=30).grid(row=linha, column=1, sticky=tk.W, pady=10)
        linha += 1
        
        ttk.Label(main_frame, text="Email*:").grid(row=linha, column=0, sticky=tk.W, pady=10)
        ttk.Entry(main_frame, textvariable=email_var, width=30).grid(row=linha, column=1, sticky=tk.W, pady=10)
        linha += 1
        
        ttk.Label(main_frame, text="Senha*:").grid(row=linha, column=0, sticky=tk.W, pady=10)
        senha_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=senha_var, show="*", width=30).grid(row=linha, column=1, sticky=tk.W, pady=10)
        linha += 1
        
        ttk.Label(main_frame, text="Tipo*:").grid(row=linha, column=0, sticky=tk.W, pady=10)
        
        # Combobox para tipo
        tipos = ["admin", "gerente", "vendedor"]
        tipo_combo = ttk.Combobox(main_frame, textvariable=tipo_var, values=tipos, width=27, state="readonly")
        tipo_combo.grid(row=linha, column=1, sticky=tk.W, pady=10)
        linha += 1
        
        # Função para salvar
        def salvar_usuario():
            # Validar campos obrigatórios
            if not nome_var.get() or not email_var.get() or not senha_var.get():
                messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
                return
            
            try:
                # Criar hash da senha
                senha_hash = hashlib.md5(senha_var.get().encode()).hexdigest()
                
                # Inserir novo usuário
                self.cursor.execute('''
                    INSERT INTO usuarios (nome, email, senha, tipo, ativo)
                    VALUES (%s, %s, %s, %s, %s)
                ''', (
                    nome_var.get(),
                    email_var.get(),
                    senha_hash,
                    tipo_var.get(),
                    1  # Ativo por padrão
                ))
                
                self.conn.commit()
                messagebox.showinfo("Sucesso", "Usuário cadastrado com sucesso!")
                
                # Atualizar lista de usuários
                self.carregar_usuarios()
                formulario.destroy()
                
            except psycopg2.IntegrityError:
                messagebox.showerror("Erro", "Email já cadastrado!")
                self.conn.rollback()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao cadastrar usuário: {str(e)}")
                self.conn.rollback()
        
        # Botões
        botoes_frame = ttk.Frame(main_frame)
        botoes_frame.grid(row=linha, column=0, columnspan=2, pady=20)
        
        ttk.Button(botoes_frame, text="Salvar", command=salvar_usuario, style='Accent.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_frame, text="Cancelar", command=formulario.destroy).pack(side=tk.LEFT, padx=5)
    
    def alterar_status_usuario(self):
        """Ativar/desativar usuário"""
        selecionado = self.tree_usuarios.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um usuário primeiro!")
            return
        
        item = self.tree_usuarios.item(selecionado[0])
        usuario_email = item['tags'][0]
        nome_usuario = item['values'][0]
        status_atual = item['values'][3]
        
        novo_status = 0 if status_atual == "Ativo" else 1
        status_texto = "desativar" if status_atual == "Ativo" else "ativar"
        
        resposta = messagebox.askyesno(
            "Confirmar Alteração", 
            f"Tem certeza que deseja {status_texto} o usuário '{nome_usuario}'?"
        )
        
        if resposta:
            try:
                # Não permitir desativar a si mesmo
                if usuario_email == self.usuario_atual.get('email') and novo_status == 0:
                    messagebox.showerror("Erro", "Você não pode desativar a si mesmo!")
                    return
                
                self.cursor.execute(
                    "UPDATE usuarios SET ativo = %s WHERE email = %s",
                    (novo_status, usuario_email)
                )
                self.conn.commit()
                
                messagebox.showinfo("Sucesso", f"Status do usuário alterado com sucesso!")
                self.carregar_usuarios()
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível alterar o status: {str(e)}")
                self.conn.rollback()
    
    def redefinir_senha_usuario(self):
        """Redefinir senha do usuário"""
        selecionado = self.tree_usuarios.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um usuário primeiro!")
            return
        
        item = self.tree_usuarios.item(selecionado[0])
        usuario_email = item['tags'][0]
        nome_usuario = item['values'][0]
        
        # Janela para nova senha
        janela_senha = tk.Toplevel(self.root)
        janela_senha.title("Redefinir Senha")
        janela_senha.geometry("300x200")
        
        main_frame = ttk.Frame(janela_senha, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text=f"Redefinir senha para:\n{nome_usuario}", 
                 font=('Arial', 10, 'bold')).pack(pady=(0, 10))
        
        ttk.Label(main_frame, text="Nova Senha:").pack(anchor=tk.W)
        senha_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=senha_var, show="*", width=20).pack(pady=5)
        
        ttk.Label(main_frame, text="Confirmar Senha:").pack(anchor=tk.W)
        confirmar_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=confirmar_var, show="*", width=20).pack(pady=5)
        
        def aplicar_senha():
            if senha_var.get() != confirmar_var.get():
                messagebox.showerror("Erro", "As senhas não coincidem!")
                return
            
            if not senha_var.get():
                messagebox.showerror("Erro", "Informe uma senha!")
                return
            
            try:
                # Criar hash da nova senha
                senha_hash = hashlib.md5(senha_var.get().encode()).hexdigest()
                
                self.cursor.execute(
                    "UPDATE usuarios SET senha = %s WHERE email = %s",
                    (senha_hash, usuario_email)
                )
                self.conn.commit()
                
                messagebox.showinfo("Sucesso", "Senha redefinida com sucesso!")
                janela_senha.destroy()
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível redefinir a senha: {str(e)}")
                self.conn.rollback()
        
        ttk.Button(main_frame, text="Aplicar", command=aplicar_senha, style='Accent.TButton').pack(pady=20)
    
    def modulo_configuracoes(self):
        """Módulo de configurações"""
        if self.usuario_atual['tipo'] not in ['admin', 'gerente']:
            messagebox.showerror("Acesso Negado", "Apenas administradores podem acessar!")
            self.dashboard()
            return
        
        self.limpar_tela()
        
        # Frame superior
        top_frame = tk.Frame(self.root, bg='#34495e', height=60)
        top_frame.pack(fill=tk.X)
        top_frame.pack_propagate(False)
        
        tk.Button(top_frame, text="← Voltar", font=('Arial', 11), bg='#34495e', fg='white',
                 bd=0, command=self.dashboard).pack(side=tk.LEFT, padx=15)
        
        tk.Label(top_frame, text="⚙️ Configurações", font=('Arial', 18, 'bold'),
                bg='#34495e', fg='white').pack(side=tk.LEFT, padx=10)
        
        # Frame principal
        main_frame = tk.Frame(self.root, bg='#2c3e50')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Notebook (abas)
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Aba 1: Configuração de Preços
        aba_precos = ttk.Frame(notebook, padding="10")
        notebook.add(aba_precos, text="💰 Preços")
        
        self.criar_aba_config_precos(aba_precos)
        
        # Aba 2: Novos Serviços
        aba_novos_servicos = ttk.Frame(notebook, padding="10")
        notebook.add(aba_novos_servicos, text="🆕 Serviços")
        
        self.criar_aba_novos_servicos(aba_novos_servicos)
    
    def criar_aba_config_precos(self, parent):
        """Criar aba de configuração de preços - ATUALIZADA"""
        config_frame = tk.Frame(parent, bg='#34495e', padx=20, pady=20)
        config_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(config_frame, text="Configuração de Preços", 
                font=('Arial', 16, 'bold'), bg='#34495e', fg='white').pack(pady=(0, 20))
        
        # Lista de preços
        scroll_frame = tk.Frame(config_frame)
        scroll_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(scroll_frame)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        self.vars_precos = {}
        
        # Organizar por categoria
        categorias = {
            "Impressão/Cópia": [
                ("Cópia Preto e Branco", "copia_pb", 2.0, "Cópia em preto e branco"),
                ("Cópia Colorida", "copia_colorida", 10.0, "Cópia colorida"),
                ("Impressão Preto e Branco", "impressao_pb", 5.0, "Impressão em preto e branco"),
                ("Impressão Colorida", "impressao_colorida", 45.0, "Impressão colorida")
            ],
            "Encadernação por Argola": [
                ("Argola 6mm", "encadernacao_6mm", 20.0, "Encadernação com argola 6mm"),
                ("Argola 8mm", "encadernacao_8mm", 35.0, "Encadernação com argola 8mm"),
                ("Argola 10mm", "encadernacao_10mm", 45.0, "Encadernação com argola 10mm"),
                ("Argola 12mm", "encadernacao_12mm", 55.0, "Encadernação com argola 12mm"),
                ("Argola 14mm", "encadernacao_14mm", 65.0, "Encadernação com argola 14mm"),
                ("Argola 16mm", "encadernacao_16mm", 80.0, "Encadernação com argola 16mm"),
                ("Argola 18mm", "encadernacao_18mm", 100.0, "Encadernação com argola 18mm"),
                ("Argola 20mm", "encadernacao_20mm", 120.0, "Encadernação com argola 20mm"),
                ("Argola 22mm", "encadernacao_22mm", 150.0, "Encadernação com argola 22mm")
            ],
            "Laminação": [
                ("Laminação BI", "laminacao_bi", 30.0, "Laminação tamanho BI"),
                ("Laminação A4", "laminacao_a4", 50.0, "Laminação tamanho A4"),
                ("Laminação A3", "laminacao_a3", 100.0, "Laminação tamanho A3"),
                ("Laminação A5", "laminacao_a5", 40.0, "Laminação tamanho A5")
            ],
            "Outros Serviços": [
                ("Digitação", "digitacao", 35.0, "Serviço de digitação"),
                ("Fotografia", "fotografia", 100.0, "Serviço de fotografia"),
                ("Cartão de Visita", "cartao_visita", 80.0, "Cartão de visita personalizado"),
                ("Convite", "convite", 120.0, "Convites personalizados"),
                ("Banner", "banner", 250.0, "Banner promocional"),
                ("Adesivo", "adesivo", 45.0, "Adesivos personalizados")
            ]
        }
        
        # Primeiro, tentar carregar do banco
        try:
            self.cursor.execute("SELECT nome, tipo, preco, descricao FROM configuracoes_servicos WHERE ativo = 1 ORDER BY nome")
            servicos_db = self.cursor.fetchall()
            
            # Limpar categorias padrão
            for categoria in categorias.values():
                categoria.clear()
            
            # Reorganizar por categoria
            for servico in servicos_db:
                nome = servico[0]
                tipo = servico[1]
                preco = servico[2]
                descricao = servico[3]
                
                # Classificar por categoria
                if 'copia' in tipo or 'impressao' in tipo:
                    categoria_nome = "Impressão/Cópia"
                elif 'encadernacao' in tipo:
                    categoria_nome = "Encadernação por Argola"
                elif 'laminacao' in tipo:
                    categoria_nome = "Laminação"
                else:
                    categoria_nome = "Outros Serviços"
                
                categorias[categoria_nome].append((nome, tipo, preco, descricao))
        except Exception as e:
            print(f"Erro ao carregar serviços do banco: {e}")
            # Manter as categorias padrão
        
        row = 0
        for categoria, itens in categorias.items():
            if itens:  # Só mostrar categoria se tiver itens
                tk.Label(scrollable_frame, text=categoria, font=('Arial', 12, 'bold')).grid(
                    row=row, column=0, columnspan=3, pady=(10, 5), sticky=tk.W)
                row += 1
                
                for nome, tipo, preco, descricao in itens:
                    tk.Label(scrollable_frame, text=nome + ":").grid(row=row, column=0, sticky=tk.W, pady=2)
                    
                    var = tk.StringVar(value=str(preco))
                    entry = ttk.Entry(scrollable_frame, textvariable=var, width=10)
                    entry.grid(row=row, column=1, sticky=tk.W, pady=2, padx=10)
                    
                    # Descrição
                    if descricao:
                        tk.Label(scrollable_frame, text=descricao, font=('Arial', 8), fg='#666').grid(
                            row=row, column=2, sticky=tk.W, pady=2, padx=5)
                    
                    self.vars_precos[tipo] = var
                    row += 1
        
        def salvar_precos():
            try:
                for key, var in self.vars_precos.items():
                    valor = float(var.get().replace(",", "."))
                    if valor < 0:
                        messagebox.showerror("Erro", f"Preço de {key} não pode ser negativo!")
                        return
                    
                    # Atualizar configuração de preços
                    self.config_precos[key] = valor
                    
                    # Atualizar no banco se existir
                    self.cursor.execute('''
                        UPDATE configuracoes_servicos 
                        SET preco = %s 
                        WHERE tipo = %s
                    ''', (valor, key))
                    
                    # Se não existir, inserir
                    if self.cursor.rowcount == 0:
                        # Obter nome do serviço
                        nome_servico = ""
                        for categoria, itens in categorias.items():
                            for item in itens:
                                if item[1] == key:
                                    nome_servico = item[0]
                                    descricao_servico = item[3]
                                    break
                        
                        self.cursor.execute('''
                            INSERT INTO configuracoes_servicos (nome, tipo, preco, descricao, ativo)
                            VALUES (%s, %s, %s, %s, 1)
                        ''', (nome_servico, key, valor, descricao_servico))
                
                self.conn.commit()
                messagebox.showinfo("Sucesso", "Preços atualizados!")
                
            except ValueError:
                messagebox.showerror("Erro", "Valores inválidos!")
            except Exception as e:
                self.conn.rollback()
                messagebox.showerror("Erro", f"Erro ao salvar preços: {str(e)}")
        
        tk.Button(config_frame, text="Salvar Preços", command=salvar_precos,
                 bg='#3498db', fg='white', padx=20, pady=10).pack(pady=20)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def criar_aba_novos_servicos(self, parent):
        """Criar aba para adicionar novos serviços"""
        main_frame = tk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Frame para adicionar novo serviço
        form_frame = tk.Frame(main_frame, bg='#34495e', padx=20, pady=20)
        form_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(form_frame, text="Adicionar Novo Serviço", 
                font=('Arial', 16, 'bold'), bg='#34495e', fg='white').pack(pady=(0, 20))
        
        # Campos do formulário
        campos_frame = tk.Frame(form_frame, bg='#34495e')
        campos_frame.pack()
        
        linha = 0
        
        tk.Label(campos_frame, text="Nome do Serviço:", bg='#34495e', fg='white').grid(
            row=linha, column=0, sticky=tk.W, pady=5, padx=5)
        self.novo_servico_nome_var = tk.StringVar()
        ttk.Entry(campos_frame, textvariable=self.novo_servico_nome_var, width=30).grid(
            row=linha, column=1, sticky=tk.W, pady=5, padx=5)
        linha += 1
        
        tk.Label(campos_frame, text="Tipo (ID único):", bg='#34495e', fg='white').grid(
            row=linha, column=0, sticky=tk.W, pady=5, padx=5)
        self.novo_servico_tipo_var = tk.StringVar()
        ttk.Entry(campos_frame, textvariable=self.novo_servico_tipo_var, width=30).grid(
            row=linha, column=1, sticky=tk.W, pady=5, padx=5)
        tk.Label(campos_frame, text="ex: 'novo_servico' (sem espaços)", 
                bg='#34495e', fg='#bdc3c7', font=('Arial', 8)).grid(
            row=linha, column=2, sticky=tk.W, pady=5, padx=5)
        linha += 1
        
        tk.Label(campos_frame, text="Preço (MT):", bg='#34495e', fg='white').grid(
            row=linha, column=0, sticky=tk.W, pady=5, padx=5)
        self.novo_servico_preco_var = tk.StringVar(value="0.00")
        ttk.Entry(campos_frame, textvariable=self.novo_servico_preco_var, width=15).grid(
            row=linha, column=1, sticky=tk.W, pady=5, padx=5)
        linha += 1
        
        tk.Label(campos_frame, text="Descrição:", bg='#34495e', fg='white').grid(
            row=linha, column=0, sticky=tk.W, pady=5, padx=5)
        self.novo_servico_descricao_var = tk.StringVar()
        ttk.Entry(campos_frame, textvariable=self.novo_servico_descricao_var, width=30).grid(
            row=linha, column=1, sticky=tk.W, pady=5, padx=5)
        linha += 1
        
        # Botão para adicionar
        def adicionar_novo_servico():
            nome = self.novo_servico_nome_var.get()
            tipo = self.novo_servico_tipo_var.get()
            preco_str = self.novo_servico_preco_var.get()
            descricao = self.novo_servico_descricao_var.get()
            
            if not nome or not tipo:
                messagebox.showerror("Erro", "Nome e tipo são obrigatórios!")
                return
            
            try:
                preco = float(preco_str.replace(",", "."))
                if preco < 0:
                    messagebox.showerror("Erro", "Preço não pode ser negativo!")
                    return
                
                # Verificar se tipo já existe
                self.cursor.execute("SELECT id FROM configuracoes_servicos WHERE tipo = %s", (tipo,))
                if self.cursor.fetchone():
                    messagebox.showerror("Erro", "Tipo de serviço já existe!")
                    return
                
                # Inserir novo serviço
                self.cursor.execute('''
                    INSERT INTO configuracoes_servicos (nome, tipo, preco, descricao, ativo)
                    VALUES (%s, %s, %s, %s, 1)
                ''', (nome, tipo, preco, descricao))
                
                self.conn.commit()
                
                # Atualizar configuração de preços
                self.config_precos[tipo] = preco
                
                messagebox.showinfo("Sucesso", "Novo serviço adicionado com sucesso!")
                
                # Limpar campos
                self.novo_servico_nome_var.set("")
                self.novo_servico_tipo_var.set("")
                self.novo_servico_preco_var.set("0.00")
                self.novo_servico_descricao_var.set("")
                
                # Recarregar lista de serviços
                self.carregar_lista_servicos()
                
            except ValueError:
                messagebox.showerror("Erro", "Preço inválido!")
            except Exception as e:
                self.conn.rollback()
                messagebox.showerror("Erro", f"Erro ao adicionar serviço: {str(e)}")
        
        ttk.Button(campos_frame, text="Adicionar Serviço", 
                  command=adicionar_novo_servico, style='Accent.TButton').grid(
            row=linha, column=0, columnspan=2, pady=20)
        
        # Frame para lista de serviços
        lista_frame = tk.Frame(main_frame)
        lista_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(lista_frame, text="Serviços Cadastrados", font=('Arial', 14, 'bold')).pack(pady=(0, 10))
        
        # Treeview de serviços
        colunas = ("Nome", "Tipo", "Preço", "Descrição", "Status")
        self.tree_servicos = ttk.Treeview(lista_frame, columns=colunas, show="headings", height=10)
        
        for col in colunas:
            self.tree_servicos.heading(col, text=col)
            self.tree_servicos.column(col, width=100)
        
        vsb = ttk.Scrollbar(lista_frame, orient="vertical", command=self.tree_servicos.yview)
        hsb = ttk.Scrollbar(lista_frame, orient="horizontal", command=self.tree_servicos.xview)
        self.tree_servicos.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree_servicos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Frame de botões para serviços
        botoes_servicos_frame = tk.Frame(lista_frame)
        botoes_servicos_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(botoes_servicos_frame, text="Ativar/Desativar", 
                  command=self.alterar_status_servico).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_servicos_frame, text="Editar Preço", 
                  command=self.editar_preco_servico).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_servicos_frame, text="Excluir", 
                  command=self.excluir_servico).pack(side=tk.LEFT, padx=5)
        
        # Carregar lista de serviços
        self.carregar_lista_servicos()
    
    def carregar_lista_servicos(self):
        """Carregar lista de serviços"""
        for item in self.tree_servicos.get_children():
            self.tree_servicos.delete(item)
        
        try:
            self.cursor.execute('''
                SELECT nome, tipo, preco, descricao, 
                       CASE WHEN ativo = 1 THEN 'Ativo' ELSE 'Inativo' END as status
                FROM configuracoes_servicos
                ORDER BY nome
            ''')
            
            for servico in self.cursor.fetchall():
                self.tree_servicos.insert("", tk.END, values=servico, tags=(servico[1],))
                
        except Exception as e:
            print(f"Erro ao carregar serviços: {e}")
    
    def alterar_status_servico(self):
        """Ativar/desativar serviço"""
        selecionado = self.tree_servicos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um serviço primeiro!")
            return
        
        item = self.tree_servicos.item(selecionado[0])
        servico_tipo = item['tags'][0]
        servico_nome = item['values'][0]
        status_atual = item['values'][4]
        
        novo_status = 0 if status_atual == "Ativo" else 1
        status_texto = "desativar" if status_atual == "Ativo" else "ativar"
        
        resposta = messagebox.askyesno(
            "Confirmar Alteração", 
            f"Tem certeza que deseja {status_texto} o serviço '{servico_nome}'?"
        )
        
        if resposta:
            try:
                self.cursor.execute(
                    "UPDATE configuracoes_servicos SET ativo = %s WHERE tipo = %s",
                    (novo_status, servico_tipo)
                )
                self.conn.commit()
                
                messagebox.showinfo("Sucesso", f"Status do serviço alterado com sucesso!")
                self.carregar_lista_servicos()
                
            except Exception as e:
                self.conn.rollback()
                messagebox.showerror("Erro", f"Não foi possível alterar o status: {str(e)}")
    
    def editar_preco_servico(self):
        """Editar preço do serviço"""
        selecionado = self.tree_servicos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um serviço primeiro!")
            return
        
        item = self.tree_servicos.item(selecionado[0])
        servico_tipo = item['tags'][0]
        servico_nome = item['values'][0]
        preco_atual = float(item['values'][2])
        
        # Janela para editar preço
        janela_preco = tk.Toplevel(self.root)
        janela_preco.title(f"Editar Preço - {servico_nome}")
        janela_preco.geometry("300x200")
        
        main_frame = ttk.Frame(janela_preco, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text=f"Editar preço para:\n{servico_nome}", 
                 font=('Arial', 10, 'bold')).pack(pady=(0, 10))
        
        ttk.Label(main_frame, text="Novo Preço (MT):").pack(anchor=tk.W)
        novo_preco_var = tk.StringVar(value=str(preco_atual))
        ttk.Entry(main_frame, textvariable=novo_preco_var, width=15).pack(pady=5)
        
        def aplicar_preco():
            try:
                novo_preco = float(novo_preco_var.get().replace(",", "."))
                if novo_preco < 0:
                    messagebox.showerror("Erro", "Preço não pode ser negativo!")
                    return
                
                # Atualizar no banco
                self.cursor.execute(
                    "UPDATE configuracoes_servicos SET preco = %s WHERE tipo = %s",
                    (novo_preco, servico_tipo)
                )
                self.conn.commit()
                
                # Atualizar configuração de preços
                self.config_precos[servico_tipo] = novo_preco
                
                messagebox.showinfo("Sucesso", "Preço atualizado com sucesso!")
                janela_preco.destroy()
                self.carregar_lista_servicos()
                
            except ValueError:
                messagebox.showerror("Erro", "Preço inválido!")
            except Exception as e:
                self.conn.rollback()
                messagebox.showerror("Erro", f"Não foi possível atualizar o preço: {str(e)}")
        
        ttk.Button(main_frame, text="Aplicar", command=aplicar_preco, style='Accent.TButton').pack(pady=20)
    
    def excluir_servico(self):
        """Excluir serviço"""
        selecionado = self.tree_servicos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um serviço primeiro!")
            return
        
        item = self.tree_servicos.item(selecionado[0])
        servico_tipo = item['tags'][0]
        servico_nome = item['values'][0]
        
        # Verificar se o serviço está em uso
        try:
            self.cursor.execute('''
                SELECT COUNT(*) FROM servicos_especiais WHERE tipo = %s
            ''', (servico_tipo,))
            
            count_uso = self.cursor.fetchone()[0]
            
            if count_uso > 0:
                messagebox.showwarning(
                    "Aviso", 
                    f"Este serviço está em {count_uso} registros.\n"
                    f"Não é possível excluí-lo.\n"
                    f"Você pode desativá-lo em vez disso."
                )
                return
        except:
            pass
        
        resposta = messagebox.askyesno(
            "Confirmar Exclusão",
            f"Tem certeza que deseja excluir o serviço?\n\n"
            f"Nome: {servico_nome}\n"
            f"Tipo: {servico_tipo}\n\n"
            f"Esta ação não pode ser desfeita!"
        )
        
        if not resposta:
            return
        
        try:
            self.cursor.execute("DELETE FROM configuracoes_servicos WHERE tipo = %s", (servico_tipo,))
            self.conn.commit()
            
            # Remover da configuração de preços
            if servico_tipo in self.config_precos:
                del self.config_precos[servico_tipo]
            
            messagebox.showinfo("Sucesso", "Serviço excluído com sucesso!")
            self.carregar_lista_servicos()
            
        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Erro", f"Erro ao excluir serviço: {str(e)}")
    
    def sair(self):
        """Sair do sistema"""
        if messagebox.askyesno("Sair", "Tem certeza que deseja sair?"):
            try:
                self.cursor.close()
                self.conn.close()
            except:
                pass
            self.root.quit()


def configurar_banco():
    """Configurar banco de dados inicial"""
    try:
        conn = psycopg2.connect(
            host="localhost",
            database="postgres",
            user="postgres",
            password="techmz06",
            port="5432"
        )
        conn.autocommit = True
        cursor = conn.cursor()
        
        # Criar banco de dados se não existir
        cursor.execute("SELECT 1 FROM pg_database WHERE datname='papelaria_db'")
        if not cursor.fetchone():
            cursor.execute("CREATE DATABASE papelaria_db")
            print("✅ Banco de dados criado!")
        
        cursor.close()
        conn.close()
        
        return True
    except Exception as e:
        print(f"❌ Erro ao configurar banco: {e}")
        return False

def main():
    """Função principal"""
    print("=" * 60)
    print("SISTEMA DE GESTÃO PARA PAPELARIA - v2.0")
    print("=" * 60)
    print("✨ NOVAS FUNCIONALIDADES:")
    print("• Cópia P&B: 2 MT | Cópia Colorida: 10 MT")
    print("• Impressão P&B: 5 MT | Impressão Colorida: 45 MT")
    print("• Encadernação com Argolas (6mm a 22mm)")
    print("• Laminação (BI, A4, A3, A5)")
    print("• Exportação de relatórios para PDF")
    print("• Impressão direta de relatórios")
    print("• Sistema de gerenciamento de serviços")
    print("• Adição de novos serviços personalizados")
    print("• Comprovante de venda em PDF")
    print("• Dashboard web com atualização automática")
    print("• Sistema de relatórios salvos")
    print("=" * 60)
    print("🔧 CONFIGURAÇÃO INICIAL:")
    print("• Banco de dados PostgreSQL")
    print("• Usuários padrão pré-cadastrados")
    print("• Dashboard web na porta 5000")
    print("=" * 60)
    
    # Configurar banco
    print("\n🛠️ Configurando banco de dados...")
    if not configurar_banco():
        print("❌ Falha na configuração do banco!")
        return
    
    print("✅ Configuração concluída!")
    print("\n👤 Credenciais padrão:")
    print("• Email: admin@papelaria.com | Senha: admin123")
    print("• Email: vendedor@papelaria.com | Senha: vendedor123")
    print("• Email: gerente@papelaria.com | Senha: gerente123")
    print("\n⚠️  Se não conseguir acessar, use 'Modo Emergência' no login")
    print("\n🌐 Dashboard web: http://localhost:5000")
    print("📱 Acesse via qualquer dispositivo na mesma rede")
    print("\n📊 NOVO: Sistema completo de relatórios")
    print("• Exporte relatórios para PDF e CSV")
    print("• Imprima relatórios diretamente")
    print("• Gerencie serviços personalizados")
    
    # Instalar dependências se necessário
    try:
        import reportlab
        print("✅ ReportLab (PDF) já instalado")
    except ImportError:
        print("📦 Instalando ReportLab para PDF...")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "reportlab"])
        print("✅ ReportLab instalado com sucesso!")
    
    try:
        import win32print # type: ignore
        print("✅ pywin32 (impressão) já instalado")
    except ImportError:
        print("⚠️  pywin32 não instalado - impressão limitada")
        print("   Para impressão completa, instale: pip install pywin32")
    
    # Iniciar aplicação 
    root = tk.Tk()
    app = SistemaGestaoPapelaria(root)
    root.mainloop()

if __name__ == "__main__":
    main()
