import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import socket
import pandas as pd
import time
import matplotlib.pyplot as plt
import numpy as np
import math

# ================= CONFIG UDP =================
UDP_IP = "192.168.4.1"  # IP do ESP8266 em modo AP
UDP_PORT = 8888
sock = None
is_connected = False

# ================= UTIL =================
def _to_number(v):
    """Converte valores que podem ter vírgula ou ser strings em float."""
    if v is None:
        return None
    if isinstance(v, (int, float, np.number)):
        return float(v)
    s = str(v).strip()
    if s == '':
        return None
    s = s.replace(',', '.')
    try:
        return float(s)
    except:
        return None

# ================= LOG =================
def log(msg):
    txt_log.config(state='normal')
    txt_log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
    txt_log.see(tk.END)
    txt_log.config(state='disabled')

# ================= UDP COMMUNICATION =================
def connect_udp():
    global sock, is_connected
    try:
        if sock:
            sock.close()
        
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.settimeout(3.0)
        sock.bind(('', 0))
        
        # Testa a conexão
        test_msg = "T1"
        sock.sendto(test_msg.encode(), (UDP_IP, UDP_PORT))
        
        try:
            response, addr = sock.recvfrom(1024)
            log(f"✅ Conectado! Resposta: {response.decode()}")
            is_connected = True
            btn_connect.config(text="Conectado", style="Connected.TButton")
        except socket.timeout:
            log("⚠️  Socket criado mas sem resposta do ESP. Verifique a conexão WiFi.")
            is_connected = False
            btn_connect.config(text="Conectar", style="TButton")
            
    except Exception as e:
        log(f"❌ Erro de conexão: {e}")
        is_connected = False
        btn_connect.config(text="Conectar", style="TButton")

def send_udp_command(cmd):
    """Envia comando via UDP e espera resposta"""
    global is_connected
    
    if not sock:
        log("❌ Socket não criado! Clique em Conectar.")
        return None
        
    try:
        sock.sendto(cmd.encode(), (UDP_IP, UDP_PORT))
        log(f"▶️ Enviado: {cmd}")
        
        # Tenta receber resposta
        try:
            response, addr = sock.recvfrom(1024)
            response_text = response.decode().strip()
            log(f"📥 Resposta: {response_text}")
            return response_text
        except socket.timeout:
            log("⏰ Timeout - Sem resposta do ESP")
            return None
        except Exception as e:
            log(f"❌ Erro recebendo resposta: {e}")
            return None
            
    except Exception as e:
        log(f"❌ Erro enviando comando: {e}")
        is_connected = False
        btn_connect.config(text="Conectar", style="TButton")
        return None

# ================= COMANDOS SIMPLES =================
def send_manual():
    cmd = ent_manual.get().strip()
    if cmd:
        send_udp_command(cmd)
        ent_manual.delete(0, tk.END)

def fazer_cruz():
    send_udp_command("C1")
    log("🎯 Comando para fazer cruz enviado")

def servo1_on():
    send_udp_command("S180")
    log("🔧 Servo 1 ativado (180°)")

def servo1_off():
    send_udp_command("S0")
    log("🔧 Servo 1 em repouso (0°)")

def servo2_on():
    send_udp_command("A180")
    log("🔧 Servo 2 ativado (180°)")

def servo2_off():
    send_udp_command("A0")
    log("🔧 Servo 2 em repouso (0°)")

def test_connection():
    send_udp_command("T1")
    log("🔍 Testando conexão...")

# ================= COMANDOS COORDENADOS =================
def enviar_comando_tab():
    """Envia comando no formato tabulado"""
    params = ent_params.get().strip()
    if not params:
        messagebox.showwarning("Aviso", "Digite os parâmetros no formato: Vx Dx Vy Dy Va Da Vz Dz")
        return
    
    # Processa entrada - pode ser com tabs ou espaços
    params_processed = params.replace('\t', ' ').replace(',', ' ')
    
    # Remove espaços extras
    while '  ' in params_processed:
        params_processed = params_processed.replace('  ', ' ')
    
    params_processed = params_processed.strip()
    
    # Verifica se tem 8 valores
    values = params_processed.split()
    if len(values) != 8:
        messagebox.showerror("Erro", f"Esperados 8 valores, encontrados {len(values)}")
        return
    
    # Converte para formato tabulado para envio
    comando = "\t".join(values)
    
    log(f"📤 Enviando comando coordenado: {params_processed}")
    send_udp_command(comando)

# ================= LEITURA DIRETA DA PLANILHA =================
def carregar_planilha():
    global _df
    path = filedialog.askopenfilename(
        filetypes=[("Excel", "*.xlsx"), ("Excel", "*.xls"), ("All files", "*.*")])
    if not path:
        return
    
    try:
        # Lê a planilha - header=None para ter controle total das linhas
        _df = pd.read_excel(path, header=None)
        log(f"📁 Planilha carregada: {path}")
        
        # Mostra informações básicas
        log(f"📊 Planilha carregada com {len(_df)} linhas e {len(_df.columns)} colunas")
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro lendo planilha: {e}")
        log(f"Erro detalhado: {e}")

def enviar_mt():
    """Envia valores da linha Mt (linha 18) - Células C até J"""
    if _df is None:
        messagebox.showerror("Erro", "Carregue a planilha primeiro.")
        return
    
    try:
        # Mt está na linha 18 (índice 17), colunas C a J (índices 2 a 9)
        valores = []
        for col in range(2, 10):  # Colunas C a J
            val = _df.iloc[17, col]  # Linha 18
            num_val = _to_number(val)
            if num_val is not None and not math.isnan(num_val):
                valores.append(num_val)
            else:
                # Se encontrar valor inválido, para aqui
                log(f"❌ Valor inválido na coluna {col}: {val}")
                messagebox.showerror("Erro", f"Valor inválido encontrado na planilha: {val}")
                return
        
        if len(valores) == 8:
            valores_int = [int(round(v)) for v in valores]
            comando = "\t".join(str(v) for v in valores_int)
            
            log("🎯 Enviando curva MT:")
            log(f"   Vx={valores_int[0]}, Dx={valores_int[1]}")
            log(f"   Vy={valores_int[2]}, Dy={valores_int[3]}") 
            log(f"   Va={valores_int[4]}, Da={valores_int[5]}")
            log(f"   Vz={valores_int[6]}, Dz={valores_int[7]}")
            
            send_udp_command(comando)
        else:
            messagebox.showerror("Erro", f"Esperados 8 valores em MT, encontrados: {len(valores)}")
            
    except Exception as e:
        messagebox.showerror("Erro", f"Erro lendo MT: {e}")
        log(f"Erro detalhado: {e}")

def enviar_mr():
    """Envia valores da linha Mr (linha 19) - Células C até J"""
    if _df is None:
        messagebox.showerror("Erro", "Carregue a planilha primeiro.")
        return
    
    try:
        # Mr está na linha 19 (índice 18), colunas C a J (índices 2 a 9)
        valores = []
        for col in range(2, 10):  # Colunas C a J
            val = _df.iloc[18, col]  # Linha 19
            num_val = _to_number(val)
            if num_val is not None and not math.isnan(num_val):
                valores.append(num_val)
            else:
                # Se encontrar valor inválido, para aqui
                log(f"❌ Valor inválido na coluna {col}: {val}")
                messagebox.showerror("Erro", f"Valor inválido encontrado na planilha: {val}")
                return
        
        if len(valores) == 8:
            valores_int = [int(round(v)) for v in valores]
            comando = "\t".join(str(v) for v in valores_int)
            
            log("🎯 Enviando curva MR:")
            log(f"   Vx={valores_int[0]}, Dx={valores_int[1]}")
            log(f"   Vy={valores_int[2]}, Dy={valores_int[3]}") 
            log(f"   Va={valores_int[4]}, Da={valores_int[5]}")
            log(f"   Vz={valores_int[6]}, Dz={valores_int[7]}")
            
            send_udp_command(comando)
        else:
            messagebox.showerror("Erro", f"Esperados 8 valores em MR, encontrados: {len(valores)}")
            
    except Exception as e:
        messagebox.showerror("Erro", f"Erro lendo MR: {e}")
        log(f"Erro detalhado: {e}")

def mostrar_valores_planilha():
    """Mostra os valores atuais da planilha carregada"""
    if _df is None:
        messagebox.showwarning("Aviso", "Nenhuma planilha carregada.")
        return
    
    try:
        # Lê valores de MT - Linha 18
        mt_valores = []
        for col in range(2, 10):
            val = _df.iloc[17, col]  # Linha 18
            num_val = _to_number(val)
            mt_valores.append(num_val if num_val is not None else "NaN")
        
        # Lê valores de MR - Linha 19
        mr_valores = []
        for col in range(2, 10):
            val = _df.iloc[18, col]  # Linha 19
            num_val = _to_number(val)
            mr_valores.append(num_val if num_val is not None else "NaN")
        
        # Mostra os valores
        info_text = "📊 VALORES DA PLANILHA:\n\n"
        info_text += "MT (Linha 18):\n"
        info_text += f"  Vx: {mt_valores[0]}, Dx: {mt_valores[1]}\n"
        info_text += f"  Vy: {mt_valores[2]}, Dy: {mt_valores[3]}\n"
        info_text += f"  Va: {mt_valores[4]}, Da: {mt_valores[5]}\n"
        info_text += f"  Vz: {mt_valores[6]}, Dz: {mt_valores[7]}\n\n"
        
        info_text += "MR (Linha 19):\n"
        info_text += f"  Vx: {mr_valores[0]}, Dx: {mr_valores[1]}\n"
        info_text += f"  Vy: {mr_valores[2]}, Dy: {mr_valores[3]}\n"
        info_text += f"  Va: {mr_valores[4]}, Da: {mr_valores[5]}\n"
        info_text += f"  Vz: {mr_valores[6]}, Dz: {mr_valores[7]}\n"
        
        messagebox.showinfo("Valores da Planilha", info_text)
        log("📋 Valores da planilha exibidos")
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro lendo valores: {e}")

# ================= MOVIMENTOS INDIVIDUAIS =================
def mover_x():
    try:
        passos = int(ent_x.get())
        send_udp_command(f"X{passos}")
        log(f"➡️  Movendo X: {passos} passos")
    except ValueError:
        messagebox.showerror("Erro", "Digite um número válido para X")

def mover_y():
    try:
        passos = int(ent_y.get())
        send_udp_command(f"Y{passos}")
        log(f"⬆️  Movendo Y: {passos} passos")
    except ValueError:
        messagebox.showerror("Erro", "Digite um número válido para Y")

def mover_z():
    try:
        passos = int(ent_z.get())
        send_udp_command(f"Z{passos}")
        log(f"⬇️  Movendo Z: {passos} passos")
    except ValueError:
        messagebox.showerror("Erro", "Digite um número válido para Z")

def mover_e():
    try:
        passos = int(ent_e.get())
        send_udp_command(f"E{passos}")
        log(f"↔️  Movendo E: {passos} passos")
    except ValueError:
        messagebox.showerror("Erro", "Digite um número válido para E")

# ================= GUI =================
def on_exit():
    if sock:
        sock.close()
    root.destroy()

root = tk.Tk()
root.title("Controle Robô - UDP")
root.geometry("800x700")
root.configure(bg="black")

# Configura estilo
style = ttk.Style()
style.theme_use('clam')
style.configure("TLabel", background="black", foreground="white")
style.configure("TButton", background="gray20", foreground="white")
style.configure("TLabelframe", background="black", foreground="white")
style.configure("TLabelframe.Label", background="black", foreground="white")
style.configure("Connected.TButton", background="green", foreground="white")

# Frame de Conexão
frm_conn = ttk.LabelFrame(root, text="Conexão UDP")
frm_conn.pack(padx=10, pady=5, fill='x')

ttk.Label(frm_conn, text=f"IP: {UDP_IP} | Porta: {UDP_PORT}", font=("Arial", 9)).pack(side='left', padx=5)
btn_connect = ttk.Button(frm_conn, text="Conectar", command=connect_udp)
btn_connect.pack(side='right', padx=5)
ttk.Button(frm_conn, text="Testar", command=test_connection).pack(side='right', padx=5)

# Comando manual
frm_manual = ttk.LabelFrame(root, text="Comando Manual (X100, S90, C1, etc)")
frm_manual.pack(padx=10, pady=5, fill='x')
ent_manual = tk.Entry(frm_manual, bg="black", fg="lime", insertbackground="lime", font=("Consolas", 10))
ent_manual.pack(side="left", fill='x', expand=True, padx=5)
ttk.Button(frm_manual, text="Enviar", command=send_manual).pack(side="right", padx=5)
ent_manual.bind('<Return>', lambda e: send_manual())

# Comando coordenado
frm_params = ttk.LabelFrame(root, text="Comando Coordenado (Vx Dx Vy Dy Va Da Vz Dz)")
frm_params.pack(padx=10, pady=5, fill='x')
ent_params = tk.Entry(frm_params, bg="black", fg="cyan", insertbackground="cyan", font=("Consolas", 10))
ent_params.pack(side="left", fill='x', expand=True, padx=5)
ttk.Button(frm_params, text="Enviar", command=enviar_comando_tab).pack(side="right", padx=5)
ent_params.bind('<Return>', lambda e: enviar_comando_tab())

# Frame de Controle de Motores
frm_motores = ttk.LabelFrame(root, text="Controle Individual de Motores")
frm_motores.pack(padx=10, pady=5, fill='x')

# X
frm_x = ttk.Frame(frm_motores)
frm_x.pack(fill='x', pady=2)
ttk.Label(frm_x, text="X:").pack(side='left', padx=5)
ent_x = tk.Entry(frm_x, width=8, bg="black", fg="white", font=("Consolas", 9))
ent_x.pack(side='left', padx=2)
ent_x.insert(0, "1000")
ttk.Button(frm_x, text="Mover X", command=mover_x, width=8).pack(side='left', padx=2)

# Y
frm_y = ttk.Frame(frm_motores)
frm_y.pack(fill='x', pady=2)
ttk.Label(frm_y, text="Y:").pack(side='left', padx=5)
ent_y = tk.Entry(frm_y, width=8, bg="black", fg="white", font=("Consolas", 9))
ent_y.pack(side='left', padx=2)
ent_y.insert(0, "1000")
ttk.Button(frm_y, text="Mover Y", command=mover_y, width=8).pack(side='left', padx=2)

# Z
frm_z = ttk.Frame(frm_motores)
frm_z.pack(fill='x', pady=2)
ttk.Label(frm_z, text="Z:").pack(side='left', padx=5)
ent_z = tk.Entry(frm_z, width=8, bg="black", fg="white", font=("Consolas", 9))
ent_z.pack(side='left', padx=2)
ent_z.insert(0, "1000")
ttk.Button(frm_z, text="Mover Z", command=mover_z, width=8).pack(side='left', padx=2)

# E (A)
frm_e = ttk.Frame(frm_motores)
frm_e.pack(fill='x', pady=2)
ttk.Label(frm_e, text="E(A):").pack(side='left', padx=5)
ent_e = tk.Entry(frm_e, width=8, bg="black", fg="white", font=("Consolas", 9))
ent_e.pack(side='left', padx=2)
ent_e.insert(0, "1000")
ttk.Button(frm_e, text="Mover E", command=mover_e, width=8).pack(side='left', padx=2)

# Frame de Controle
frm_control = ttk.Frame(root)
frm_control.pack(pady=8, fill='x')

# Servos
frm_servos = ttk.LabelFrame(frm_control, text="Controle de Servos")
frm_servos.pack(side='left', padx=5, fill='x', expand=True)
ttk.Button(frm_servos, text="Servo1 ON", command=servo1_on).grid(row=0, column=0, padx=2, pady=2, sticky='ew')
ttk.Button(frm_servos, text="Servo1 OFF", command=servo1_off).grid(row=0, column=1, padx=2, pady=2, sticky='ew')
ttk.Button(frm_servos, text="Servo2 ON", command=servo2_on).grid(row=1, column=0, padx=2, pady=2, sticky='ew')
ttk.Button(frm_servos, text="Servo2 OFF", command=servo2_off).grid(row=1, column=1, padx=2, pady=2, sticky='ew')
frm_servos.grid_columnconfigure(0, weight=1)
frm_servos.grid_columnconfigure(1, weight=1)

# Ações
frm_actions = ttk.LabelFrame(frm_control, text="Ações")
frm_actions.pack(side='left', padx=5, fill='x', expand=True)
ttk.Button(frm_actions, text="🞩 Fazer Cruz", command=fazer_cruz).pack(padx=5, pady=5, fill='x')

# Frame de Planilha
frm_excel = ttk.LabelFrame(root, text="Controle por Planilha")
frm_excel.pack(pady=8, fill='x')

frm_excel_buttons = ttk.Frame(frm_excel)
frm_excel_buttons.pack(fill='x', padx=5, pady=5)

ttk.Button(frm_excel_buttons, text="📂 Carregar Planilha", command=carregar_planilha).pack(side='left', padx=2)
ttk.Button(frm_excel_buttons, text="📊 Ver Valores", command=mostrar_valores_planilha).pack(side='left', padx=2)
ttk.Button(frm_excel_buttons, text="🔄 Enviar MT", command=enviar_mt).pack(side='left', padx=2)
ttk.Button(frm_excel_buttons, text="🔄 Enviar MR", command=enviar_mr).pack(side='left', padx=2)

# Log
frm_log = ttk.LabelFrame(root, text="Log de Comunicação")
frm_log.pack(padx=10, pady=5, fill='both', expand=True)

txt_log = tk.Text(frm_log, height=15, state='disabled', wrap='word', 
                  bg="black", fg="lime", insertbackground="lime", 
                  font=("Consolas", 9))
scrollbar = ttk.Scrollbar(frm_log, orient="vertical", command=txt_log.yview)
txt_log.configure(yscrollcommand=scrollbar.set)

txt_log.pack(side='left', fill='both', expand=True, padx=5, pady=5)
scrollbar.pack(side='right', fill='y', pady=5)

# Dicas
frm_tips = ttk.Frame(root)
frm_tips.pack(padx=10, pady=5, fill='x')
ttk.Label(frm_tips, text="💡 Conecte-se ao WiFi 'ROBO_CNPEM' (senha: 12345678) antes de usar", 
          font=("Arial", 8), foreground="yellow").pack()

root.protocol("WM_DELETE_WINDOW", on_exit)

# Conectar automaticamente após 1 segundo
root.after(1000, connect_udp)

# Focar no campo de comando manual
root.after(1500, lambda: ent_manual.focus())

root.mainloop()