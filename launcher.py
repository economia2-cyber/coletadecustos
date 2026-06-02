import tkinter as tk
from tkinter import scrolledtext
import subprocess
import threading
import sys
import os
from datetime import datetime

# Quando rodando como .exe do PyInstaller, sys.executable aponta para o .exe
# Quando rodando como .py normal, __file__ aponta para o script
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_streamlit_proc = None


def run_cmd(cmd):
    r = subprocess.run(cmd, capture_output=True, text=True, cwd=BASE_DIR,
                       encoding='utf-8', errors='replace')
    return r.stdout.strip(), r.stderr.strip(), r.returncode


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Aprosoja/MS — Coleta de Custos")
        self.geometry("440x420")
        self.resizable(False, False)
        self.configure(bg='#0e1a12')
        self.protocol("WM_DELETE_WINDOW", self.fechar)

        # ── Header ────────────────────────────────────────────────────────────
        tk.Label(self, text="  APROSOJA / MS  ", bg='#2D5416', fg='#ffffff',
                 font=('Segoe UI', 14, 'bold'), pady=13).pack(fill='x')
        tk.Label(self, text="Gerenciador — Coleta de Custos · Safra 2026/2027",
                 bg='#0e1a12', fg='#7aa88a', font=('Segoe UI', 8)).pack(pady=(6, 2))

        # ── Status ────────────────────────────────────────────────────────────
        self._status = tk.StringVar(value="Pronto.")
        tk.Label(self, textvariable=self._status, bg='#0e1a12', fg='#c8a415',
                 font=('Segoe UI', 8, 'italic'), wraplength=400).pack(pady=(0, 6))

        # ── Log ───────────────────────────────────────────────────────────────
        self.log = scrolledtext.ScrolledText(
            self, height=7, bg='#0a1a10', fg='#a8d8b0',
            font=('Consolas', 8), state='disabled',
            relief='flat', bd=1, highlightthickness=1,
            highlightbackground='#1c3d28',
        )
        self.log.pack(fill='x', padx=16, pady=(0, 10))

        # ── Botões ────────────────────────────────────────────────────────────
        def btn(texto, cmd, bg='#1B4332', fg='#c8e6a0'):
            b = tk.Button(self, text=texto, command=cmd,
                          bg=bg, fg=fg, activebackground='#2d5a3d',
                          activeforeground='#ffffff', font=('Segoe UI', 10, 'bold'),
                          relief='flat', cursor='hand2', pady=11, bd=0)
            b.pack(fill='x', padx=16, pady=3)
            b.bind('<Enter>', lambda e: b.config(bg='#254d30'))
            b.bind('<Leave>', lambda e: b.config(bg=bg))
            return b

        self.btn_app = btn("🚀  Abrir App de Coleta",
                           self.abrir_app, bg='#b8940f', fg='#0e1a12')
        btn("⬇️  Atualizar do GitHub  (puxar versão mais recente)",
            lambda: threading.Thread(target=self.atualizar, daemon=True).start())
        btn("⬆️  Salvar e Enviar Alterações para o GitHub",
            lambda: threading.Thread(target=self.enviar, daemon=True).start())

        # ── Rodapé ────────────────────────────────────────────────────────────
        tk.Label(self, text="github.com/economia2-cyber/coletadecustos",
                 bg='#0e1a12', fg='#2d5a3d', font=('Segoe UI', 7)).pack(pady=(4, 0))

        self._log("Sistema iniciado. Clique em 🚀 para abrir o app.")

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _log(self, msg):
        self.log.config(state='normal')
        ts = datetime.now().strftime('%H:%M:%S')
        self.log.insert('end', f"[{ts}] {msg}\n")
        self.log.see('end')
        self.log.config(state='disabled')
        self.update_idletasks()

    def _status_set(self, msg):
        self._status.set(msg)
        self.update_idletasks()

    # ── Ações ─────────────────────────────────────────────────────────────────
    def _encontrar_python(self):
        import shutil
        # 1. streamlit direto no PATH
        sl = shutil.which('streamlit')
        if sl:
            return [sl, 'run']
        # 2. python no PATH
        py = shutil.which('python') or shutil.which('python3')
        if py:
            return [py, '-m', 'streamlit', 'run']
        # 3. locais comuns no Windows
        usuario = os.path.expanduser('~')
        candidatos = [
            rf'{usuario}\AppData\Local\Programs\Python\Python312\python.exe',
            rf'{usuario}\AppData\Local\Programs\Python\Python311\python.exe',
            rf'{usuario}\AppData\Local\Programs\Python\Python310\python.exe',
            r'C:\Python312\python.exe',
            r'C:\Python311\python.exe',
        ]
        for p in candidatos:
            if os.path.exists(p):
                return [p, '-m', 'streamlit', 'run']
        return None

    def abrir_app(self):
        global _streamlit_proc
        self._status_set("Localizando Python/Streamlit…")
        self._log("→ Procurando Python no sistema…")

        cmd = self._encontrar_python()

        if cmd is None:
            self._log("❌ Python não encontrado. Abrindo terminal como alternativa…")
            subprocess.Popen(
                ['cmd', '/k', f'cd /d "{BASE_DIR}" && streamlit run app.py --server.port 8501'],
            )
            self._status_set("Terminal aberto — aguarde o navegador.")
            return

        self._log(f"→ Usando: {cmd[0]}")
        self._log("→ Iniciando app.py na porta 8501…")
        app_py = os.path.join(BASE_DIR, 'app.py')
        self._log(f"→ BASE_DIR : {BASE_DIR}")
        self._log(f"→ app.py   : {app_py}")
        self._log(f"→ cmd      : {cmd}")

        # cmd já vem pronto de _encontrar_python:
        #   ['streamlit', 'run']  ou  ['python.exe', '-m', 'streamlit', 'run']
        # Só precisamos acrescentar o app e as opções
        args = cmd + [app_py,
                      '--server.port', '8501',
                      '--server.headless', 'true',
                      '--browser.gatherUsageStats', 'false']

        # Gera bat temporário para manter janela aberta e ver erros
        bat_path = os.path.join(BASE_DIR, '_start.bat')
        linha = ' '.join(f'"{a}"' if ' ' in str(a) else str(a) for a in args)
        with open(bat_path, 'w', encoding='utf-8') as f:
            f.write('@echo off\n')
            f.write(f'cd /d "{BASE_DIR}"\n')
            f.write(f'{linha}\n')
            f.write('pause\n')

        global _streamlit_proc
        _streamlit_proc = subprocess.Popen(
            ['cmd', '/k', bat_path],
            creationflags=subprocess.CREATE_NEW_CONSOLE,
        )
        self._status_set("✅ Terminal aberto — abrindo navegador em 6 segundos…")
        self._log("✅ Aguarde o terminal iniciar o Streamlit…")
        self._log("   (Uma janela preta deve ter aparecido)")
        self.btn_app.config(text="🔄  Reiniciar App de Coleta")
        self.after(6000, self._abrir_navegador)

    def _abrir_navegador(self):
        import webbrowser
        webbrowser.open("http://localhost:8501")
        self._status_set("✅ App rodando em http://localhost:8501")
        self._log("→ Navegador aberto em http://localhost:8501")

    def atualizar(self):
        self._status_set("Atualizando do GitHub…")
        self._log("→ git pull origin main…")
        out, err, code = run_cmd(['git', 'pull', 'origin', 'main'])
        if code == 0:
            self._log(out or "Já na versão mais recente.")
            self._status_set("✅ Atualizado com sucesso!")
        else:
            self._log(f"❌ {err or out}")
            self._status_set("Erro ao atualizar.")

    def enviar(self):
        self._status_set("Verificando alterações…")
        self._log("→ git add .")
        run_cmd(['git', 'add', '.'])

        self._log("→ Verificando se há alterações…")
        out, _, _ = run_cmd(['git', 'status', '--porcelain'])
        if not out:
            self._log("Nenhuma alteração encontrada.")
            self._status_set("Nada para enviar.")
            return

        msg = f"Atualizacao — {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        self._log(f"→ git commit -m \"{msg}\"")
        run_cmd(['git', 'commit', '-m', msg])

        self._log("→ git push origin main…")
        out, err, code = run_cmd(['git', 'push', 'origin', 'main'])
        if code == 0:
            self._log("✅ Enviado para o GitHub!")
            self._log("   Os tablets receberão a versão nova ao recarregar o app.")
            self._status_set("✅ Alterações enviadas com sucesso!")
        else:
            self._log(f"❌ {err or out}")
            self._status_set("Erro ao enviar. Verifique o token.")

    def fechar(self):
        global _streamlit_proc
        if _streamlit_proc and _streamlit_proc.poll() is None:
            _streamlit_proc.terminate()
        self.destroy()


if __name__ == '__main__':
    App().mainloop()
