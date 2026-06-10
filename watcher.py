# watcher.py — Monitora a planilha TOTVS e importa automaticamente ao detectar alteração
# Deixe este script rodando em segundo plano. Sempre que a planilha for salva, a importação é disparada.

import sys
import time
import os
from watchdog.observers.polling import PollingObserver
from watchdog.events import FileSystemEventHandler
from datetime import datetime
from importar import importar

# Evita UnicodeEncodeError ao usar emojis nos prints (console Windows usa cp1252)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ─── CONFIGURAÇÃO ────────────────────────────────────────────────────────────
ARQUIVO_MONITORADO = r"P:\QUALIDADE\USUARIOS\00. Dept. Qualidade\07. Controle de Refugo\2026\Planilha de refugo diário Totvs.xlsx"
# ─────────────────────────────────────────────────────────────────────────────

PASTA    = os.path.dirname(ARQUIVO_MONITORADO)
ARQUIVO  = os.path.basename(ARQUIVO_MONITORADO)

class MonitorPlanilha(FileSystemEventHandler):
    def __init__(self):
        self.ultimo_evento = 0  # Evita disparos duplos em sequência rápida

    def _processar(self, caminho):
        # Verifica se é exatamente o arquivo monitorado
        if os.path.basename(caminho) != ARQUIVO:
            return

        agora = time.time()
        # Evita disparos duplos em sequência rápida
        if agora - self.ultimo_evento < 5:
            return
        self.ultimo_evento = agora

        print(f"\n{'='*55}")
        print(f"  📂 Alteração detectada em: {ARQUIVO}")
        print(f"  🕐 {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        print(f"{'='*55}")
        print("  Aguardando Excel finalizar o salvamento...")
        time.sleep(3)

        try:
            importar()
        except Exception as e:
            print(f"\n❌ Erro durante importação: {e}")

    def on_modified(self, event):
        self._processar(event.src_path)

    def on_created(self, event):
        self._processar(event.src_path)

    def on_moved(self, event):
        # O Excel salva criando um arquivo temporário e renomeando-o
        # para o nome original, o que gera um evento "moved".
        self._processar(event.dest_path)

def main():
    print("=" * 55)
    print("  👁️  MONITOR DE PLANILHA — ATIVO")
    print(f"  Arquivo: {ARQUIVO}")
    print(f"  Pasta  : {PASTA}")
    print("  Aguardando alterações... (Ctrl+C para parar)")
    print("=" * 55)

    handler  = MonitorPlanilha()
    observer = PollingObserver(timeout=2)
    observer.schedule(handler, path=PASTA, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n\n⛔ Monitor encerrado.")
        observer.stop()

    observer.join()

if __name__ == "__main__":
    main()
