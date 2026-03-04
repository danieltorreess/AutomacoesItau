from datetime import datetime

def log(msg):

    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print(f"[{agora}] {msg}")