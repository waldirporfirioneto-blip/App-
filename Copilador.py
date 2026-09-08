import os
import subprocess
import sys
from PIL import Image

NOME_SCRIPT_PRINCIPAL = "Hub_Gree.pyw"  
NOME_PNG = "logo_icone.png"
NOME_ICO = "gree.ico"
# Adicionamos a variável com o nome exato da nova logo
NOME_LOGO_COR = "logo gree colorida.png"

print("=" * 60)
print("     INICIANDO GERADOR DE EXECUTÁVEL - HUB GREE LOGÍSTICA     ")
print("=" * 60)

# PASSO 1: Gerar o arquivo gree.ico
if os.path.exists(NOME_PNG):
    try:
        print(f"\n1. Criando o contêiner de ícone '{NOME_ICO}'...")
        img = Image.open(NOME_PNG)
        img.save(NOME_ICO, format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
        print("✓ Arquivo 'gree.ico' gerado com sucesso!")
    except Exception as e:
        print(f"❌ Erro ao gerar o arquivo .ico: {e}")
        sys.exit(1)
else:
    print(f"❌ Erro: O arquivo '{NOME_PNG}' não foi encontrado!")
    sys.exit(1)

# PASSO 2: Verificar o script principal e a nova logo
if not os.path.exists(NOME_SCRIPT_PRINCIPAL):
    print(f"❌ Erro: O arquivo principal '{NOME_SCRIPT_PRINCIPAL}' não foi encontrado!")
    sys.exit(1)

# Verificação de segurança para a logo colorida
if not os.path.exists(NOME_LOGO_COR):
    print(f"❌ Erro: O arquivo de imagem '{NOME_LOGO_COR}' não foi encontrado na pasta!")
    sys.exit(1)

# PASSO 3: Compilar embutindo a logo interna
try:
    print("\n2. Verificando dependência do PyInstaller no sistema...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    print(f"\n3. Compilando o arquivo '{NOME_SCRIPT_PRINCIPAL}' com recursos embutidos...")
    print("   (Isso pode levar de 1 a 2 minutos. Aguarde...)")
    
    # Adicionamos os parâmetros para injetar o PNG, o ICO e a NOVA LOGO COLORIDA dentro do .exe
    comando = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconsole",
        "--onefile",
        "--add-data", "gree.png;.",         # EMBUTE A LOGO CLARA NO EXECUTÁVEL
        "--add-data", f"{NOME_LOGO_COR};.", # EMBUTE A LOGO COLORIDA NO EXECUTÁVEL
        "--add-data", "gree.ico;.",         # EMBUTE O ÍCONE ICO NO EXECUTÁVEL
        f"--icon={NOME_ICO}",               # ÍCONE DO ARQUIVO .EXE NO WINDOWS
        NOME_SCRIPT_PRINCIPAL
    ]
    
    subprocess.run(comando, check=True)
    
    print("\n" + "=" * 60)
    print("✓ PROCESSAMENTO FINALIZADO COM SUCESSO!")
    print("=" * 60)
    print("O seu aplicativo da GREE está pronto para distribuição!")
    print(f"🚀 Arquivo pronto: Hub_Gree.exe (dentro da pasta 'dist')")
    print("Este executável agora é independente e já contém as logos dentro dele!")
    print("=" * 60)
    
except Exception as e:
    print(f"\n❌ Erro durante a compilação: {e}")