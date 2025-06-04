import os
import sys
import subprocess
import shutil
import zipfile
from urllib.request import urlretrieve

# Required Python packages
REQUIRED_PACKAGES = [
    'tabula-py',
    'pandas',
    'fpdf2',
    'PyMuPDF',
]

JDK_URL = 'https://download.oracle.com/java/17/latest/jdk-17_windows-x64_bin.zip'
JDK_ZIP = 'java.zip'
JDK_DIR = 'java_runtime'


def is_command_available(cmd):
    try:
        subprocess.run([cmd, '--version'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except FileNotFoundError:
        return False


def install_packages():
    for package in REQUIRED_PACKAGES:
        try:
            __import__(package.split('-')[0])
        except ImportError:
            print(f"Installing missing package: {package}")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])


def install_java():
    if is_command_available('java'):
        return

    print('Java not found. Downloading bundled JDK...')
    urlretrieve(JDK_URL, JDK_ZIP)
    with zipfile.ZipFile(JDK_ZIP, 'r') as zf:
        zf.extractall(JDK_DIR)
    os.remove(JDK_ZIP)

    # JDK directory structure: jdk-17*/bin
    extracted_folders = [f for f in os.listdir(JDK_DIR) if os.path.isdir(os.path.join(JDK_DIR, f))]
    if extracted_folders:
        jdk_bin = os.path.abspath(os.path.join(JDK_DIR, extracted_folders[0], 'bin'))
        os.environ['PATH'] = jdk_bin + os.pathsep + os.environ.get('PATH', '')
        print('Java installed locally.')
    else:
        print('Failed to install Java automatically. Please install Java manually and rerun.')


def run_main():
    exec_path = os.path.join(os.path.dirname(__file__), 'main.py')
    subprocess.run([sys.executable, exec_path])


if __name__ == '__main__':
    install_packages()
    install_java()
    run_main()
