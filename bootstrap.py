import os
import sys
import subprocess
import shutil
import zipfile
from urllib.request import urlretrieve
import runpy

# Directory to persistently store Python packages
PACKAGES_DIR = os.path.join(os.path.dirname(__file__), 'packages')

# Ensure persistent packages directory is on sys.path
if os.path.isdir(PACKAGES_DIR) and PACKAGES_DIR not in sys.path:
    sys.path.append(PACKAGES_DIR)

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
    """Ensure all required packages are installed in PACKAGES_DIR."""
    os.makedirs(PACKAGES_DIR, exist_ok=True)
    if PACKAGES_DIR not in sys.path:
        sys.path.append(PACKAGES_DIR)
    for package in REQUIRED_PACKAGES:
        try:
            __import__(package.split('-')[0])
        except ImportError:
            print(f"Installing missing package: {package}")
            try:
                subprocess.check_call([
                    sys.executable,
                    '-m',
                    'pip',
                    'install',
                    '--upgrade',
                    '--target', PACKAGES_DIR,
                    package,
                ])
            except subprocess.CalledProcessError as e:
                print(f"Failed to install {package}: {e}")
                sys.exit(1)


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
    """Run the main application using the current Python interpreter."""
    exec_path = os.path.join(os.path.dirname(__file__), 'main.py')
    runpy.run_path(exec_path, run_name="__main__")


if __name__ == '__main__':
    install_packages()
    install_java()
    run_main()
