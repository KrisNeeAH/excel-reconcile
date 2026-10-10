"""Build a native standalone folder with PyInstaller on the target OS."""
from pathlib import Path
import subprocess
import sys

root = Path(__file__).resolve().parent
subprocess.run([sys.executable, str(root/'generate_samples.py')], check=True)
subprocess.run([sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir',
                '--name', 'Reconcile', '--distpath', str(root/'dist'),
                '--workpath', str(root/'build'), '--specpath', str(root/'build'),
                '--add-data', str(root/'interface.html')+':.',
                '--add-data', str(root/'samples')+':samples', str(root/'reconcile.py')], check=True)
