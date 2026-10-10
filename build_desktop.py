"""Build and package a native candidate on the target operating system."""
from pathlib import Path
import platform
import shutil
import subprocess
import sys

root = Path(__file__).resolve().parent
subprocess.run([sys.executable, str(root/'generate_samples.py')], check=True)
subprocess.run([sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir',
                '--name', 'Reconcile', '--distpath', str(root/'dist'),
                '--workpath', str(root/'build'), '--specpath', str(root/'build'),
                '--add-data', str(root/'interface.html')+':.',
                '--add-data', str(root/'samples')+':samples', str(root/'reconcile.py')], check=True)
bundle = root/'dist'/'Reconcile'
for name in ('交付与使用手册.md', 'SELLING_READINESS.md', 'VALIDATION.md'):
    if (root/name).exists():
        shutil.copy2(root/name, bundle/name)
shutil.copytree(root/'third_party', bundle/'third_party', dirs_exist_ok=True)
if sys.platform == 'darwin':
    launcher = bundle/'启动核对台.command'
    launcher.write_text('#!/bin/zsh\ncd "${0:A:h}"\n./Reconcile\nprintf "\\n程序已退出。按回车关闭窗口。\\n"\nread\n', encoding='utf-8')
    launcher.chmod(0o755)
elif sys.platform == 'win32':
    (bundle/'Start-Reconcile.bat').write_text('@echo off\r\ncd /d "%~dp0"\r\nReconcile.exe\r\npause\r\n', encoding='ascii')
archive = root/'dist'/('Reconcile-1.1.0-rc.1-'+platform.system()+'-'+platform.machine())
shutil.make_archive(str(archive), 'zip', root/'dist', 'Reconcile')
print('Native candidate archive:', str(archive)+'.zip')
