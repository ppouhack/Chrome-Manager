import os
import sys
import shutil
import subprocess
from typing import List

def check_and_install_packages(packages: List[str]):
    """필요한 패키지 확인 및 설치"""
    print("필요한 패키지 확인 및 설치 중...")
    
    for package in packages:
        try:
            __import__(package)
            print(f"✓ {package} 이(가) 설치되어 있습니다")
        except ImportError:
            print(f"{package} 설치 중...")
            subprocess.run([sys.executable, "-m", "pip", "install", package], check=True)
            print(f"✓ {package} 설치 완료")

def install_requirements():
    """필요한 의존성 패키지 설치"""
    required_packages = [
        "pyinstaller",  # 패키징용
        "sv_ttk",      # 테마용
        "keyboard",    # 키보드 감지용
        "mouse",       # 마우스 감지용
        "pywin32",     # Windows API용
        "typing-extensions"  # 타입 힌트용
    ]
    
    
    for package in required_packages:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        except subprocess.CalledProcessError as e:
            print(f"설치 {package} 실패: {str(e)}")
            return False
    return True

def build():
    """프로그램 패키징"""
    # 필요한 패키지 목록
    required_packages = [
        "pyinstaller",
        "sv_ttk",
        "keyboard",
        "mouse",
        "pywin32"
    ]
    
    # 필요한 패키지 확인 및 설치
    check_and_install_packages(required_packages)
    
    # 필요한 모듈 임포트 (설치 후 임포트)
    import sv_ttk
    
    print("\n프로그램 패키징 시작...")
    
    # 기존 빌드 파일 정리
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    
    # sv_ttk 경로 가져오기
    sv_ttk_path = os.path.dirname(sv_ttk.__file__)
    
    # spec 파일 내용 작성
    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['chrome_manager.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app.ico', '.'),
        (r'{sv_ttk_path}', 'sv_ttk')
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [('app.manifest', 'app.manifest', 'DATA')],
    name='chrome_manager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=['app.ico'],
    manifest="app.manifest"
)
'''
    
    # spec 파일 작성
    with open('chrome_manager.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    # app.manifest 파일 작성
    with open('app.manifest', 'w') as f:
        f.write('''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
    <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
        <security>
            <requestedPrivileges>
                <requestedExecutionLevel level="requireAdministrator" uiAccess="false"/>
            </requestedPrivileges>
        </security>
    </trustInfo>
    </assembly>''')
    
    # PyInstaller 실행
    subprocess.run(['pyinstaller', 'chrome_manager.spec'])
    
    print("\n패키징 완료! 프로그램 파일이 dist 폴더에 생성되었습니다.")

if __name__ == "__main__":
    try:
        build()
    except Exception as e:
        print(f"\n오류: {str(e)}")
        input("\n종료하려면 엔터 키를 누르세요...") 