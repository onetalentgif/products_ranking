import PyInstaller.__main__
import os

if __name__ == '__main__':
    print(">>> PyInstaller 빌드를 시작합니다...")
    
    PyInstaller.__main__.run([
        'main.py',
        '--onefile',
        '--clean',
        '--name=products_ranking_bot',  # exe 파일 이름 지정 (옵션)
        # '--noconsole',  # 콘솔 숨기려면 주석 해제
    ])
    
    print(">>> 빌드가 완료되었습니다. 'dist' 폴더를 확인하세요.")
