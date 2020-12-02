# add
python --version == 3.7.9

requirements.txt - 설치해야 하는 라이브러리 hwp의 경우 이 방법으로 설치가 안되어 폴더를 직접옮기는 방식으로 해결
  asn1crypto==0.24.0
  cffi==1.12.2
  cryptography==2.6.1
  enum34==1.1.6
  ipaddress==1.0.22
  lxml==4.2.4
  olefile==0.44
  pycparser==2.19
  six==1.12.0

test.zip 사용 방법
가상환경 구축
C:\project>python -m venv example
C:\project>cd example
C:\project\example>Scripts\activate.bat
(example) C:\project\example>
test 폴더의 휠 파일을사용하여 라이브러리 설치
python -m pip install --no-index --find-links="./" -r .\requirements.txt
test2 폴더의 hwp라이브러리를 직접 옮겨줌으로서 설치
