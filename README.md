

# Description

이 툴은 엑셀 파일에 작성한 단어를 한번에 번역하는 도구 입니다.
번역은 Google translator API를 이용했습니다.

# How to run code

py translate.py

# How to build as exe

pyinstaller --onefile --noconsole translate.py

(After build, the file will be placed in /dist folder.)

# How to use this tool

1) 번역할 소스 언어 코드를 입력.  언어코드는 예: en, ko, es, jp ...등
2) 번역할 언어코드 입력.  예: 한국어를 영어로 번역하려면  "소스 언어 코드"에는 'ko'입력,  '번역할 언어코드'는 'en' 입력 
3) 번역할 컬럼명 입력. 예: 번역할 컬럼명이 '단어'면  단어를 입력 
4) 번역할 엑셀 파일 선택 
5) 'Start Translate'버튼 클릭  





