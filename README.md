# TxtToExel-SECS

SECS (SEMI Equipment Communication Standard) 로그 파일을 Excel 파일로 변환하는 Java 기반 도구입니다.

## 📁 프로젝트 구조

```
TxtToExel-SECS/
├── src/
│   └── TxtToExelConverter.java      # 메인 실행 파일
├── lib/                             # 사용되는 라이브러리 (선택적)
├── README.md                        # 본 문서
├── .gitignore
```

## ⚙️ 주요 기능

* SECS 로그 텍스트 파일을 엑셀(.xlsx) 파일로 변환
* 변환 대상 SVID 추출 및 정렬 처리
* 사용자 친화적인 GUI 파일 선택 인터페이스 제공

## 🚀 사용 방법

1. **저장소 클론**

   ```bash
   git clone https://github.com/kchs0529/TxtToExel-SECS.git
   cd TxtToExel-SECS
   ```

2. **실행 방법**

   * Java 8 이상 필요
   * `TxtToExelConverter.java`를 실행하면 GUI 파일 선택창이 나타납니다.
   * 변환 대상 텍스트 파일을 선택하고, 출력 엑셀 파일명을 입력하면 변환이 시작됩니다.

3. **출력 파일 확인**

   * 동일 폴더 내에 지정한 이름으로 `.xlsx` 파일이 생성됩니다.

## 📝 주의사항

* 로그 파일은 SVID가 포함된 SECS 로그 형식을 따릅니다.
* 입력 형식이 맞지 않으면 변환이 실패할 수 있습니다.
