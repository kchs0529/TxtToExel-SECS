# TxtToExel-SECS

SECS (SEMI Equipment Communication Standard) 로그 파일을 Excel 파일로 변환하는 Java 기반 도구입니다.

## 📁 프로젝트 구조

```
TxtToExel-SECS/
├── .settings/                      # Eclipse 설정 파일
├── src/
│   └── main/
│       └── java/
│           └── TxtToExelConverter.java   # 메인 실행 파일
├── target/                         # 빌드 결과물 (.jar 등)
├── .classpath                     # Eclipse 클래스 경로 설정
├── .project                       # Eclipse 프로젝트 설정
├── dependency-reduced-pom.xml    # 쉐이딩된 Maven 의존성 설정
├── pom.xml                        # Maven 기본 설정파일 (Apache POI 등 포함)
```


##  주요 기능

* SECS 로그 텍스트 파일을 엑셀(.xlsx) 파일로 변환
* 변환 대상 SVID 추출 및 정렬 처리
* 사용자 친화적인 GUI 파일 선택 인터페이스 제공

## 사용 방법

1. **저장소 클론**

   ```bash
   git clone https://github.com/kchs0529/TxtToExel-SECS.git
   cd TxtToExel-SECS
   ```

2. **실행 방법**

   * Java 8 이상 필요
   * `TxtToExelConverter.java`를 실행하면 GUI 파일 선택창이 나타납니다.
   * 변환 대상 텍스트 파일을 선택하고, 출력 엑셀 파일명을 입력하면 변환이 시작됩니다.
   * target 폴더 내에 있는 TxtToExelConverter.jar를 TxtToExelConverter.bat 파일과 같은 폴더에 넣은 후 bat 파일을 실행하셔도 됩니다.

3. **출력 파일 확인**

   * 동일 폴더 내에 지정한 이름으로 `.xlsx` 파일이 생성됩니다.

## 주의사항

* 로그 파일은 SVID가 포함된 SECS 로그 형식을 따라야 합니다.
* 입력 형식이 맞지 않으면 변환이 실패할 수 있습니다.
* 첫 번째 List 이후에 나오는 List들은 List 내부에 있는 데이터 값은 추출하지 않고 "List item"이라는 문자열로 출력됩니다.

## 입력 로그 예시

```
[2025-05-20 09:00:00.000 S1F4V11]
L[2][]
    [Item1]
    L[2][]
        [NestedItem1]
        [NestedItem2]
[2025-05-20 09:01:00.000 S1F4V11]
L[3][]
    [100]
    [200]
    [300]
```

## 출력 엑셀 예시 (Excel)

| Unit/Time               | 항목1       | 항목2   | 항목3       | 항목4 |
| ----------------------- | --------- | ----- | --------- | --- |
| 2025-05-20 09:00:00.000 | Item1 | List item| |     |
| 2025-05-20 09:01:00.000 | 100| 200   | 300       |  |

>  **보조 설명**
>
> * 첫 번째 List는 생략됩니다
> * 중첩 List (`indentLevel >= 2`) 내부의 값은 무시됩니다. (2번 이상 들여쓰기 한 값)
> * 두 번째 이후 `L[...][]`는 "List item"으로 표시됩니다.
> * `[ ]` 안의 값은 그대로 Excel에 입력되며, 빈 `[]`는 공백("")으로 채워집니다.
