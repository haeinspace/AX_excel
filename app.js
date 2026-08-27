const PAGES = {
  "powerquery-overview": {
    title: "파워쿼리로 할 수 있는 작업",
    subtitle: "파워쿼리",
    image: "powerquery-overview.png",
  },
  filemerge: {
    title: "파일취합자동화",
    subtitle: "파워쿼리",
    sections: [
      {
        heading: "여러 파일 통합",
        steps: [
          "Data 컬럼을 제외한 열 제거",
          "Data 컬럼 &gt;&gt; 결합 버튼 클릭",
          "첫 행을 머리글로 사용",
          "날짜에서 날짜텍스트 필터링",
          "각 컬럼의 타입 체크",
          "홈 &gt;&gt; 쿼리 이름 변경 — 각 파일에 있었던 불필요한 행이 삭제됨",
          "홈 &gt;&gt; 닫기 및 로드",
        ],
      },
      {
        heading: "여러 파일 통합 — Merge",
        steps: [
          "병합할 데이터 파워쿼리로 연결하기 — 앞에서 통합한 자료에 대분류 병합하기",
          "닫기 및 다음으로 로드 선택",
          "엑셀에 데이터를 가져오지 않고 연결만 설정",
          "쿼리 편집기에서 쿼리통합을 클릭합니다",
          "통합파일과 중분류코드를 병합합니다",
          "닫기 및 로드",
        ],
      },
    ],
  },
  "mcode-merge": {
    title: "M코드로 여러파일 합치기",
    subtitle: "파워쿼리",
    sections: [
      {
        heading: "M코드란?",
        callout: {
          icon: "💡",
          paragraphs: [
            "M코드(M Language, Power Query Formula Language)는 파워쿼리가 내부적으로 동작하는 프로그래밍 언어입니다.",
            "우리가 파워쿼리 편집기에서 클릭·드래그로 \"폴더에서 가져오기\", \"열 삭제\", \"확장\" 같은 작업을 할 때마다, 파워쿼리는 그 클릭들을 자동으로 M코드로 번역해서 기록합니다.",
          ],
        },
      },
      {
        heading: "AI 프롬프트 (복사해서 사용하세요)",
        group: [
          {
            heading: "AI 프롬프트 (복사해서 사용하세요)",
            prompt:
              "너는 Microsoft Power Query(M 언어) 전문가야.\n" +
              "아래 조건에 맞는 M코드를 작성해줘.\n" +
              "\n" +
              "[목표]\n" +
              "지정한 폴더 안에 있는 모든 엑셀 파일(.xlsx)을 하나의 표로 자동 통합하는 Power Query M코드를 작성해줘.\n" +
              "\n" +
              "[조건]\n" +
              "- 대상 폴더 경로: 내 파일이 있는 경로 표기\n" +
              "- 폴더 안에는 여러 개의 .xlsx 파일이 있고, 각 파일은 동일한 컬럼 구조(같은 머리글)를 가지고 있어\n" +
              "- 각 파일의 첫 번째 시트만 가져오면 돼\n" +
              "- 숨김 파일이나 .xlsx가 아닌 파일은 자동으로 제외해줘\n" +
              "\n" +
              "[출력 형식]\n" +
              "- 각 단계(step) 이름은 한글로, 무슨 작업을 하는 단계인지 알 수 있게 지어줘",
            promptHighlight: "내 파일이 있는 경로 표기",
          },
          {
            heading:
              '예시: "C:\\Users\\Haein\\Desktop\\교재_업무자동화\\실습자료\\경비지출" 폴더의 파일 합치기',
            callout: {
              icon: "📌",
              steps: [
                '위의 AI 프롬프트에서 "내 파일이 있는 경로 표기" 부분에 내 실제 경로로 바꾸고 ChatGPT, Claude 등에 붙여넣기 합니다',
                "[데이터 탭] &gt;&gt; [데이터 가져오기] &gt;&gt; [기타 원본에서] &gt;&gt; [빈 쿼리] 선택",
                "파워쿼리 편집기 홈 &gt;&gt; [고급 편집기] 선택",
                "AI가 만들어 준 M코드를 고급 편집기에 붙여넣기",
              ],
            },
          },
          {
            heading: "완성된 M코드 예시",
            note: "아래 코드에서 경로 부분은 내 경로에 맞게 수정합니다",
            prompt: String.raw`let
    // 1. 지정한 폴더에 있는 모든 파일 목록 불러오기
    폴더파일불러오기 = Folder.Files("C:\Users\Haein\Desktop\교재_업무자동화\실습자료\경비지출"),

    // 2. 확장자가 .xlsx인 파일만 남기기
    엑셀파일만선택 = Table.SelectRows(폴더파일불러오기, each Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~$")),

    // 3. 숨김 파일 제외하기
    숨김파일제외 = Table.SelectRows(엑셀파일만선택, each [Attributes]?[Hidden]? <> true),

    // 4. 각 엑셀 파일의 내용을 읽기
    엑셀내용읽기 = Table.AddColumn(숨김파일제외, "엑셀내용", each Excel.Workbook([Content], false)),

    // 5. 각 파일에서 첫 번째 시트의 데이터만 가져오기
    첫번째시트가져오기 = Table.AddColumn(엑셀내용읽기, "첫번째시트", each let 시트목록 = Table.SelectRows([엑셀내용], each [Kind] = "Sheet") in if Table.RowCount(시트목록) > 0 then 시트목록{0}[Data] else null),

    // 6. 첫 번째 시트가 없는 파일은 제외하기
    시트없는파일제외 = Table.SelectRows(첫번째시트가져오기, each [첫번째시트] <> null),

    // 7. 각 시트의 첫 번째 행을 머리글로 사용하기
    머리글적용 = Table.TransformColumns(시트없는파일제외, {"첫번째시트", each Table.PromoteHeaders(_, [PromoteAllScalars = true])}),

    // 8. 모든 파일의 첫 번째 시트를 하나의 표로 합치기
    모든파일통합 = Table.Combine(머리글적용[첫번째시트])

in
    모든파일통합`,
            promptHighlight: String.raw`C:\Users\Haein\Desktop\교재_업무자동화\실습자료\경비지출`,
            promptHighlightClass: "code-highlight",
          },
          {
            heading: "실행 결과",
            image: "mcode-result.png",
          },
          {
            heading: "쿼리 이름 변경",
            callout: {
              icon: "📌",
              paragraphs: ["쿼리 이름을 경비지출로 변경합니다"],
            },
          },
        ],
      },
      {
        heading: "전처리",
        callout: {
          icon: "⚠️",
          paragraphs: [
            '수량 컬럼에 "3개", "6개"로 보이는 데이터는 "개"를 빼고 숫자만 보이도록 전처리해야 합니다.',
          ],
        },
      },
      {
        heading: "AI 프롬프트 — 수량 컬럼 숫자만 추출하는 단계 추가",
        group: [
          {
            heading: "AI 프롬프트 — 수량 컬럼 숫자만 추출하는 단계 추가",
            note: "위 M코드를 함께 붙여넣고 사용하세요",
            prompt:
              "너는 Microsoft Power Query(M 언어) 전문가야.\n" +
              "아래는 여러 엑셀 파일을 하나의 표로 통합하는 Power Query M코드야.\n" +
              "\n" +
              "[기존 M코드]\n" +
              "여기에 위 \"완성된 M코드 예시\"를 붙여넣기\n" +
              "\n" +
              "[요청 사항]\n" +
              "이 M코드의 마지막 단계로, \"수량\" 컬럼에서 숫자만 추출하는 단계를 추가해줘.\n" +
              "\n" +
              "[조건]\n" +
              "- \"3개\", \"6개\"처럼 숫자와 문자가 섞여 있는 값은 숫자만 남겨줘\n" +
              "- 이미 숫자만 있는 값은 그대로 유지해줘\n" +
              "- 추출한 값은 정수(숫자) 형식으로 변환해줘\n" +
              "\n" +
              "[출력 형식]\n" +
              "- 기존 단계는 그대로 두고, 마지막에 새 단계만 추가한 전체 M코드를 처음부터 끝까지 다시 작성해줘\n" +
              "- 추가하는 단계 이름은 한글로, 무슨 작업을 하는 단계인지 알 수 있게 지어줘",
            promptHighlight: '여기에 위 "완성된 M코드 예시"를 붙여넣기',
            promptHighlightClass: "prompt-highlight-alt",
          },
          {
            heading: "결과 M코드",
            note: "수량 컬럼 숫자 추출 단계까지 반영된 코드입니다",
            prompt: String.raw`let
    // 1. 지정한 폴더에 있는 모든 파일 목록 불러오기
    폴더파일불러오기 = Folder.Files("C:\Users\Haein\Desktop\교재_업무자동화\실습자료\경비지출"),

    // 2. 확장자가 .xlsx인 파일만 남기기
    엑셀파일만선택 = Table.SelectRows(폴더파일불러오기, each Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~$")),

    // 3. 숨김 파일 제외하기
    숨김파일제외 = Table.SelectRows(엑셀파일만선택, each [Attributes]?[Hidden]? <> true),

    // 4. 각 엑셀 파일의 내용을 읽기
    엑셀내용읽기 = Table.AddColumn(숨김파일제외, "엑셀내용", each Excel.Workbook([Content], false)),

    // 5. 각 파일에서 첫 번째 시트의 데이터만 가져오기
    첫번째시트가져오기 = Table.AddColumn(엑셀내용읽기, "첫번째시트", each let 시트목록 = Table.SelectRows([엑셀내용], each [Kind] = "Sheet") in if Table.RowCount(시트목록) > 0 then 시트목록{0}[Data] else null),

    // 6. 첫 번째 시트가 없는 파일은 제외하기
    시트없는파일제외 = Table.SelectRows(첫번째시트가져오기, each [첫번째시트] <> null),

    // 7. 각 시트의 첫 번째 행을 머리글로 사용하기
    머리글적용 = Table.TransformColumns(시트없는파일제외, {"첫번째시트", each Table.PromoteHeaders(_, [PromoteAllScalars = true])}),

    // 8. 모든 파일의 첫 번째 시트를 하나의 표로 합치기
    모든파일통합 = Table.Combine(머리글적용[첫번째시트]),

    // 9. 수량 컬럼에서 숫자만 추출하고 정수 형식으로 변환하기
    수량컬럼숫자추출 = Table.TransformColumns(모든파일통합, {{"수량", each if _ = null then null else Int64.From(Text.Select(Text.From(_), {"0".."9"})), Int64.Type}})

in
    수량컬럼숫자추출`,
          },
        ],
      },
      {
        heading: "AI 프롬프트 — 날짜/금액 컬럼 타입 변환 단계 추가",
        group: [
          {
            heading: "AI 프롬프트 — 날짜/금액 컬럼 타입 변환 단계 추가",
            note: "위 결과 M코드를 함께 붙여넣고 사용하세요",
            prompt:
              "너는 Microsoft Power Query(M 언어) 전문가야.\n" +
              "아래는 여러 엑셀 파일을 하나의 표로 통합하는 Power Query M코드야.\n" +
              "\n" +
              "[기존 M코드]\n" +
              "여기에 위 \"결과 M코드\"를 붙여넣기\n" +
              "\n" +
              "[요청 사항]\n" +
              "이 M코드의 마지막 단계로, 다음 두 컬럼의 타입을 변환하는 단계를 추가해줘.\n" +
              "\n" +
              "[조건]\n" +
              "- \"날짜\" 컬럼은 날짜(Date) 타입으로 변환해줘\n" +
              "- \"금액\" 컬럼은 숫자(Number) 타입으로 변환해줘\n" +
              "- 변환 실패한 값은 null로 처리하고 실패 건수를 세어줘. 단, 실패 건수가 0건이면 오류 개수 컬럼은 만들지 말아줘\n" +
              "\n" +
              "[출력 형식]\n" +
              "- 기존 단계는 그대로 두고, 마지막에 새 단계만 추가한 전체 M코드를 처음부터 끝까지 다시 작성해줘\n" +
              "- 추가하는 단계 이름은 한글로, 무슨 작업을 하는 단계인지 알 수 있게 지어줘",
            promptHighlight: '여기에 위 "결과 M코드"를 붙여넣기',
            promptHighlightClass: "prompt-highlight-alt",
          },
          {
            heading: "결과 M코드",
            note: "날짜/금액 타입 변환과 오류 개수 처리까지 반영된 최종 코드입니다",
            prompt: String.raw`let
    // 1. 지정한 폴더에 있는 모든 파일 목록 불러오기
    폴더파일불러오기 = Folder.Files("C:\Users\Haein\Desktop\교재_업무자동화\경비지출 폴더"),

    // 2. 확장자가 .xlsx인 파일만 남기기
    엑셀파일만선택 = Table.SelectRows(폴더파일불러오기, each Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~$")),

    // 3. 숨김 파일 제외하기
    숨김파일제외 = Table.SelectRows(엑셀파일만선택, each [Attributes]?[Hidden]? <> true),

    // 4. 각 엑셀 파일의 내용을 읽기
    엑셀내용읽기 = Table.AddColumn(숨김파일제외, "엑셀내용", each Excel.Workbook([Content], false)),

    // 5. 각 파일에서 첫 번째 시트의 데이터만 가져오기
    첫번째시트가져오기 = Table.AddColumn(엑셀내용읽기, "첫번째시트", each let 시트목록 = Table.SelectRows([엑셀내용], each [Kind] = "Sheet") in if Table.RowCount(시트목록) > 0 then 시트목록{0}[Data] else null),

    // 6. 첫 번째 시트가 없는 파일은 제외하기
    시트없는파일제외 = Table.SelectRows(첫번째시트가져오기, each [첫번째시트] <> null),

    // 7. 각 시트의 첫 번째 행을 머리글로 사용하기
    머리글적용 = Table.TransformColumns(시트없는파일제외, {"첫번째시트", each Table.PromoteHeaders(_, [PromoteAllScalars = true])}),

    // 8. 모든 파일의 첫 번째 시트를 하나의 표로 합치기
    모든파일통합 = Table.Combine(머리글적용[첫번째시트]),

    // 9. 수량 컬럼에서 숫자만 추출하고 정수 형식으로 변환하기
    수량컬럼숫자추출 = Table.TransformColumns(모든파일통합, {{"수량", each if _ = null then null else Int64.From(Text.Select(Text.From(_), {"0".."9"})), Int64.Type}}),

    // 10. 날짜와 금액의 변환 실패 개수 확인하기
    변환오류개수추가 = Table.AddColumn(수량컬럼숫자추출, "변환오류개수", each
        (if [날짜] <> null and (try Date.From([날짜]))[HasError] then 1 else 0) +
        (if [금액] <> null and (try Number.From([금액]))[HasError] then 1 else 0),
        Int64.Type),

    // 11. 날짜와 금액 타입 변환하고 실패한 값은 null로 처리하기
    날짜금액타입변환 = Table.TransformColumns(변환오류개수추가, {
        {"날짜", each try Date.From(_) otherwise null, type date},
        {"금액", each try Number.From(_) otherwise null, type number}
    }),

    // 12. 전체 실패 건수가 0이면 오류 개수 컬럼 제거하기
    오류개수컬럼정리 = if List.Sum(날짜금액타입변환[변환오류개수]) = 0
        then Table.RemoveColumns(날짜금액타입변환, {"변환오류개수"})
        else 날짜금액타입변환

in
    오류개수컬럼정리`,
            promptHighlight: String.raw`C:\Users\Haein\Desktop\교재_업무자동화\경비지출 폴더`,
            promptHighlightClass: "path-highlight",
          },
        ],
      },
      {
        heading: "AI 프롬프트 — 부서코드 쿼리 병합 단계 추가",
        group: [
          {
            heading: "AI 프롬프트 — 부서코드 쿼리 병합 단계 추가",
            note: "위 결과 M코드 (최종)를 함께 붙여넣고 사용하세요",
            prompt:
              "너는 Microsoft Power Query(M 언어) 전문가야.\n" +
              "아래는 여러 엑셀 파일을 하나의 표로 통합한 Power Query M코드야.\n" +
              "\n" +
              "[기존 M코드]\n" +
              "여기에 위 \"결과 M코드 (최종)\"를 붙여넣기\n" +
              "\n" +
              "[요청 사항]\n" +
              "이 M코드의 마지막 단계로, 이미 파워쿼리에 로드되어 있는 \"부서코드\" 쿼리와 쿼리 병합(Merge)하는 단계를 추가해줘.\n" +
              "\n" +
              "[조건]\n" +
              "- \"부서코드\" 쿼리에는 \"부서명\", \"본부\" 컬럼이 있어\n" +
              "- 기존 M코드의 \"부서명\" 컬럼과 \"부서코드\" 쿼리의 \"부서명\" 컬럼을 기준으로 병합해줘\n" +
              "- 조인 방식은 왼쪽 외부 조인(Left Outer Join)으로 해줘\n" +
              "- 병합 결과에서 \"본부\" 컬럼만 확장(Expand)해서 가져와줘\n" +
              "\n" +
              "[출력 형식]\n" +
              "- 기존 단계는 그대로 두고, 마지막에 새 단계만 추가한 전체 M코드를 처음부터 끝까지 다시 작성해줘\n" +
              "- 추가하는 단계 이름은 한글로, 무슨 작업을 하는 단계인지 알 수 있게 지어줘",
            promptHighlight: '여기에 위 "결과 M코드 (최종)"를 붙여넣기',
            promptHighlightClass: "prompt-highlight-alt",
          },
          {
            heading: "결과 M코드",
            note: "부서코드 쿼리 병합 단계까지 반영된 코드입니다",
            prompt: String.raw`let
    // 1. 지정한 폴더에 있는 모든 파일 목록 불러오기
    폴더파일불러오기 = Folder.Files("C:\Users\Haein\Desktop\교재_업무자동화\경비지출 폴더"),

    // 2. 확장자가 .xlsx인 파일만 남기기
    엑셀파일만선택 = Table.SelectRows(폴더파일불러오기, each Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~$")),

    // 3. 숨김 파일 제외하기
    숨김파일제외 = Table.SelectRows(엑셀파일만선택, each [Attributes]?[Hidden]? <> true),

    // 4. 각 엑셀 파일의 내용을 읽기
    엑셀내용읽기 = Table.AddColumn(숨김파일제외, "엑셀내용", each Excel.Workbook([Content], false)),

    // 5. 각 파일에서 첫 번째 시트의 데이터만 가져오기
    첫번째시트가져오기 = Table.AddColumn(엑셀내용읽기, "첫번째시트", each let 시트목록 = Table.SelectRows([엑셀내용], each [Kind] = "Sheet") in if Table.RowCount(시트목록) > 0 then 시트목록{0}[Data] else null),

    // 6. 첫 번째 시트가 없는 파일은 제외하기
    시트없는파일제외 = Table.SelectRows(첫번째시트가져오기, each [첫번째시트] <> null),

    // 7. 각 시트의 첫 번째 행을 머리글로 사용하기
    머리글적용 = Table.TransformColumns(시트없는파일제외, {"첫번째시트", each Table.PromoteHeaders(_, [PromoteAllScalars = true])}),

    // 8. 모든 파일의 첫 번째 시트를 하나의 표로 합치기
    모든파일통합 = Table.Combine(머리글적용[첫번째시트]),

    // 9. 수량 컬럼에서 숫자만 추출하고 정수 형식으로 변환하기
    수량컬럼숫자추출 = Table.TransformColumns(모든파일통합, {{"수량", each if _ = null then null else Int64.From(Text.Select(Text.From(_), {"0".."9"})), Int64.Type}}),

    // 10. 날짜와 금액의 변환 실패 개수 확인하기
    변환오류개수추가 = Table.AddColumn(수량컬럼숫자추출, "변환오류개수", each
        (if [날짜] <> null and (try Date.From([날짜]))[HasError] then 1 else 0) +
        (if [금액] <> null and (try Number.From([금액]))[HasError] then 1 else 0),
        Int64.Type),

    // 11. 날짜와 금액 타입 변환하고 실패한 값은 null로 처리하기
    날짜금액타입변환 = Table.TransformColumns(변환오류개수추가, {
        {"날짜", each try Date.From(_) otherwise null, type date},
        {"금액", each try Number.From(_) otherwise null, type number}
    }),

    // 12. 전체 실패 건수가 0이면 오류 개수 컬럼 제거하기
    오류개수컬럼정리 = if List.Sum(날짜금액타입변환[변환오류개수]) = 0
        then Table.RemoveColumns(날짜금액타입변환, {"변환오류개수"})
        else 날짜금액타입변환,

    // 13. 현재 엑셀 파일의 부서코드 표 불러오기
    부서코드불러오기 = Excel.CurrentWorkbook(){[Name="부서코드"]}[Content],

    // 14. 부서명을 기준으로 부서코드 표와 Left Outer Join
    부서정보병합 = Table.NestedJoin(오류개수컬럼정리, {"부서명"}, 부서코드불러오기, {"부서명"}, "부서정보", JoinKind.LeftOuter),

    // 15. 부서코드 표에서 본부 컬럼만 가져오기
    본부컬럼추가 = Table.ExpandTableColumn(부서정보병합, "부서정보", {"본부"}, {"본부"})

in
    본부컬럼추가`,
          },
        ],
      },
      {
        heading: "AI 프롬프트 — 중분류코드 쿼리 병합 단계 추가",
        group: [
          {
            heading: "AI 프롬프트 — 중분류코드 쿼리 병합 단계 추가",
            note: "위 결과 M코드를 함께 붙여넣고 사용하세요",
            prompt:
              "너는 Microsoft Power Query(M 언어) 전문가야.\n" +
              "\n" +
              "아래 기존 M코드에 \"중분류코드\" 시트의 정보를 병합하는 단계를 추가해줘.\n" +
              "\n" +
              "[병합 조건]\n" +
              "\n" +
              "- 현재 쿼리의 \"중분류\" 컬럼과 \"중분류코드\" 시트의 \"중분류\" 컬럼을 기준으로 병합해줘.\n" +
              "- 두 컬럼의 값이 같은 행끼리 연결해줘.\n" +
              "- \"중분류코드\" 시트에서는 \"대분류\" 컬럼만 가져와줘.\n" +
              "- 기존 데이터의 모든 행이 유지되도록 Left Outer Join을 사용해줘.\n" +
              "- \"중분류\" 값이 일치하지 않으면 \"대분류\" 값은 null로 남겨줘.\n" +
              "- \"중분류코드\" 시트의 데이터는 현재 엑셀 파일 안에 있고, Excel 표(Table) 이름도 \"중분류코드\"라고 가정해줘.\n" +
              "\n" +
              "[출력 형식]\n" +
              "\n" +
              "- 기존 M코드는 수정하지 말고 마지막에 병합 단계를 추가해줘.\n" +
              "- 추가하는 단계 이름은 한글로 작성해줘.\n" +
              "- 완성된 전체 M코드를 처음부터 끝까지 보여줘.",
          },
          {
            heading: "결과 M코드",
            note: "중분류코드 쿼리 병합 단계까지 반영된 코드입니다",
            prompt: String.raw`let
    // 1. 지정한 폴더에 있는 모든 파일 목록 불러오기
    폴더파일불러오기 = Folder.Files("C:\Users\Haein\Desktop\교재_업무자동화\실습자료\경비지출 폴더"),

    // 2. 확장자가 .xlsx인 파일만 남기기
    엑셀파일만선택 = Table.SelectRows(폴더파일불러오기, each Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~$")),

    // 3. 숨김 파일 제외하기
    숨김파일제외 = Table.SelectRows(엑셀파일만선택, each [Attributes]?[Hidden]? <> true),

    // 4. 각 엑셀 파일의 내용을 읽기
    엑셀내용읽기 = Table.AddColumn(숨김파일제외, "엑셀내용", each Excel.Workbook([Content], false)),

    // 5. 각 파일에서 첫 번째 시트의 데이터만 가져오기
    첫번째시트가져오기 = Table.AddColumn(엑셀내용읽기, "첫번째시트", each let 시트목록 = Table.SelectRows([엑셀내용], each [Kind] = "Sheet") in if Table.RowCount(시트목록) > 0 then 시트목록{0}[Data] else null),

    // 6. 첫 번째 시트가 없는 파일은 제외하기
    시트없는파일제외 = Table.SelectRows(첫번째시트가져오기, each [첫번째시트] <> null),

    // 7. 각 시트의 첫 번째 행을 머리글로 사용하기
    머리글적용 = Table.TransformColumns(시트없는파일제외, {"첫번째시트", each Table.PromoteHeaders(_, [PromoteAllScalars = true])}),

    // 8. 모든 파일의 첫 번째 시트를 하나의 표로 합치기
    모든파일통합 = Table.Combine(머리글적용[첫번째시트]),

    // 9. 수량 컬럼에서 숫자만 추출하고 정수 형식으로 변환하기
    수량컬럼숫자추출 = Table.TransformColumns(모든파일통합, {{"수량", each if _ = null then null else Int64.From(Text.Select(Text.From(_), {"0".."9"})), Int64.Type}}),

    // 10. 날짜와 금액의 변환 실패 개수 확인하기
    변환오류개수추가 = Table.AddColumn(수량컬럼숫자추출, "변환오류개수", each
        (if [날짜] <> null and (try Date.From([날짜]))[HasError] then 1 else 0) +
        (if [금액] <> null and (try Number.From([금액]))[HasError] then 1 else 0),
        Int64.Type),

    // 11. 날짜와 금액 타입 변환하고 실패한 값은 null로 처리하기
    날짜금액타입변환 = Table.TransformColumns(변환오류개수추가, {
        {"날짜", each try Date.From(_) otherwise null, type date},
        {"금액", each try Number.From(_) otherwise null, type number}
    }),

    // 12. 전체 실패 건수가 0이면 오류 개수 컬럼 제거하기
    오류개수컬럼정리 = if List.Sum(날짜금액타입변환[변환오류개수]) = 0
        then Table.RemoveColumns(날짜금액타입변환, {"변환오류개수"})
        else 날짜금액타입변환,

    // 13. 현재 엑셀 파일의 부서코드 표 불러오기
    부서코드불러오기 = Excel.CurrentWorkbook(){[Name="부서코드"]}[Content],

    // 14. 부서명을 기준으로 부서코드 표와 병합하기
    부서정보병합 = Table.NestedJoin(오류개수컬럼정리, {"부서명"}, 부서코드불러오기, {"부서명"}, "부서정보", JoinKind.LeftOuter),

    // 15. 부서코드 표에서 본부 컬럼 가져오기
    본부컬럼추가 = Table.ExpandTableColumn(부서정보병합, "부서정보", {"본부"}, {"본부"}),

    // 16. 현재 엑셀 파일의 중분류코드 표 불러오기
    중분류코드불러오기 = Excel.CurrentWorkbook(){[Name="중분류코드"]}[Content],

    // 17. 중분류를 기준으로 중분류코드 표와 병합하기
    중분류정보병합 = Table.NestedJoin(본부컬럼추가, {"중분류"}, 중분류코드불러오기, {"중분류"}, "중분류정보", JoinKind.LeftOuter),

    // 18. 중분류코드 표에서 대분류 컬럼 가져오기
    대분류컬럼추가 = Table.ExpandTableColumn(중분류정보병합, "중분류정보", {"대분류"}, {"대분류"})

in
    대분류컬럼추가`,
          },
        ],
      },
    ],
  },
  practice1: {
    title: "빈셀채우기",
    subtitle: "엑셀VBA",
    noFold: true,
    sections: [
      {
        heading: "실습 자료",
        callout: {
          icon: "📎",
          paragraphs: ['실습자료 폴더의 "실습1_빈셀채우기.xlsx"'],
        },
      },
      {
        heading: "AI 프롬프트 — 프롬프트 생성 요청",
        note: "첨부 이미지와 함께 붙여넣으면 더 정확한 프롬프트를 받을 수 있습니다",
        prompt:
          "부서명 열에서 비어있는 셀 값을 위의 셀의 값으로 채우는 vba코드를 작성하려고 해.\n" +
          "프롬프트를 작성해줘",
      },
      {
        heading: "AI 프롬프트 — 빈 셀을 위의 셀 값으로 채우기",
        note: "복사해서 ChatGPT, Claude 등에 붙여넣고 사용하세요",
        prompt:
          "엑셀 VBA 매크로를 작성해줘.\n" +
          "\n" +
          "[상황]\n" +
          "A열(부서명)에 같은 부서가 연속될 때 첫 행에만 값이 있고, 나머지 행은 빈 셀로 되어 있어.\n" +
          "(예: A2=총무과, A3=공백, A4=공백, A5=민원과, A6=공백 ...)\n" +
          "\n" +
          "[원하는 동작]\n" +
          "- A열에서 빈 셀을 찾아, 바로 위 셀의 값으로 채워줘\n" +
          "- 데이터가 몇 행까지 있는지 몰라서, 마지막 데이터 행까지 자동으로 인식해서 처리해줘 (행 번호 하드코딩 금지)\n" +
          "- B열, C열 등 다른 열은 건드리지 마\n" +
          "- 실행 완료 후 몇 개의 셀을 채웠는지 메시지박스로 알려줘\n" +
          "\n" +
          "[참고]\n" +
          "- 병합된 셀이 아니라 실제로 값이 비어 있는 빈 셀 상태야",
      },
      {
        heading: "결과 코드",
        prompt: String.raw`Sub FillBlankDeptCells()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fillCount As Long

    Set ws = ActiveSheet

    ' A열 기준 마지막 데이터 행 자동 인식 (하드코딩 없음)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ' A열만으로는 부족할 수 있으니 B열 기준도 함께 확인해서 더 큰 값 사용
    Dim lastRowB As Long
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRowB > lastRow Then lastRow = lastRowB

    fillCount = 0

    ' 2행부터 시작 (1행은 헤더로 가정)
    For i = 2 To lastRow
        If ws.Cells(i, "A").Value = "" Then
            ws.Cells(i, "A").Value = ws.Cells(i - 1, "A").Value
            fillCount = fillCount + 1
        End If
    Next i

    MsgBox fillCount & "개의 빈 셀을 채웠습니다.", vbInformation, "완료"
End Sub`,
      },
      {
        heading: "실행 방법",
        callout: {
          icon: "▶️",
          steps: [
            "Alt + F11로 VBA 편집기 열기",
            "삽입 → 모듈 클릭",
            "AI가 생성한 코드 붙여넣기",
            "F5 또는 Ctrl + F5로 실행",
          ],
        },
      },
    ],
  },
  practice2: {
    title: "셀 병합하기",
    subtitle: "엑셀VBA",
    noFold: true,
    sections: [
      {
        heading: "실습 자료",
        callout: {
          icon: "📎",
          paragraphs: ['실습자료 폴더의 "실습2_셀병합히기_실습.xlsx"'],
        },
      },
      {
        heading: "AI 프롬프트 — 프롬프트 생성 요청",
        note: "첨부 이미지와 함께 붙여넣으면 더 정확한 프롬프트를 받을 수 있습니다",
        prompt:
          "첨부이미지의 A열(부서명)에서 같은 값이 연속되는 셀을 병합하는 vba코드가 나오도록\n" +
          "프롬프트를 작성해줘",
      },
      {
        heading: "AI 프롬프트 — 같은 값 셀 병합하기",
        note: "복사해서 ChatGPT, Claude 등에 붙여넣고 사용하세요",
        prompt:
          "엑셀 VBA 매크로를 작성해줘.\n" +
          "\n" +
          "[목표]\n" +
          "부서명 열에서 같은 값이 연속되는 셀을 하나로 병합하는 VBA 코드를 작성한다.\n" +
          "\n" +
          "[상황]\n" +
          "A열(부서명)에 같은 값이 연속으로 여러 행에 걸쳐 입력되어 있어.\n" +
          "(예: A2~A4=총무과, A5~A8=민원과, A9~A10=세무과, A11~A13=환경과 ...)\n" +
          "현재는 병합된 셀이 아니라, 각 행마다 값이 개별적으로 입력되어 있는 상태야.\n" +
          "\n" +
          "[원하는 동작]\n" +
          "- A열에서 값이 같은 연속된 셀들을 찾아서 하나로 병합해줘\n" +
          "- 병합된 셀의 값은 맨 위 셀의 값으로 표시되고, 세로/가로 가운데 정렬해줘\n" +
          "- 데이터가 몇 행까지 있는지 몰라서, 마지막 데이터 행까지 자동으로 인식해서 처리해줘 (행 번호 하드코딩 금지)\n" +
          "- B열, C열 등 다른 열은 건드리지 마\n" +
          "\n" +
          "[참고]\n" +
          "- 헤더는 1행이고, 데이터는 2행부터 시작해\n" +
          "- 부서가 바뀌는 지점(값이 달라지는 지점)에서 병합 구간이 끊겨야 해",
      },
      {
        heading: "결과 코드",
        prompt: String.raw`Sub 부서명_연속셀_병합()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim i As Long

    Set ws = ActiveSheet

    ' A열의 마지막 데이터 행 자동 확인
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 데이터가 없는 경우 종료
    If lastRow < 2 Then
        MsgBox "병합할 데이터가 없습니다.", vbInformation
        Exit Sub
    End If

    ' 병합 경고 메시지 잠시 끄기
    Application.DisplayAlerts = False

    startRow = 2

    ' 3행부터 비교
    ' 마지막 구간 처리를 위해 lastRow + 1까지 반복
    For i = 3 To lastRow + 1

        ' 이전 행과 값이 달라지면 하나의 구간 종료
        If i > lastRow Or ws.Cells(i, "A").Value <> ws.Cells(startRow, "A").Value Then

            ' 같은 값이 2행 이상 연속된 경우에만 병합
            If i - startRow >= 2 Then

                With ws.Range("A" & startRow & ":A" & i - 1)
                    .Merge

                    ' 가로 가운데 정렬
                    .HorizontalAlignment = xlCenter

                    ' 세로 가운데 정렬
                    .VerticalAlignment = xlCenter
                End With

            Else
                ' 한 셀짜리 구간도 가운데 정렬
                With ws.Cells(startRow, "A")
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
            End If

            ' 다음 부서의 시작 행
            startRow = i

        End If

    Next i

    Application.DisplayAlerts = True

    MsgBox "부서명 병합이 완료되었습니다.", vbInformation

End Sub`,
      },
      {
        heading: "실행 방법",
        callout: {
          icon: "▶️",
          steps: [
            "Alt + F11로 VBA 편집기 열기",
            "삽입 → 모듈 클릭",
            "AI가 생성한 코드 붙여넣기",
            "F5 또는 Ctrl + F5로 실행",
          ],
        },
      },
    ],
  },
  practice3: {
    title: "특정조건데이터 구분하기",
    subtitle: "엑셀VBA",
    noFold: true,
    sections: [
      {
        heading: "실습 자료",
        callout: {
          icon: "📎",
          paragraphs: [
            '실습자료 폴더의 "실습3_특정조건을만족하는데이터구분_실습.xlsx"',
          ],
        },
      },
      {
        heading: "AI 프롬프트 — 프롬프트 생성 요청",
        note: "첨부 이미지와 함께 붙여넣으면 더 정확한 프롬프트를 받을 수 있습니다",
        prompt:
          "계약만료일 기준으로 상태 표시 열을 추가하고, 만료일-오늘 <= 0이면 빨강, <= 30이면 노랑으로 행 색상을\n" +
          "표시하는 vba코드가 나오도록 프롬프트를 작성해줘",
      },
      {
        heading: "AI 프롬프트 — 특정 조건 데이터 강조 표시",
        note: "복사해서 ChatGPT, Claude 등에 붙여넣고 사용하세요",
        prompt:
          "엑셀 VBA 매크로를 작성해줘.\n" +
          "\n" +
          "[목표]\n" +
          "계약만료일을 기준으로 상태를 표시하는 열을 추가하고, 조건에 따라 행 색상을 다르게 표시하는 VBA 코드를 작성한다.\n" +
          "\n" +
          "[상황]\n" +
          "E열에 계약만료일이 날짜로 입력되어 있어.\n" +
          "데이터는 2행부터 시작하고, 마지막 행은 자동으로 인식해줘.\n" +
          "\n" +
          "[원하는 동작]\n" +
          "1. F열에 \"상태\"라는 헤더를 추가하고, 각 행마다 아래 기준으로 상태 값을 표시해줘\n" +
          "- 계약만료일 - 오늘 <= 0 → \"만료됨\"\n" +
          "- 계약만료일 - 오늘 <= 30 (0보다 큰 경우) → \"30일 이내\"\n" +
          "- 그 외 → \"여유\"\n" +
          "2. 상태에 따라 해당 행 전체(A열~F열)의 배경색을 다르게 칠해줘\n" +
          "- \"만료됨\" → 빨강 계열\n" +
          "- \"30일 이내\" → 노랑 계열\n" +
          "- \"여유\" → 색 없음(기본 흰색)\n" +
          "3. 매크로를 다시 실행해도 이전 색상이 남지 않도록, 실행 시작 시 기존 배경색을 먼저 초기화해줘\n" +
          "\n" +
          "[참고]\n" +
          "- 헤더는 1행이고, 데이터는 2행부터 시작해\n" +
          "- \"오늘\" 기준은 VBA의 Date 함수를 사용해줘",
      },
      {
        heading: "결과 코드",
        prompt: String.raw`Sub 계약만료_상태표시()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim expireDate As Date
    Dim remainDays As Long

    Set ws = ActiveSheet

    ' E열을 기준으로 마지막 데이터 행 자동 확인
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' 데이터가 없는 경우 종료
    If lastRow < 2 Then
        MsgBox "처리할 데이터가 없습니다.", vbInformation
        Exit Sub
    End If

    ' F열 헤더 입력
    ws.Range("F1").Value = "상태"

    ' 기존 배경색 초기화
    ws.Range("A2:F" & lastRow).Interior.Pattern = xlNone

    ' 2행부터 마지막 행까지 반복
    For i = 2 To lastRow

        ' E열에 정상적인 날짜가 입력되어 있는 경우
        If IsDate(ws.Cells(i, "E").Value) Then

            expireDate = CDate(ws.Cells(i, "E").Value)

            ' 계약만료일까지 남은 날짜 계산
            remainDays = expireDate - Date

            If remainDays <= 0 Then

                ' 이미 만료되었거나 오늘이 만료일
                ws.Cells(i, "F").Value = "만료됨"

                ' A~F열을 연한 빨강으로 표시
                ws.Range("A" & i & ":F" & i).Interior.Color = RGB(255, 199, 206)

            ElseIf remainDays <= 30 Then

                ' 만료일까지 1~30일 남은 경우
                ws.Cells(i, "F").Value = "30일 이내"

                ' A~F열을 연한 노랑으로 표시
                ws.Range("A" & i & ":F" & i).Interior.Color = RGB(255, 235, 156)

            Else

                ' 30일보다 많이 남은 경우
                ws.Cells(i, "F").Value = "여유"

                ' 별도 색상 없음

            End If

        Else

            ' E열이 비어 있거나 날짜가 아닌 경우
            ws.Cells(i, "F").Value = ""

        End If

    Next i

    MsgBox "계약 상태 확인이 완료되었습니다.", vbInformation

End Sub`,
      },
      {
        heading: "실행 방법",
        callout: {
          icon: "▶️",
          steps: [
            "Alt + F11로 VBA 편집기 열기",
            "삽입 → 모듈 클릭",
            "AI가 생성한 코드 붙여넣기",
            "F5 또는 Ctrl + F5로 실행",
          ],
        },
      },
      {
        heading: "엑셀을 열면 자동실행되게 하는 방법",
        callout: {
          icon: "📌",
          steps: [
            "앞에서 만든 HighlightContractStatus 코드는 그대로 표준 모듈에 둡니다 (실제 로직)",
            '<pre class="step-code">Private Sub Workbook_Open()\n    Call HighlightContractStatus\nEnd Sub</pre>',
            '현재 통합문서 모듈(VBA 편집기 왼쪽 프로젝트 탐색기에서 "Microsoft Excel 개체 → 현재통합문서")에 Workbook_Open 코드를 추가합니다',
            "코드를 넣은 후 Ctrl + S로 저장 (매크로가 있으니 .xlsm 형식 유지 확인)",
            "파일을 완전히 닫고 다시 파일을 열기",
            "별도로 F5를 누르지 않아도 F열에 상태값이 채워지고 행 색깔이 자동으로 칠해지면 성공입니다",
          ],
        },
      },
    ],
  },
  certificate: {
    title: "수료증 출력",
    subtitle: "엑셀VBA",
    noFold: true,
    sections: [
      {
        heading: "실습 자료",
        callout: {
          icon: "📎",
          paragraphs: [
            '실습자료 폴더의 "업무자동화_수료증생성_실습.xlsm" (프롬프트 시트 참고)',
          ],
        },
      },
      {
        heading: "AI 프롬프트 — 방법론 요청",
        note: "먼저 AI에게 어떻게 접근하면 좋을지 물어봅니다",
        prompt:
          "[역할]\n" +
          "너는 엑셀 VBA 작성 전문가야.\n" +
          "나는 VBA를 전혀 모르는 초보 사용자야.\n" +
          "어떤식으로 자동화할 지 알려줘\n" +
          "\n" +
          "[업무 목표]\n" +
          "엑셀 파일 안에 있는 \"명단\" 시트의 데이터를 한 행씩 읽어서,\n" +
          "수료증 시트의 이미지 수료증 양식 위에 있는 내용을 바꾸고,\n" +
          "사람별로 수료증을 PDF로 하나씩 저장하고 싶어.\n" +
          "\n" +
          "[현재 상황]\n" +
          "수료증 시트에는 수료증 디자인이 이미지로 들어가 있음",
      },
      {
        heading: "AI 프롬프트 — 프롬프트 요청",
        note: "수료증, 명단 시트의 이미지를 함께 첨부해서 사용하세요",
        prompt:
          "[역할]\n" +
          "너는 엑셀 VBA 작성 전문가야.\n" +
          "나는 VBA를 전혀 모르는 사용자야\n" +
          "업무 설명을 할테니 VBA 코드가 나오도록 프롬프트를 작성해줘\n" +
          "\n" +
          "[업무설명]\n" +
          "시트명 : 수료증\n" +
          "시트명 : 명단\n" +
          "출력명 : 명단에 있는 사람들의 정보가 상장에 들어가 사람별로 하나씩 수료증이 PDF로 저장되도롤 하고 싶어.\n" +
          "이미지 파일을 참고해서 어떻게 하면 되는지 말해줘.",
      },
      {
        heading: "AI 프롬프트 — VBA 코드 생성",
        note: "복사해서 ChatGPT, Claude 등에 붙여넣고 사용하세요",
        prompt:
          "[역할]\n" +
          "너는 엑셀 VBA 작성 전문가야.\n" +
          "나는 VBA를 전혀 모르는 초보 사용자야.\n" +
          "지금부터 내가 설명하는 엑셀 파일 구조와 작업 목표를 바탕으로, 바로 복사해서 사용할 수 있는 실무용 VBA 코드를 작성해줘.\n" +
          "\n" +
          "[목표]\n" +
          "엑셀 파일 안의 \"명단\" 시트 데이터를 한 행씩 읽어서,\n" +
          "수료증 시트에 있는 텍스트박스 내용을 사람별로 바꾸고,\n" +
          "각 사람의 수료증을 PDF로 하나씩 자동 저장하는 매크로를 만들고 싶어.\n" +
          "\n" +
          "[엑셀 파일 구조]\n" +
          "\n" +
          "시트명: 명단\n" +
          "1행은 제목\n" +
          "2행부터 데이터가 있음\n" +
          "열 구성은 다음과 같음\n" +
          "A열: 성명\n" +
          "B열: 소속\n" +
          "C열: 교육명\n" +
          "D열: 교육기간\n" +
          "E열: 날짜\n" +
          "시트명: 수료증\n" +
          "수료증 배경은 이미지로 되어 있음\n" +
          "변경되는 값은 텍스트박스로 만들어져 있음\n" +
          "텍스트박스 이름은 다음과 같음\n" +
          "성명: txtName\n" +
          "소속: txtDept\n" +
          "교육명: txtCourse\n" +
          "교육기간: txtPeriod\n" +
          "날짜: txtDate\n" +
          "\n" +
          "[원하는 동작]\n" +
          "\n" +
          "명단 시트의 2행부터 마지막 행까지 반복\n" +
          "각 행의 값을 수료증 시트의 텍스트박스에 넣기\n" +
          "수료증 시트를 PDF로 저장하기\n" +
          "저장 파일명은 \"성명_수료증.pdf\" 형식으로 하기\n" +
          "PDF는 현재 엑셀 파일과 같은 폴더 안에 \"수료증 PDF\" 폴더를 만들어 저장하기\n" +
          "폴더가 없으면 자동 생성하기\n" +
          "폴더가 이미 있으면 그대로 사용하기\n" +
          "PDF 저장 후 다음 사람으로 넘어가기\n" +
          "모든 작업이 끝나면 완료 메시지를 띄우기\n" +
          "\n" +
          "[추가 조건]\n" +
          "\n" +
          "초보자도 이해할 수 있도록 모든 주요 코드에 자세한 주석을 달아줘\n" +
          "왜 이 코드가 필요한지도 짧게 설명해줘\n" +
          "참조 설정 없이 실행 가능하도록 작성해줘\n" +
          "외부 라이브러리 없이 엑셀 기본 기능만 사용해줘\n" +
          "오류가 발생할 수 있는 경우를 고려해서 안정적으로 작성해줘\n" +
          "오류 발생 시 MsgBox로 원인과 조치 방법을 안내해줘\n" +
          "시트를 Select하거나 Activate하지 말고 객체로 직접 참조해줘\n" +
          "Windows 환경 기준으로 작성해줘\n" +
          "날짜 값은 수료증에 \"2026년 4월 1일\" 형식으로 들어가게 해줘\n" +
          "파일명에 사용할 수 없는 문자(\\ / : * ? \" < > |)는 자동으로 제거하거나 바꿔줘\n" +
          "PDF 저장 전에 텍스트박스가 비어 있으면 경고가 뜨도록 해줘\n" +
          "명단 시트에 성명이 없는 행은 건너뛰게 해줘\n" +
          "\n" +
          "[출력 형식]\n" +
          "다음 순서로 답변해줘\n" +
          "\n" +
          "전체 VBA 코드\n" +
          "코드를 어디에 넣는지 초보자용 설명\n" +
          "실행 방법\n" +
          "자주 발생하는 오류와 해결 방법",
      },
      {
        heading: "결과 코드",
        prompt: String.raw`Option Explicit

'=========================================================
' 명단 시트의 데이터를 이용해 사람별 수료증 PDF를 생성하는 매크로
'
' [명단 시트]
' A열 : 성명
' B열 : 소속
' C열 : 교육명
' D열 : 교육기간
' E열 : 날짜
'
' [수료증 시트 텍스트박스 이름]
' txtName   : 성명
' txtDept   : 소속
' txtCourse : 교육명
' txtPeriod : 교육기간
' txtDate   : 날짜
'=========================================================
Sub 수료증_PDF_일괄생성()

    '-----------------------------------------------------
    ' 변수 선언
    '-----------------------------------------------------
    Dim wb As Workbook
    Dim wsList As Worksheet
    Dim wsCert As Worksheet

    Dim lastRow As Long
    Dim i As Long

    Dim personName As String
    Dim dept As String
    Dim courseName As String
    Dim period As String
    Dim certDate As String

    Dim saveFolder As String
    Dim saveFile As String
    Dim safeName As String

    Dim createdCount As Long
    Dim skippedCount As Long

    Dim blankFields As String
    Dim answer As VbMsgBoxResult

    '오류가 발생하면 아래의 오류 처리 부분으로 이동
    On Error GoTo ErrorHandler

    '-----------------------------------------------------
    ' 1. 현재 매크로가 들어 있는 엑셀 파일 지정
    '-----------------------------------------------------
    Set wb = ThisWorkbook


    '-----------------------------------------------------
    ' 2. 엑셀 파일이 먼저 저장되어 있는지 확인
    '
    ' PDF를 "현재 엑셀 파일과 같은 폴더"에 저장해야 하므로
    ' 엑셀 파일 자체가 먼저 저장되어 있어야 합니다.
    '-----------------------------------------------------
    If wb.Path = "" Then
        MsgBox "현재 엑셀 파일이 아직 저장되지 않았습니다." & vbCrLf & vbCrLf & _
               "먼저 엑셀 파일을 원하는 폴더에 저장한 후" & vbCrLf & _
               "매크로를 다시 실행해 주세요.", _
               vbExclamation, "엑셀 파일 저장 필요"
        Exit Sub
    End If


    '-----------------------------------------------------
    ' 3. "명단", "수료증" 시트가 있는지 확인
    '-----------------------------------------------------
    If Not WorksheetExists(wb, "명단") Then
        MsgBox "'명단' 시트를 찾을 수 없습니다." & vbCrLf & vbCrLf & _
               "시트 이름이 정확히 '명단'인지 확인해 주세요.", _
               vbCritical, "시트 확인"
        Exit Sub
    End If

    If Not WorksheetExists(wb, "수료증") Then
        MsgBox "'수료증' 시트를 찾을 수 없습니다." & vbCrLf & vbCrLf & _
               "시트 이름이 정확히 '수료증'인지 확인해 주세요.", _
               vbCritical, "시트 확인"
        Exit Sub
    End If


    '시트 객체 지정
    'Select 또는 Activate를 사용하지 않고 직접 참조합니다.
    Set wsList = wb.Worksheets("명단")
    Set wsCert = wb.Worksheets("수료증")


    '-----------------------------------------------------
    ' 4. 수료증 시트에 필요한 텍스트박스가 모두 있는지 확인
    '-----------------------------------------------------
    If Not ShapeExists(wsCert, "txtName") Then
        MsgBox "수료증 시트에서 'txtName' 텍스트박스를 찾을 수 없습니다." & vbCrLf & _
               "텍스트박스의 이름을 확인해 주세요.", _
               vbCritical, "텍스트박스 확인"
        Exit Sub
    End If

    If Not ShapeExists(wsCert, "txtDept") Then
        MsgBox "수료증 시트에서 'txtDept' 텍스트박스를 찾을 수 없습니다." & vbCrLf & _
               "텍스트박스의 이름을 확인해 주세요.", _
               vbCritical, "텍스트박스 확인"
        Exit Sub
    End If

    If Not ShapeExists(wsCert, "txtCourse") Then
        MsgBox "수료증 시트에서 'txtCourse' 텍스트박스를 찾을 수 없습니다." & vbCrLf & _
               "텍스트박스의 이름을 확인해 주세요.", _
               vbCritical, "텍스트박스 확인"
        Exit Sub
    End If

    If Not ShapeExists(wsCert, "txtPeriod") Then
        MsgBox "수료증 시트에서 'txtPeriod' 텍스트박스를 찾을 수 없습니다." & vbCrLf & _
               "텍스트박스의 이름을 확인해 주세요.", _
               vbCritical, "텍스트박스 확인"
        Exit Sub
    End If

    If Not ShapeExists(wsCert, "txtDate") Then
        MsgBox "수료증 시트에서 'txtDate' 텍스트박스를 찾을 수 없습니다." & vbCrLf & _
               "텍스트박스의 이름을 확인해 주세요.", _
               vbCritical, "텍스트박스 확인"
        Exit Sub
    End If


    '-----------------------------------------------------
    ' 5. 명단 시트의 마지막 데이터 행 찾기
    '
    ' A열(성명)을 기준으로 마지막 행을 자동으로 찾습니다.
    '-----------------------------------------------------
    lastRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "'명단' 시트에 처리할 데이터가 없습니다." & vbCrLf & _
               "2행부터 명단을 입력해 주세요.", _
               vbExclamation, "데이터 없음"
        Exit Sub
    End If


    '-----------------------------------------------------
    ' 6. PDF 저장 폴더 만들기
    '
    ' 현재 엑셀 파일이 있는 폴더 아래에
    ' "수료증 PDF" 폴더를 사용합니다.
    '-----------------------------------------------------
    saveFolder = wb.Path & Application.PathSeparator & "수료증 PDF"

    '폴더가 없을 때만 새로 생성
    If Dir(saveFolder, vbDirectory) = "" Then
        MkDir saveFolder
    End If


    '-----------------------------------------------------
    ' 화면 갱신을 잠시 중지
    '
    ' 많은 사람의 PDF를 만들 때 화면이 계속 깜빡이는 것을
    ' 방지하고 처리 속도를 높이기 위한 설정입니다.
    '-----------------------------------------------------
    Application.ScreenUpdating = False


    '-----------------------------------------------------
    ' 7. 명단의 2행부터 마지막 행까지 반복
    '-----------------------------------------------------
    For i = 2 To lastRow

        '-------------------------------------------------
        ' 성명 읽기
        '-------------------------------------------------
        personName = Trim(CStr(wsList.Cells(i, "A").Value))


        '-------------------------------------------------
        ' 성명이 없는 행은 건너뜀
        '-------------------------------------------------
        If personName = "" Then
            skippedCount = skippedCount + 1
            GoTo NextPerson
        End If


        '-------------------------------------------------
        ' 현재 행의 데이터를 변수에 저장
        '-------------------------------------------------
        dept = Trim(CStr(wsList.Cells(i, "B").Value))
        courseName = Trim(CStr(wsList.Cells(i, "C").Value))
        period = Trim(CStr(wsList.Cells(i, "D").Value))


        '-------------------------------------------------
        ' 날짜 처리
        '
        ' E열의 값이 정상적인 날짜이면
        ' 예: 2026년 4월 1일
        ' 형식으로 변환합니다.
        '
        ' 정상적인 날짜가 아니면 빈 문자열로 처리합니다.
        '-------------------------------------------------
        If IsDate(wsList.Cells(i, "E").Value) Then

            certDate = Format(CDate(wsList.Cells(i, "E").Value), _
                              "yyyy년 m월 d일")
        Else
            certDate = ""
        End If


        '-------------------------------------------------
        ' 8. 수료증 시트의 텍스트박스 내용 변경
        '-------------------------------------------------
        wsCert.Shapes("txtName").TextFrame2.TextRange.Text = personName
        wsCert.Shapes("txtDept").TextFrame2.TextRange.Text = dept
        wsCert.Shapes("txtCourse").TextFrame2.TextRange.Text = courseName
        wsCert.Shapes("txtPeriod").TextFrame2.TextRange.Text = period
        wsCert.Shapes("txtDate").TextFrame2.TextRange.Text = certDate


        '-------------------------------------------------
        ' 9. 텍스트박스 중 비어 있는 값이 있는지 확인
        '
        ' 비어 있는 항목이 있으면 어떤 항목인지 알려주고
        ' PDF를 그대로 만들 것인지 사용자에게 확인합니다.
        '-------------------------------------------------
        blankFields = ""

        If Trim(wsCert.Shapes("txtName").TextFrame2.TextRange.Text) = "" Then
            blankFields = blankFields & "ㆍ성명" & vbCrLf
        End If

        If Trim(wsCert.Shapes("txtDept").TextFrame2.TextRange.Text) = "" Then
            blankFields = blankFields & "ㆍ소속" & vbCrLf
        End If

        If Trim(wsCert.Shapes("txtCourse").TextFrame2.TextRange.Text) = "" Then
            blankFields = blankFields & "ㆍ교육명" & vbCrLf
        End If

        If Trim(wsCert.Shapes("txtPeriod").TextFrame2.TextRange.Text) = "" Then
            blankFields = blankFields & "ㆍ교육기간" & vbCrLf
        End If

        If Trim(wsCert.Shapes("txtDate").TextFrame2.TextRange.Text) = "" Then
            blankFields = blankFields & "ㆍ날짜" & vbCrLf
        End If


        '비어 있는 텍스트박스가 하나라도 있을 경우
        If blankFields <> "" Then

            answer = MsgBox( _
                personName & "님의 수료증에 다음 항목이 비어 있습니다." & _
                vbCrLf & vbCrLf & _
                blankFields & vbCrLf & _
                "그래도 PDF를 저장하시겠습니까?" & vbCrLf & vbCrLf & _
                "[예] : 그대로 PDF 저장" & vbCrLf & _
                "[아니오] : 이 사람은 건너뛰기", _
                vbYesNo + vbExclamation, _
                "빈 항목 확인")

            '아니오를 선택하면 해당 사람을 건너뜀
            If answer = vbNo Then
                skippedCount = skippedCount + 1
                GoTo NextPerson
            End If

        End If


        '-------------------------------------------------
        ' 10. PDF 파일명 만들기
        '
        ' Windows 파일명으로 사용할 수 없는 문자는
        ' CleanFileName 함수에서 "_"로 바꿉니다.
        '
        ' 예:
        ' 홍길동 → 홍길동_수료증.pdf
        '-------------------------------------------------
        safeName = CleanFileName(personName)

        saveFile = saveFolder & Application.PathSeparator & _
                   safeName & "_수료증.pdf"


        '-------------------------------------------------
        ' 11. 수료증 시트를 PDF로 저장
        '
        ' Select/Activate 없이 해당 워크시트 객체에서
        ' 바로 ExportAsFixedFormat을 실행합니다.
        '-------------------------------------------------
        wsCert.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=saveFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False


        '생성된 PDF 개수 증가
        createdCount = createdCount + 1


NextPerson:

        '다음 사람으로 진행

    Next i


    '-----------------------------------------------------
    ' 12. 화면 갱신 다시 활성화
    '-----------------------------------------------------
    Application.ScreenUpdating = True


    '-----------------------------------------------------
    ' 13. 모든 작업 완료 메시지
    '-----------------------------------------------------
    MsgBox "수료증 PDF 생성이 완료되었습니다." & vbCrLf & vbCrLf & _
           "생성된 PDF : " & createdCount & "개" & vbCrLf & _
           "건너뛴 명단 : " & skippedCount & "개" & vbCrLf & vbCrLf & _
           "저장 위치:" & vbCrLf & _
           saveFolder, _
           vbInformation, "작업 완료"

    Exit Sub


'=========================================================
' 오류 처리
'=========================================================
ErrorHandler:

    '오류 발생 후에도 화면 갱신 기능을 원래대로 복원
    Application.ScreenUpdating = True

    MsgBox "수료증 PDF 생성 중 오류가 발생했습니다." & vbCrLf & vbCrLf & _
           "오류 번호: " & Err.Number & vbCrLf & _
           "오류 내용: " & Err.Description & vbCrLf & vbCrLf & _
           "다음 사항을 확인해 주세요." & vbCrLf & _
           "1. '명단'과 '수료증' 시트 이름이 정확한지" & vbCrLf & _
           "2. 텍스트박스 이름이 정확한지" & vbCrLf & _
           "3. 엑셀 파일이 먼저 저장되어 있는지" & vbCrLf & _
           "4. PDF 저장 폴더에 쓰기 권한이 있는지" & vbCrLf & _
           "5. 같은 PDF 파일이 다른 프로그램에서 열려 있지 않은지", _
           vbCritical, "오류 발생"

End Sub


'=========================================================
' 특정 이름의 워크시트가 존재하는지 확인하는 함수
'
' 예:
' WorksheetExists(ThisWorkbook, "명단")
'
' 있으면 True
' 없으면 False
'=========================================================
Private Function WorksheetExists(ByVal wb As Workbook, _
                                 ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = wb.Worksheets(sheetName)

    On Error GoTo 0

    WorksheetExists = Not ws Is Nothing

End Function


'=========================================================
' 특정 텍스트박스 또는 도형이 수료증 시트에
' 존재하는지 확인하는 함수
'
' 텍스트박스 이름이 잘못되어 있을 경우
' 매크로 실행 중 갑자기 멈추는 것을 방지합니다.
'=========================================================
Private Function ShapeExists(ByVal ws As Worksheet, _
                             ByVal shapeName As String) As Boolean

    Dim shp As Shape

    On Error Resume Next

    Set shp = ws.Shapes(shapeName)

    On Error GoTo 0

    ShapeExists = Not shp Is Nothing

End Function


'=========================================================
' Windows 파일명으로 사용할 수 없는 문자를
' "_" 문자로 바꾸는 함수
'
' 사용할 수 없는 문자:
' \ / : * ? " < > |
'
' 예:
' 홍길동/총무과
'
' ↓
'
' 홍길동_총무과
'=========================================================
Private Function CleanFileName(ByVal fileName As String) As String

    Dim invalidChars As Variant
    Dim ch As Variant

    'Windows 파일명에서 사용할 수 없는 문자 목록
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    '각 금지 문자를 "_"로 변경
    For Each ch In invalidChars
        fileName = Replace(fileName, ch, "_")
    Next ch

    '앞뒤의 불필요한 공백 제거
    fileName = Trim(fileName)

    'Windows에서는 파일명의 마지막이 마침표(.)이면 문제가 될 수 있으므로 제거
    Do While Right(fileName, 1) = "."
        fileName = Left(fileName, Len(fileName) - 1)
    Loop

    '혹시 이름이 전부 제거되어 빈 값이 된 경우 대비
    If fileName = "" Then
        fileName = "이름없음"
    End If

    CleanFileName = fileName

End Function`,
      },
      {
        heading: "실행 방법",
        callout: {
          icon: "▶️",
          steps: [
            "Alt + F11로 VBA 편집기 열기",
            "삽입 → 모듈 클릭",
            "AI가 생성한 코드 붙여넣기",
            "F5 또는 Ctrl + F5로 실행",
          ],
        },
      },
    ],
  },
  education: {
    title: "파워쿼리로 하기",
    subtitle: "직접 해보기 · 교육 수료자 현황",
    noFold: true,
    sections: [
      {
        heading: "목표",
        callout: {
          icon: "🎯",
          paragraphs: [
            "월별 법정의무교육 수료자 명단을 통합하고, 직제 기준으로 표준화해 분석 가능한 형태로 정리한다.",
          ],
        },
      },
      {
        heading: "업무 흐름",
        cards: [
          {
            icon: "🔗",
            title: "데이터 통합",
            desc: "여러 교육기수 데이터를 통합하여 일관된 기준으로 관리 가능한 데이터셋을 구축한다",
          },
          {
            icon: "🏷️",
            title: "직급 표준화",
            desc: "직급 정보를 급수 기준으로 표준화하여 조직별 분석 및 보고서 작성이 가능하도록 한다",
          },
          {
            icon: "📊",
            title: "정렬 및 구조화",
            desc: "직제 순서에 맞는 데이터 구조를 만들어 의사결정에 활용 가능한 형태로 정리한다",
          },
        ],
      },
      {
        heading: "파일 구성 · 도구",
        callout: {
          icon: "📁",
          paragraphs: [
            "교육현황 폴더에 월별 이수자 명단 파일 존재",
            "동일한 구조의 엑셀 파일이 반복적으로 누적됨",
            "사용 도구 → 엑셀 Power Query",
          ],
        },
      },
      {
        heading: "전처리 단계",
        group: [
          {
            heading: "1. 교육현황 폴더에서 파일 통합",
            callout: {
              icon: "📌",
              paragraphs: [
                "목적 → 월별 파일을 하나의 데이터로 자동 통합",
                "사용 기능 → 파워쿼리 > 데이터 가져오기 > 파일에서 > 폴더 / 파일 결합",
              ],
            },
          },
          {
            heading: "2. 직급 정보를 직제로 표준화",
            callout: {
              icon: "📌",
              paragraphs: [
                "목적 → 직급 텍스트를 규칙에 따라 직제로 변환",
                "사용 기능 → 파워쿼리 > 열 추가 > 조건 열 추가",
                "실제 직급 값은 \"지방시설주사\"처럼 [직렬명 + 직급명] 형태이므로, 조건에서 \"다음을 포함함\" 대신 \"다음으로 끝남\"을 사용하면 아래 표의 직급명과 그대로 매칭할 수 있습니다.",
                "조건 열에서 \"다음을 포함함\"을 사용할 때는, \"주사\"는 \"주사보\"에, \"서기\"는 \"서기관\"·\"서기보\"에 포함되므로 조건 순서를 반드시 긴 직급명(주사보, 서기관, 서기보)을 먼저, 짧은 직급명(주사, 서기)을 나중에 배치해야 합니다.",
              ],
            },
            table: {
              headers: ["직급", "직제"],
              rows: [
                ["관리관", "1급"],
                ["이사관", "2급"],
                ["부이사관", "3급"],
                ["서기관", "4급"],
                ["사무관", "5급"],
                ["주사", "6급"],
                ["주사보", "7급"],
                ["서기", "8급"],
                ["서기보", "9급"],
              ],
            },
          },
          {
            heading: "3. 직제 기준 정렬",
            callout: {
              icon: "📌",
              paragraphs: [
                "목적 → 조직 체계에 맞는 순서로 정렬",
                "사용 기능 → 파워쿼리 > 열 정렬(Sort) / 생성된 급수 컬럼 기준 정렬",
              ],
            },
          },
        ],
      },
      {
        heading: "STEP 4 — 분석 시트 만들기 (미션)",
        group: [
          {
            heading: "미션",
            callout: {
              icon: "🎯",
              paragraphs: [
                "정리된 데이터를 바탕으로 피벗테이블을 만들어서, 직제별 · 직급별 · 소속별 이수자 수를 집계한 \"분석\" 시트를 만들어보세요.",
                "완성 결과는 아래 이미지처럼 3개의 표가 한 시트에 나란히 배치되어야 합니다.",
              ],
            },
          },
          {
            heading: "결과 예시",
            image: "education-analysis.png",
          },
          {
            heading: "힌트",
            callout: {
              icon: "💡",
              steps: [
                "정리된 데이터 범위를 선택한 뒤 [삽입] &gt;&gt; [피벗 테이블]로 표를 만들어보세요",
                "행 레이블에는 직제, 직급, 소속을 각각 넣고, 값에는 \"번호\" 필드를 \"개수\"로 집계해보세요",
                "직제별 / 직급별 / 소속별로 총 3개의 피벗테이블을 만들어서 한 시트에 나란히 배치해보세요",
                "각 표는 인원수가 많은 순서로 정렬하면 이미지와 더 비슷해져요 (값 기준 내림차순 정렬)",
              ],
            },
          },
        ],
      },
    ],
  },
  "education-vba": {
    title: "VBA로 하기",
    subtitle: "직접 해보기 · 교육수료자 현황 VBA 미니프로젝트",
    noFold: true,
    sections: [
      {
        heading: "실습 자료",
        callout: {
          icon: "📎",
          paragraphs: [
            '실습워크북 폴더의 "(기본형)교육수료자 현황 실습.xlsx"',
            '실습워크북 폴더의 "교육수료자 현황" 폴더 (1~4월 월별 명단 파일)',
          ],
        },
      },
      {
        heading: "프로젝트 개요",
        callout: {
          icon: "🎯",
          paragraphs: [
            "월별로 쌓이는 교육 이수자 명단 파일을 통합하고, 직급을 직제(급수)로 표준화한 뒤 정렬하고, 마지막으로 직제·직급·소속별 인원수를 요약하는 VBA 미니프로젝트입니다.",
            "아래 4단계는 각각 독립된 매크로(Sub)로 만들어서, 나중에 버튼 4개에 하나씩 연결해 순서대로 눌러 실행할 수 있도록 구성했습니다.",
          ],
        },
      },
      {
        heading: "STEP 1 — 폴더 내 파일 통합",
        group: [
          {
            heading: "미션",
            callout: {
              icon: "🎯",
              paragraphs: [
                '"교육수료자 현황" 폴더 안에 있는 월별 명단 파일들을 하나로 합쳐서, "통합결과_교육수료자 현황"이라는 새 시트에 모아보세요.',
                "각 파일의 '교육수료자 명단' 시트(헤더: 번호·소속·직급·성명)를 읽어서 이어붙이고, 통합이 끝나면 몇 개의 파일과 몇 건의 데이터를 처리했는지 메시지로 알려주세요.",
              ],
            },
          },
          {
            heading: "힌트 — 프롬프트에 이런 내용을 담아보세요",
            callout: {
              icon: "💡",
              steps: [
                "월별 파일이 어느 폴더에 있고, 몇 개나 쌓여 있는지 구체적으로 설명하기",
                "각 파일의 시트 이름, 헤더가 몇 행에 있는지, 데이터가 몇 행부터 시작하는지 알려주기",
                "결과를 모을 새 시트 이름을 정해서 알려주고, 원본 파일은 절대 수정하면 안 된다는 조건 넣기",
                "데이터가 몇 행까지 있을지 모른다는 점을 알려주고, 자동으로 마지막 행까지 인식하도록 요청하기",
                "작업이 끝나면 몇 개 파일에서 몇 건을 처리했는지 알려달라고 요청하기",
                "나중에 버튼에 연결할 것이므로, 독립적으로 실행되는 매크로로 만들어달라고 요청하기",
              ],
            },
          },
          {
            heading: "실행 방법",
            callout: {
              icon: "▶️",
              steps: [
                "Alt + F11로 VBA 편집기 열기",
                "삽입 → 모듈 클릭",
                "완성한 코드 붙여넣기",
                "F5 또는 Ctrl + F5로 실행",
              ],
            },
          },
        ],
      },
      {
        heading: "STEP 2 — 직급을 직제로 변환",
        group: [
          {
            heading: "미션",
            callout: {
              icon: "🎯",
              paragraphs: [
                '"통합결과_교육수료자 현황" 시트의 직급 값을 보고 몇 급인지 판단해서, "직제"라는 새 열에 채워보세요.',
                "직급 값은 \"지방전산서기\", \"지방시설주사\"처럼 앞에 직렬명이 붙고 뒤에 실제 직급명이 오는 형태예요.",
              ],
            },
          },
          {
            heading: "직급 → 직제 매핑표",
            table: {
              headers: ["직급(끝 단어)", "직제"],
              rows: [
                ["관리관", "1급"],
                ["이사관", "2급"],
                ["부이사관", "3급"],
                ["서기관", "4급"],
                ["사무관", "5급"],
                ["주사보", "7급"],
                ["주사", "6급"],
                ["서기보", "9급"],
                ["서기", "8급"],
              ],
            },
          },
          {
            heading: "힌트 — 프롬프트에 이런 내용을 담아보세요",
            callout: {
              icon: "💡",
              steps: [
                "직급 값이 어떤 형태인지 예시를 들어서 설명하기 (직렬명 + 직급명이 붙어있는 구조)",
                "위 매핑표를 그대로 프롬프트에 붙여넣어서, 어떤 기준으로 직제를 판단해야 하는지 알려주기",
                "\"주사보\"와 \"주사\"처럼 한쪽 단어가 다른 단어에 포함되는 경우가 있다는 점을 미리 알려주기",
                "매핑표에 없는 값이 나오면 어떻게 처리할지(예: '미분류'로 표시) 정해서 요청하기",
                "직제를 추가할 위치(어느 열 다음에 넣을지)를 구체적으로 알려주기",
              ],
            },
          },
        ],
      },
      {
        heading: "STEP 3 — 직제 기준 정렬",
        group: [
          {
            heading: "미션",
            callout: {
              icon: "🎯",
              paragraphs: [
                '"통합결과_교육수료자 현황" 시트를 직제(급수) 기준으로 오름차순 정렬하고, 같은 급수 안에서는 소속 기준으로 정렬해보세요.',
                "정렬 후에는 번호 열도 1부터 다시 채워주세요.",
              ],
            },
          },
          {
            heading: "힌트 — 프롬프트에 이런 내용을 담아보세요",
            callout: {
              icon: "💡",
              steps: [
                "정렬 기준이 2단계라는 것을 명확히 알려주기 (먼저 직제, 같은 직제 안에서는 소속)",
                "직제 값이 \"1급\"~\"9급\"처럼 텍스트라서 그대로 정렬하면 순서가 꼬일 수 있다는 점을 알려주고, 숫자 기준으로 정렬해달라고 요청하기",
                "\"미분류\" 값은 맨 뒤로 보내달라고 조건을 넣기",
                "정렬 후 번호 열을 1부터 다시 채워달라고 요청하기",
              ],
            },
          },
        ],
      },
      {
        heading: "STEP 4 — 분석 시트 만들기",
        group: [
          {
            heading: "미션",
            callout: {
              icon: "🎯",
              paragraphs: [
                '"통합결과_교육수료자 현황" 시트를 바탕으로 직제별 / 직급별 / 소속별로 몇 명이 이수했는지 세어서, "분석" 시트에 표 3개로 정리해보세요.',
                "각 표는 인원수가 많은 순서대로 정렬하고, 맨 아래에 총합계 행도 넣어보세요.",
              ],
            },
          },
          {
            heading: "힌트 — 프롬프트에 이런 내용을 담아보세요",
            callout: {
              icon: "💡",
              steps: [
                "어떤 3가지 기준(직제/직급/소속)으로 인원수를 집계해야 하는지 명확히 나열하기",
                "결과를 넣을 새 시트 이름과, 표 3개를 어떻게 배치할지(예: 나란히) 설명하기",
                "각 표는 인원수가 많은 순서로 정렬해달라고 요청하기",
                "각 표 맨 아래에 총합계 행을 추가해달라고 요청하기",
              ],
            },
          },
        ],
      },
      {
        heading: "버튼 연결하기",
        callout: {
          icon: "📌",
          steps: [
            "STEP 1~4에서 만든 4개의 매크로가 모두 모듈에 들어있는지 확인",
            "시트에서 [개발 도구] &gt;&gt; [삽입] &gt;&gt; [양식 컨트롤] &gt;&gt; [단추] 선택 후 원하는 위치에 그리기",
            "매크로 지정 창이 뜨면 STEP 1에서 만든 매크로를 선택",
            "버튼 텍스트를 각 단계 이름으로 수정 (예: \"① 파일 통합\", \"② 직제 변환\", \"③ 정렬\", \"④ 분석표 생성\")",
            "같은 방법으로 나머지 3개 버튼도 STEP 2, 3, 4의 매크로에 각각 연결",
            "버튼을 순서대로 눌러가며 1~4단계가 차례로 실행되는지 확인",
          ],
        },
      },
    ],
  },
  "vba-answers": {
    title: "VBA 미니프로젝트 정답 모음",
    subtitle: "",
    sections: [
      {
        heading: "STEP 1 — MergeMonthlyFiles",
        note: "복사해서 ChatGPT, Claude 등에 붙여넣고 사용하세요",
        prompt:
          "엑셀 VBA 매크로를 작성해줘.\n" +
          "\n" +
          "[목표]\n" +
          "\"교육수료자 현황\" 폴더 안에 있는 여러 개의 월별 교육 이수자 명단 엑셀 파일을 하나로 통합한다.\n" +
          "\n" +
          "[상황]\n" +
          "이 매크로가 들어있는 엑셀 파일과 같은 폴더 안에 \"교육수료자 현황\"이라는 하위 폴더가 있고,\n" +
          "그 안에 1월, 2월, 3월, 4월처럼 매달 쌓이는 동일한 구조의 엑셀 파일들이 있어.\n" +
          "각 파일에는 \"교육수료자 명단\"이라는 시트가 있고, 1행은 제목, 2행은 헤더(번호, 소속, 직급, 성명), 3행부터 데이터가 있어.\n" +
          "\n" +
          "[원하는 동작]\n" +
          "1. 현재 통합문서에 \"통합결과_교육수료자 현황\"이라는 새 시트를 만들어줘 (이미 있으면 내용을 지우고 다시 씀)\n" +
          "2. 헤더로 번호, 소속, 직급, 성명을 입력해줘\n" +
          "3. \"교육수료자 현황\" 폴더 안의 모든 엑셀 파일을 순서대로 열어서 \"교육수료자 명단\" 시트의 데이터를 읽고,\n" +
          "   \"통합결과_교육수료자 현황\" 시트에 이어붙여줘 (원본 파일은 수정하지 말고 읽기 전용으로 열고 닫기)\n" +
          "4. 통합 후 번호 열은 1부터 다시 순서대로 채워줘\n" +
          "5. 통합이 끝나면 몇 개의 파일에서 몇 건의 데이터를 통합했는지 완료 메시지로 알려줘\n" +
          "\n" +
          "[참고]\n" +
          "- 매크로 이름은 MergeMonthlyFiles로 해줘\n" +
          "- 나중에 버튼에 연결해서 독립적으로 실행할 것이므로, 다른 매크로에 의존하지 않는 별도의 Sub로 작성해줘\n" +
          "- 데이터가 몇 행까지 있는지 몰라서 마지막 데이터 행까지 자동으로 인식해서 처리해줘 (행 번호 하드코딩 금지)\n" +
          "- 화면 깜빡임 방지를 위해 Application.ScreenUpdating을 활용해줘",
        promptHighlight: "MergeMonthlyFiles",
        promptHighlightClass: "prompt-highlight-alt",
      },
      {
        heading: "STEP 2 — ConvertRankToGrade",
        note: "복사해서 ChatGPT, Claude 등에 붙여넣고 사용하세요",
        prompt:
          "엑셀 VBA 매크로를 작성해줘.\n" +
          "\n" +
          "[목표]\n" +
          "\"통합결과_교육수료자 현황\" 시트의 직급 값을 직제(급수)로 변환해서 새 열에 추가한다.\n" +
          "\n" +
          "[상황]\n" +
          "직급 열에는 \"지방전산서기\", \"지방시설주사\", \"지방행정주사보\"처럼\n" +
          "앞에 직렬명이 붙고 뒤에 실제 직급명이 붙는 형태로 값이 들어 있어.\n" +
          "직급명은 항상 아래 9가지 중 하나로 끝나.\n" +
          "\n" +
          "[직급 → 직제 매핑 (뒤에서부터 일치하는 것으로 판단)]\n" +
          "관리관 → 1급\n" +
          "이사관 → 2급\n" +
          "부이사관 → 3급\n" +
          "서기관 → 4급\n" +
          "사무관 → 5급\n" +
          "주사보 → 7급\n" +
          "주사 → 6급\n" +
          "서기보 → 9급\n" +
          "서기 → 8급\n" +
          "\n" +
          "[원하는 동작]\n" +
          "1. \"통합결과_교육수료자 현황\" 시트의 B열(소속) 다음에 \"직제\"라는 열을 추가해줘\n" +
          "2. 각 행의 직급 값이 어떤 직급명으로 끝나는지 판별해서 직제(급수)를 입력해줘\n" +
          "   (\"주사보\"와 \"주사\"처럼 한쪽이 다른 쪽 뒤에 포함되는 경우가 있으니, 더 긴 이름부터 먼저 비교해줘)\n" +
          "3. 어떤 직급명과도 일치하지 않으면 \"미분류\"라고 표시해줘\n" +
          "\n" +
          "[참고]\n" +
          "- 매크로 이름은 ConvertRankToGrade로 해줘\n" +
          "- 나중에 버튼에 연결해서 독립적으로 실행할 것이므로, 다른 매크로에 의존하지 않는 별도의 Sub로 작성해줘\n" +
          "- 데이터가 몇 행까지 있는지 몰라서 마지막 데이터 행까지 자동으로 인식해서 처리해줘 (행 번호 하드코딩 금지)\n" +
          "- \"통합결과_교육수료자 현황\" 시트가 없으면 안내 메시지를 띄우고 종료해줘",
        promptHighlight: "ConvertRankToGrade",
        promptHighlightClass: "prompt-highlight-alt",
      },
      {
        heading: "STEP 3 — SortByGrade",
        note: "복사해서 ChatGPT, Claude 등에 붙여넣고 사용하세요",
        prompt:
          "엑셀 VBA 매크로를 작성해줘.\n" +
          "\n" +
          "[목표]\n" +
          "\"통합결과_교육수료자 현황\" 시트의 데이터를 직제(급수) 기준으로 정렬한다.\n" +
          "\n" +
          "[상황]\n" +
          "시트에는 번호, 소속, 직제, 직급, 성명 열이 있고, 직제 열에는 \"1급\"~\"9급\", \"미분류\" 값이 들어 있어.\n" +
          "\n" +
          "[원하는 동작]\n" +
          "1. 직제 열의 숫자(1~9) 기준으로 오름차순 정렬해줘 (\"미분류\"는 맨 아래로)\n" +
          "2. 같은 직제 안에서는 소속 기준으로 오름차순 정렬해줘\n" +
          "3. 정렬 후 번호 열은 1부터 다시 순서대로 채워줘\n" +
          "4. 정렬이 끝나면 완료 메시지를 띄워줘\n" +
          "\n" +
          "[참고]\n" +
          "- 매크로 이름은 SortByGrade로 해줘\n" +
          "- 나중에 버튼에 연결해서 독립적으로 실행할 것이므로, 다른 매크로에 의존하지 않는 별도의 Sub로 작성해줘\n" +
          "- 데이터가 몇 행까지 있는지 몰라서 마지막 데이터 행까지 자동으로 인식해서 처리해줘 (행 번호 하드코딩 금지)\n" +
          "- \"통합결과_교육수료자 현황\" 시트가 없으면 안내 메시지를 띄우고 종료해줘",
        promptHighlight: "SortByGrade",
        promptHighlightClass: "prompt-highlight-alt",
      },
      {
        heading: "STEP 4 — BuildSummarySheet",
        note: "복사해서 ChatGPT, Claude 등에 붙여넣고 사용하세요",
        prompt:
          "엑셀 VBA 매크로를 작성해줘.\n" +
          "\n" +
          "[목표]\n" +
          "\"통합결과_교육수료자 현황\" 시트의 데이터를 바탕으로 직제별, 직급별, 소속별 이수자 인원수를 집계해\n" +
          "\"분석\"이라는 시트에 표로 정리한다.\n" +
          "\n" +
          "[상황]\n" +
          "\"통합결과_교육수료자 현황\" 시트에는 번호, 소속, 직제, 직급, 성명 열이 있어.\n" +
          "\n" +
          "[원하는 동작]\n" +
          "1. \"분석\"이라는 새 시트를 만들어줘 (이미 있으면 내용을 지우고 다시 씀)\n" +
          "2. 왼쪽부터 순서대로 3개의 표를 만들어줘\n" +
          "   - 직제별 이수자 현황 (직제, 인원수)\n" +
          "   - 직급별 이수자 현황 (직급, 인원수)\n" +
          "   - 소속별 이수자 현황 (소속, 인원수)\n" +
          "3. 각 표는 인원수가 많은 순서대로 정렬해줘\n" +
          "4. 각 표 맨 아래에 총합계 행을 추가해줘\n" +
          "5. 집계가 끝나면 완료 메시지를 띄워줘\n" +
          "\n" +
          "[참고]\n" +
          "- 매크로 이름은 BuildSummarySheet로 해줘\n" +
          "- 나중에 버튼에 연결해서 독립적으로 실행할 것이므로, 다른 매크로에 의존하지 않는 별도의 Sub로 작성해줘\n" +
          "- 집계에는 Dictionary(Scripting.Dictionary 또는 Collection)를 활용해줘\n" +
          "- \"통합결과_교육수료자 현황\" 시트가 없으면 안내 메시지를 띄우고 종료해줘",
        promptHighlight: "BuildSummarySheet",
        promptHighlightClass: "prompt-highlight-alt",
      },
    ],
  },
};

const contentEl = document.getElementById("content");
const menuEl = document.getElementById("menu");

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function highlightPrompt(escapedText, highlight, className) {
  if (!highlight) return escapedText;
  const cls = className || "prompt-highlight";
  return escapedText.split(escapeHtml(highlight)).join(
    `<span class="${cls}">${escapeHtml(highlight)}</span>`
  );
}

function renderTable(table) {
  return `
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>${table.headers.map((h) => `<th>${h}</th>`).join("")}</tr>
        </thead>
        <tbody>
          ${table.rows
            .map(
              (row) =>
                `<tr>${row.map((cell) => `<td>${cell}</td>`).join("")}</tr>`
            )
            .join("")}
        </tbody>
      </table>
    </div>
  `;
}

function sectionBody(section) {
  const main = section.steps
    ? `
        <ol class="step-list">
          ${section.steps.map((step) => `<li>${step}</li>`).join("")}
        </ol>
      `
    : section.callout
    ? `
        <div class="callout">
          <span class="callout-icon">${section.callout.icon}</span>
          <div class="callout-body">
            ${
              section.callout.steps
                ? `<ol class="step-list">${section.callout.steps
                    .map((step) => `<li>${step}</li>`)
                    .join("")}</ol>`
                : section.callout.paragraphs
                    .map((p) => `<p>${p}</p>`)
                    .join("")
            }
          </div>
        </div>
      `
    : section.prompt
    ? `
        <div class="prompt-box">
          <div class="prompt-actions">
            <button class="expand-btn" type="button" data-prompt="${escapeHtml(
              section.prompt
            )}" data-title="${escapeHtml(section.heading)}" title="확대해서 보기">⤢</button>
            <button class="copy-btn" type="button" data-prompt="${escapeHtml(
              section.prompt
            )}">복사</button>
          </div>
          <pre class="prompt-text">${highlightPrompt(
            escapeHtml(section.prompt),
            section.promptHighlight,
            section.promptHighlightClass
          )}</pre>
        </div>
      `
    : section.image
    ? `<img class="page-image" src="${section.image}" alt="${section.heading}">`
    : section.cards
    ? renderCards(section.cards)
    : section.paragraphs
    ? `
        <div class="text-block">
          ${section.paragraphs.map((p) => `<p>${p}</p>`).join("")}
        </div>
      `
    : "";

  const table = section.table ? renderTable(section.table) : "";

  return main + table;
}

function renderSubsection(section, skipHeading) {
  const note = section.note
    ? `<p class="section-note">${section.note}</p>`
    : "";
  const heading = skipHeading
    ? ""
    : `<h3 class="subsection-heading">${section.heading}</h3>`;
  return `
    <div class="subsection">
      ${heading}
      ${note}
      ${sectionBody(section)}
    </div>
  `;
}

function renderSections(sections, foldable = true) {
  return sections
    .map((section) => {
      // hidden sections always render as a closed <details>, regardless of
      // the page's fold setting — used to tuck instructor-only content away
      const useFold = section.hidden ? true : foldable;
      const tag = useFold ? "details" : "div";
      const headingTag = useFold ? "summary" : "h2";
      const openAttr = useFold && !section.hidden ? " open" : "";
      const hiddenClass = section.hidden ? " hidden-section" : "";

      if (section.group) {
        const inner = section.group
          .map((s, i) => renderSubsection(s, i === 0))
          .join("");
        return `
      <${tag} class="content-section${hiddenClass}"${openAttr}>
        <${headingTag} class="section-heading">${section.heading}</${headingTag}>
        <div class="content-section-body">
          ${inner}
        </div>
      </${tag}>
    `;
      }

      const body = sectionBody(section);
      const note = section.note
        ? `<p class="section-note">${section.note}</p>`
        : "";
      return `
      <${tag} class="content-section${hiddenClass}"${openAttr}>
        <${headingTag} class="section-heading">${section.heading}</${headingTag}>
        <div class="content-section-body">
          ${note}
          ${body}
        </div>
      </${tag}>
    `;
    })
    .join("");
}

function renderCards(cards) {
  return `
    <div class="card-grid">
      ${cards
        .map(
          (card, i) => `
        <div class="card">
          <div class="card-header">
            <span class="card-num">${i + 1}</span>
            <span class="card-title">${card.title}</span>
          </div>
          <div class="card-icon">${card.icon}</div>
          <div class="card-divider"></div>
          <div class="card-desc">${card.desc}</div>
        </div>
      `
        )
        .join("")}
    </div>
  `;
}

function renderPage(pageId) {
  const page = PAGES[pageId];
  if (!page) return;

  const body = page.sections
    ? renderSections(page.sections, !page.noFold)
    : page.cards
    ? renderCards(page.cards)
    : page.image
    ? `<img class="page-image" src="${page.image}" alt="${page.title}">`
    : `<div class="placeholder-box">준비 중입니다.</div>`;

  const subtitle = page.subtitle
    ? `<p class="page-subtitle">${page.subtitle}</p>`
    : "";

  contentEl.innerHTML = `
    <h1 class="page-title">${page.title}</h1>
    ${subtitle}
    ${body}
  `;

  menuEl.querySelectorAll(".menu-item").forEach((el) => {
    el.classList.toggle("active", el.dataset.page === pageId);
  });
}

menuEl.addEventListener("click", (e) => {
  const btn = e.target.closest(".menu-item");
  if (!btn) return;
  renderPage(btn.dataset.page);
});

const sidebarSecret = document.querySelector(".sidebar-secret");
if (sidebarSecret) {
  sidebarSecret.addEventListener("click", () => {
    renderPage(sidebarSecret.dataset.page);
  });
}

contentEl.addEventListener("click", (e) => {
  const copyBtn = e.target.closest(".copy-btn");
  if (copyBtn) {
    navigator.clipboard.writeText(copyBtn.dataset.prompt).then(() => {
      const original = copyBtn.textContent;
      copyBtn.textContent = "복사됨";
      copyBtn.classList.add("copied");
      setTimeout(() => {
        copyBtn.textContent = original;
        copyBtn.classList.remove("copied");
      }, 1500);
    });
    return;
  }

  const expandBtn = e.target.closest(".expand-btn");
  if (expandBtn) {
    openPromptModal(expandBtn.dataset.title, expandBtn.dataset.prompt);
  }
});

const promptModal = document.getElementById("promptModal");
const modalTitle = document.getElementById("modalTitle");
const modalText = document.getElementById("modalText");
const modalClose = document.getElementById("modalClose");

function openPromptModal(title, text) {
  modalTitle.textContent = title;
  modalText.textContent = text;
  promptModal.classList.add("open");
}

function closePromptModal() {
  promptModal.classList.remove("open");
}

modalClose.addEventListener("click", closePromptModal);
promptModal.addEventListener("click", (e) => {
  if (e.target === promptModal) closePromptModal();
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") closePromptModal();
});

renderPage("powerquery-overview");

const sidebarEl = document.getElementById("sidebar");
const sidebarResizer = document.getElementById("sidebarResizer");
const SIDEBAR_MIN = 160;
const SIDEBAR_MAX = 480;

sidebarResizer.addEventListener("mousedown", (e) => {
  e.preventDefault();
  sidebarEl.classList.add("resizing");

  function onMouseMove(moveEvent) {
    const newWidth = moveEvent.clientX - sidebarEl.getBoundingClientRect().left;
    const clamped = Math.min(SIDEBAR_MAX, Math.max(SIDEBAR_MIN, newWidth));
    sidebarEl.style.width = clamped + "px";
  }

  function onMouseUp() {
    sidebarEl.classList.remove("resizing");
    document.removeEventListener("mousemove", onMouseMove);
    document.removeEventListener("mouseup", onMouseUp);
  }

  document.addEventListener("mousemove", onMouseMove);
  document.addEventListener("mouseup", onMouseUp);
});
