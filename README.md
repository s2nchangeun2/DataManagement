# DataManagement
Excel 시트를 JSON(또는 XML)으로 변환하고, Unity에서 클래스 구조로 게임 데이터를 관리하는 툴입니다.  

![converter2](https://github.com/s2nchangeun2/DataManagement/assets/132358969/44618f11-5482-4957-846c-3dd9bb2f7d38)

### 주요 기능
### 1. Excel -> JSON/XML 변환(Editor전용)
- 메뉴 : Tools > ExcelConverter  
- 엑셀 첫 번째 시트의 첫 행을 필드명(ID, RANK 등)으로 사용, 나머지 행을 데이터로 파싱  
- 셀 값에 콤마가 있으면, int[] 배열로 변환  
- 출력 포맷 : JSON 또는 XML 선택 가능  
- 텍스트 변환 토글 체크 시 값을 문자열로, 미체크 시 int로 저장  
- 결과는 지정한 폴더에 '파일명.json(또는 .xml)'로 저장  

### 2. 런타임 데이터 관리
- `Data<T, TT>` 제네릭 베이스 클래스가 `Assets/Data/Json/{타입이름}.Json`을 읽어 리스트로 역직렬화
- `DataManager` 싱글턴이 `Dictionary<Type, Data>`로 모든 데이터 클래스를 등록/조회
- 데이터 클래스별로 `SetData_Inner()`를 오버라이드해 원하는 자료구조(예: ID 기준 Dictionary)로 가공 가능
