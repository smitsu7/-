# r11 VBAをブラウザからコピペで移す手順

## 用意したもの

`paste`の24個のテキストが、ExcelのVBA画面へ貼り付けるコードです。標準モジュール23個と、ThisWorkbookのイベントコード1個があります。現行r11の取り込み用コードから、VBA画面には貼れない書き出し情報だけを除きました。計算処理は変更していません。

`readable`には日本語のコメントを残した正本を置きました。処理の確認にはこちらを使います。貼り付けには、Windowsの言語による文字化けを避けるため、`paste`のASCII版を使います。ASCII版でも実行時の日本語表示は残ります。

このリポジトリでは、Excelへ貼り付けるためのテキストと導入手順を提供します。原指数、計算結果、Excelブック、X-13実行ファイルは含みません。最初に1～3章を順に読み、貼り付けと初期シートの作成を済ませてください。計算を始める前に、4章のファイル配置を確認します。

## 1. 社用PCで新しいブックを作る

1. 新しいExcelブックを作り、`IIP_X13_Tool_r11.xlsm`として保存します。既存の業務ブックには貼り付けません。
2. `Alt+F11`でVBAの画面を開きます。左側で、いま保存したブックのVBAProjectを選びます。
3. 「挿入」→「標準モジュール」を選びます。
4. `F4`でプロパティを開き、モジュールの「(オブジェクト名)」または「(Name)」を、下の表の貼り付け先と同じ名前にします。
5. 対応する`paste`のテキストをブラウザで開き、コードだけを全文コピーします。GitHubならファイルの「Raw」表示か、ファイル内容をコピーするボタンを使います。行番号やページの見出しは含めません。
6. VBAのコード欄で既存の空コードを`Ctrl+A`で選び、貼り付けます。自動で入っている`Option Explicit`も置き換え、重複させません。
7. 01～23について、別々の標準モジュールを作って同じ操作を行います。全部を一つのモジュールへまとめないでください。モジュール名を使った呼出しがあるため、名前をそろえる必要があります。

## 2. ThisWorkbookへ貼り付ける

24だけは標準モジュールではありません。左側の「Microsoft Excel Objects」内の`ThisWorkbook`を開き、24のコードを貼り付けます。新しいクラスモジュールは作りません。

## 3. シートとボタンを作る

1. 全24個の貼り付けが終わってから、「デバッグ」→「VBAProjectのコンパイル」を実行します。
2. Excelへ戻り、`Alt+F8`から`modSetup.BuildWorkbook`を実行します。表示が`BuildWorkbook`だけの場合は、いま作ったブックのものを選びます。
3. ホームの「季節調整」やControl、設定・結果のシートが作られます。この初期化だけではX-13の実行やネット取得は始まりません。
4. 保存して閉じ、開き直してホームとボタンが残ることを確かめます。

## 4. 季節調整に必要なファイル

コードを貼ることで再現できるのはExcel側です。計算にはWindows版X-13と祝日CSVも必要です。社用PCで取得できる場合は、以下の公式配布を使います。会社側ですでに用意されているX-13がある場合は、Controlでその場所を指定できます。

- X-13 ASCII版 Version 1.1 Build 62：<https://www2.census.gov/software/x-13arima-seats/x13as/windows/program-archives/x13as_ascii-v1-1-b62.zip>
- 内閣府の祝日CSV：<https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv>

ブックと同じフォルダを基準に、次のように配置します。原指数は利用者が用意します。

```text
IIP_X13_Package/
  IIP_X13_Tool_r11.xlsm
  01_engine/
    x13as_ascii.exe
  02_input/
    public_excel/        ← 原指数Excel
    series/              ← 原値CSVの保存先
  03_holiday/
    source/
      cabinet_holidays.csv
    generated/           ← 祝日変数の保存先
  04_results/
    work/
    candidates/
    final/
    summaries/
```

内閣府のCSVを`cabinet_holidays.csv`の名前で置き、ホームの「祝日変数の作成と確認」で変数を作ります。実績は連続144か月、初期値は2014年1月～2025年12月、予測は別に12か月です。設定期間と予測期間をカバーする祝日CSVが必要です。

X-13の作業フォルダには半角英数字だけのパスを指定します。ブックを日本語を含む場所に置いた場合は、ControlのX13_WORK_DIRに`C:\IIP_X13_work`などを指定します。

## 貼り付け先一覧

| 番号 | 貼り付け先の名前 | コピペ用コード | 行数 |
|---|---|---|---:|
| 01 | `modMain` | [01_modMain.txt](paste/01_modMain.txt) | 190 |
| 02 | `modSheetNames` | [02_modSheetNames.txt](paste/02_modSheetNames.txt) | 145 |
| 03 | `modX13Options` | [03_modX13Options.txt](paste/03_modX13Options.txt) | 362 |
| 04 | `modPackagePaths` | [04_modPackagePaths.txt](paste/04_modPackagePaths.txt) | 147 |
| 05 | `modLog` | [05_modLog.txt](paste/05_modLog.txt) | 59 |
| 06 | `modHash` | [06_modHash.txt](paste/06_modHash.txt) | 189 |
| 07 | `modIIPTextTable` | [07_modIIPTextTable.txt](paste/07_modIIPTextTable.txt) | 236 |
| 08 | `modPublicImport` | [08_modPublicImport.txt](paste/08_modPublicImport.txt) | 1195 |
| 09 | `modHolidayIIP` | [09_modHolidayIIP.txt](paste/09_modHolidayIIP.txt) | 559 |
| 10 | `modZeroReplace` | [10_modZeroReplace.txt](paste/10_modZeroReplace.txt) | 157 |
| 11 | `modSeriesSettings` | [11_modSeriesSettings.txt](paste/11_modSeriesSettings.txt) | 305 |
| 12 | `modX13Interventions` | [12_modX13Interventions.txt](paste/12_modX13Interventions.txt) | 150 |
| 13 | `modX13Calendar` | [13_modX13Calendar.txt](paste/13_modX13Calendar.txt) | 177 |
| 14 | `modX13Spec` | [14_modX13Spec.txt](paste/14_modX13Spec.txt) | 191 |
| 15 | `modX13Native` | [15_modX13Native.txt](paste/15_modX13Native.txt) | 228 |
| 16 | `modX13IO` | [16_modX13IO.txt](paste/16_modX13IO.txt) | 503 |
| 17 | `modX13Results` | [17_modX13Results.txt](paste/17_modX13Results.txt) | 464 |
| 18 | `modX13Store` | [18_modX13Store.txt](paste/18_modX13Store.txt) | 589 |
| 19 | `modX13Run` | [19_modX13Run.txt](paste/19_modX13Run.txt) | 623 |
| 20 | `modX13Provisional` | [20_modX13Provisional.txt](paste/20_modX13Provisional.txt) | 83 |
| 21 | `modX13Summary` | [21_modX13Summary.txt](paste/21_modX13Summary.txt) | 713 |
| 22 | `modUI` | [22_modUI.txt](paste/22_modUI.txt) | 111 |
| 23 | `modSetup` | [23_modSetup.txt](paste/23_modSetup.txt) | 169 |
| 24 | `ThisWorkbook` | [24_ThisWorkbook.txt](paste/24_ThisWorkbook.txt) | 13 |

## 確認した範囲

24個すべてで、r11の取り込み用コードから書き出し用ヘッダーだけを除いたことを照合しました。標準モジュール名は現行版と一致し、ThisWorkbookの2個のイベントも保持しています。

社用PCでのコピペ、コンパイル、シート作成、X-13実行はまだ確認していません。現行r11のWindows受入記録と、今回の貼り付け用テキストの照合を区別しています。

## 操作の参考

- GitHub公式：ファイル内容の表示・コピー <https://docs.github.com/en/repositories/working-with-files/using-files/viewing-and-understanding-files>
- Microsoft公式：VBAの挿入メニュー <https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/insert-menu>
- Microsoft公式：VBAモジュールのコピー <https://support.microsoft.com/en-us/excel/copy-a-macro-module-to-another-workbook>

## 用語とエラー時の確認

- 標準モジュール：VBAの処理を置くコード欄です。このツールでは23個を名前で区別します。
- ThisWorkbook：このブックを開くときなどの処理を置く、Excelに最初からあるコード欄です。
- Control：ファイルの場所や計算条件を設定するシートです。
- 原指数：季節調整前の指数です。利用者が公表原指数のExcelを用意します。

貼り付け後にコンパイルエラーが出た場合は、モジュール名、コードの先頭と末尾、ThisWorkbookの貼り付け先を確認します。
解決しない場合は、エラーメッセージと該当するモジュール名を控えてください。相談の際に原指数や社内パスを公開する必要はありません。
