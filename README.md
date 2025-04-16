# カレンダー登録機能

紙で入手した年間計画を、OCRで電子化し、手動で一覧として整え、カレンダーAPIを使用して自動登録します。

一覧をメンテして登録機能を起動するとメンテ内容が反映されます。

## 読み取りから登録までの手順

### OCR

YomiToku を利用します。PDFまたは画像を入力としてhtmlの表形式のデータを作成します。

https://github.com/kotaro-kinoshita/yomitoku

画像またはPDFを```./in```に格納し、以下のコマンドでOCRで電子化したhtmlを作成してください。結果は```./results```に格納されます。

```
yomitoku in -f html -o results -v --figure
```

読み取り精度は入力とした画像ファイルの品質によります。完璧な予定表は再現されないので、次のステップで整形してください。

##### yomitokuの説明

```
yomitoku ${path_data} -f html -o ${outdir} -v --figure --lite
```

- ${path_data} 解析対象の画像が含まれたディレクトリか画像ファイルのパスを直接して指定してください。ディレクトリを対象とした場合はディレクトリのサブディレクトリ内の画像も含めて処理を実行します。
- -f, --format 出力形式のファイルフォーマットを指定します。(json, csv, html, md をサポート)
- -o, --outdir 出力先のディレクトリ名を指定します。存在しない場合は新規で作成されます。
- -v, --vis を指定すると解析結果を可視化した画像を出力します。
- -l, --lite を指定すると軽量モデルで推論を実行します。通常より高速に推論できますが、若干、精度が低下する可能性があります。



※ YomiToku は日本語に特化した AI 文章画像解析エンジン(Document AI)です。画像内の文字の全文 OCR およびレイアウト解析機能を有しており、画像内の文字情報や図表を認識、抽出、変換します。

### 一覧作成

OCRで読み取ったhtmlをブラウザで表示し、表計算ソフトに貼り付け(コピー&ペースト)、電子データとして並べ、読み取り誤りを修正したり、表形式を整えてください。修正したら'Excel2007-365形式'で保存してください。

ここはゴリゴリの手作業です。なんとか頑張って以下のようなフォーマットの表を作成してください。(OCRで電子データがあるのでゼロから手入力よりはマシでしょ？という考え方)

※ sample.xlsx をサンプルとして参考にしてください。


| 開始日 | 終了日 | 行事 | 内容 | 作成日 | 更新日 | イベントID |
| ------ | ------ | ---- | ---- | ------ | ------ | ---------- |
|        |        |      |      |        |        |            |

- 開始日
  イベントの開始日。時間を記入する必要はない。同じ日のイベントでも別の行に記載すること。1イベント1行です。
- 終了日
  イベントの終了日。この日を含む予定として作成される。一日のみのイベントの場合は記入不要。
- 行事
  イベントの名称
- 内容
  イベントの内容。補足事項など
- 作成日
  システムで管理する値。記入不要。カレンダーイベントの作成日
- 更新日
  システムで管理する値。記入不要。カレンダーイベントの更新日
- イベントID
  システムで管理するID。記入不要。プログラムを実行して登録処理を行うとここにIDが記入される。この値は変更してはいけない。

> コマンドでGoogleカレンダーをダウンロードすると同じ形式のファイルが作成されるので雛形として使用できる。
> 新年度、イベント登録がない状態でダウンロードすると空の雛形ファイルが作成される。



### カレンダー登録

一覧をもとにカレンダー登録機能でGoogleカレンダーに一気に大量に登録します。

登録時にイベントIDが生成されるため、一覧を修正して登録しなおせばカレンダー側の該当イベントも修正されます。

まずは以下の記事を参考にGoogleカレンダーへアクセスできるように準備してください。

https://developers.google.com/calendar/api/quickstart/python?hl=ja

上記のリンク先で以下のような作業を行っています。

- APIを有効にする
- OAuth同意画面を構成する
- デスクトップアプリケーションの認証情報を承認する
  - credentials.jsonを作成します
- Googleクライアントライブラリをインストールする
  - [実行環境を準備する手順](#実行環境を準備する手順)でインストールされます


カレンダーへのアクセスの準備と一覧の作成ができたら以下のコマンドで登録処理を行ってください。

```
(.venv) $ python hinocal.py upload -sy 2025

※ -sy: school year. 対象とする年度を指定する
```

初回実行時にはGoogleの認証処理が行われます。ブラウザを操作して適宜認証してください。


##### hinocal.pyの説明

Googleカレンダーを正として、hinocalコマンドでデータをダウンロードし、ダウンロードしてできたExcelファイルを編集し、編集内容をアップロードしてカレンダーに反映する。

1. ダウンロード
  ```python hinocal.py download -sy 2025```

2. ファイル編集
   ```calendar_sy2025.xlsx```

3. アップロード
  ```python hinocal.py upload -sy 2025```


```

(.venv) $ python hinocal.py -h
usage: hinocal.py [-h] [-re] [-sd STARTDATE] [-f FILE] [-cf CALENDAR_FILE] [-sy SCHOOL_YEAR]
                  {list,sync,calendar,download,upload}

Googleカレンダーにイベント(予定)を登録・更新する。Googleカレンダーの情報を正と捉え、更新するためにダウンロードし、編集後にアップロードしてカレンダーを更新する。

positional arguments:
  {list,sync,calendar,download,upload}
                        list: Get and print events from Google calender. sync: Sync local to Google. calendar: Get and print
                        calenders from Google.

options:
  -h, --help            show this help message and exit
  -re, --relogin        サインイン情報をクリアしてから実行する
  -sd STARTDATE, --startdate STARTDATE
                        開始年月 yyyy-mm
  -f FILE, --file FILE  行事予定一覧excelファイル
  -cf CALENDAR_FILE, --calendar_file CALENDAR_FILE
                        カレンダーの内容を書き出す/読み込むexcelファイル名。指定しない場合は'calendar_syYYYY.xlsx'。既存のファイルは上書きされる。
  -sy SCHOOL_YEAR, --school_year SCHOOL_YEAR
                        年度 yyyy。downloadとuploadの際に使用する。


```

## 実行環境を準備する手順

※ まずはこちらからですね。環境を整えてください。

1. リポジトリのクローン

```
git clone https://github.com/ECR33/hinocal.git
```

2. python仮想環境作成

```
cd hinocal
python3 -m venv .venv
source .venv/bin/activate
```

3. ライブラリインストール

```
pip install -r requirements.txt
```

## 目的

3年間手作業でカレンダーをメンテしてみて、以下を改善する必要があると感じたため、誤りの少なく、かつ、なるべく自動化できる環境を作成しました。

- 手作業での登録が面倒・タイポしやすい。
  - カレンダーに予定として登録して、紙とのチェック、とオペレーションが煩雑。
  - どれ(どの予定を)見てたかわからなくなる。
- メンテ作業が面倒
  - 月単位で更新版のデータ(紙)をもらいますが、当初の予定に記載されていた予定が見当たらないことがありました。これが、予定を削除したのか、別の月へ変更されたのかとてもわかりにくかったため、一覧性の高いメンテ方法が必要となった。Googleカレンダーでも一覧表示できるが、一覧を表示したままひとつひとつの予定を修正できないため不便

なんだかんだ言って、エクセルが一般の利用者にいちばんなじみがあり、ミスが少ないインタフェースと考えられるため、メンテ用UIとして採用。

エクセルのまま登録までできるとユーザにとって優しいが、エクセルを持っていないので開発不可。Google spreadsheetはカレンダーとの親和性が高いかもしれないが、ローカルでの開発がとても面倒なので採用しなかった。

pythonなら逸般人は誰でも使えるので採用。(期待してます。)
