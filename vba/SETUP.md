# VBAセットアップ手順

## 準備
1. `○在庫帳○27期_テスト用少データ.xlsx` を **名前を付けて保存** → `.xlsm` 形式で保存
2. `Alt + F11` でVBAエディターを開く

---

## モジュール追加

### mMain（標準モジュール）
1. メニュー「挿入」→「標準モジュール」
2. モジュール名を `mMain` に変更（プロパティウィンドウ）
3. `mMain.bas` の内容を貼り付け

---

## フォーム作成

### frmSettings（設定フォーム）
1. メニュー「挿入」→「ユーザーフォーム」
2. フォーム名を `frmSettings` に変更
3. 以下のコントロールを配置：

| コントロール種類 | Name | Caption/Text | 位置の目安 |
|---|---|---|---|
| Label | lblSheetName | 対象シート名 | 上部 |
| TextBox | txtSheetName | （空） | lblSheetNameの右 |
| Label | lblFolder | 画像フォルダ | 中段 |
| TextBox | txtFolder | （空） | lblFolderの右、幅広め |
| CommandButton | btnBrowse | 参照... | txtFolderの右 |
| CommandButton | btnOK | OK | 下部 |
| CommandButton | btnCancel | キャンセル | 下部 |

4. `frmSettings.bas` の内容をフォームのコードウィンドウに貼り付け

---

### frmInput（入力フォーム）
1. メニュー「挿入」→「ユーザーフォーム」
2. フォーム名を `frmInput` に変更
3. フォームを横長に（幅 720pt × 高さ 520pt 程度）
4. 以下のコントロールを配置：

**左ペイン（プレビューエリア）**

| コントロール種類 | Name | Caption | 備考 |
|---|---|---|---|
| Label | lblProgress | 1 / 5 | 右上 |
| Image | imgPreview | | 大きめ（幅280×高360pt）、PictureSizeMode=3 |
| Label | lblFileName | （空） | imgPreviewの下 |

**右ペイン（入力エリア）15項目**

| Label Name | TextBox Name | Caption |
|---|---|---|
| lblCarName | txtCarName | 車名 |
| lblGrade | txtGrade | グレード |
| lblYear | txtYear | 年式 |
| lblMonth | txtMonth | 月 |
| lblColor | txtColor | 色 |
| lblChassis | txtChassis | 車台番号 |
| lblScore | txtScore | 評価点 |
| lblMileage | txtMileage | 走行距離(km) |
| lblPrice | txtPrice | 落札価格(円) |
| lblTax | txtTax | 消費税(円) |
| lblCarTax | txtCarTax | 自動車税(円) |
| lblRecycle | txtRecycle | リサイクル料(円) |
| lblFee | txtFee | 落札手数料(円) |
| lblVenue | txtVenue | オークション会場 |
| lblLotNum | txtLotNum | 出品番号 |

**ボタン（フォーム下部）**

| Name | Caption | Default |
|---|---|---|
| btnImport | 取込 | True |
| btnCancel | 中止 | False |

5. `frmInput.bas` の内容をフォームのコードウィンドウに貼り付け

---

## 実行方法
- `Alt + F8` → `Main` を選択して実行
- またはシートにボタンを配置して `Main` をマクロとして割り当て

---

## テスト用画像フォルダ
任意のフォルダにJPG/PNG/BMPファイルを数枚置いてテストしてください。
処理後は同フォルダ内に `処理済み/` フォルダが自動作成され、画像が移動します。
