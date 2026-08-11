# PR #13 承認済み修正反映の安全テスト

## 目的

⑧「承認済み修正候補を反映」を本番データで試す前に、まず同じ中核処理を一時シートだけで検証し、その後コピー版スプレッドシートで⑧そのものをUI経由で検証します。

通常の鳥取データでは⑧を実行しません。

## 1. 中核処理の隔離テスト

`runHotelDbV2ApprovedCorrectionSafetyTests` は一時的に3枚のテストシートを作成し、次を確認します。

1. `承認` の候補だけが処理対象になる
2. 正常な承認候補1件だけが元データへ反映される
3. `未確認` の候補は反映されない
4. 候補作成後に元データが変わった競合ケースは反映されない
5. 競合候補は `要再確認` になる
6. 正常反映候補は `反映済み` になる
7. 正常反映候補に反映日時が入る
8. 正常反映の履歴が残る
9. 競合・未反映の履歴が残る
10. 本物の `修正候補` シートの行数が変わらない
11. 本物の `修正履歴` シートの行数が変わらない
12. 一時テストシートはテスト終了後に削除される

内部では合計16項目をチェックします。

## 2. ⑧そのもののUIテスト

コピー版スプレッドシート名に `PR13` と `⑧反映テスト` の両方が含まれている場合だけ実行できます。

新規ファイル `V2ApprovedCorrectionUiTest.gs` の3関数を使います。

### セットアップ

`setupHotelDbV2ApprovedCorrectionUiTest`

- 通常の `修正候補` と `修正履歴` をバックアップ名へ退避
- `PR13_UI反映元` を作成
- ⑧が使う標準名 `修正候補` と `修正履歴` にはテスト専用データだけを配置
- 承認2件（正常1件・競合1件）、未確認1件を用意

セットアップ完了後に、スプレッドシートのメニューから⑧「承認済み修正候補を反映」を実行します。

期待結果は次です。

```text
承認対象: 2
反映済み: 1
要再確認: 1
エラー: 0
```

### 検証

⑧実行後に `verifyHotelDbV2ApprovedCorrectionUiTest` を実行します。

12項目を確認し、正常なら次を表示します。

```text
PR #13 ⑧UIテスト 成功

成功件数: 12件
失敗件数: 0件
正常承認: 1件だけ反映
未承認: 未反映
競合: 要再確認・未反映
履歴: 2件記録
```

### 復元

検証後に `cleanupHotelDbV2ApprovedCorrectionUiTest` を実行します。

- UIテスト専用シートを削除
- 退避した `修正候補` と `修正履歴` を元の標準名へ復元

## 本番コードの変更

`hotelDbV2ApplyApprovedCorrections_()` の外部動作は変更しません。

本番関数は、共通中核処理 `hotelDbV2ApplyApprovedCorrectionsWithContext_({})` を呼ぶだけです。

隔離テスト時だけ、共通中核処理へ一時の修正候補シート・履歴シートを明示的に渡します。

## Apps Scriptでの手順

1. PR #13の `V2Operations.gs` でApps Scriptの同名ファイルを置き換える
2. `V2ApprovedCorrectionSafetyTests.gs` を追加する
3. `runHotelDbV2ApprovedCorrectionSafetyTests` を実行し、16件成功を確認する
4. スプレッドシートをコピーし、コピー名に `PR13_⑧反映テスト` を含める
5. コピー版のApps Scriptへ `V2ApprovedCorrectionUiTest.gs` を追加する
6. `setupHotelDbV2ApprovedCorrectionUiTest` を実行する
7. コピー版だけでメニュー⑧を実行する
8. `verifyHotelDbV2ApprovedCorrectionUiTest` を実行する
9. `cleanupHotelDbV2ApprovedCorrectionUiTest` を実行する
10. 問題がなければPR #13をマージする

## 注意

通常のテストDBや鳥取本体では、PR #13検証中に⑧を実行しません。
