# PR #17 承認済み新規追加候補の安全追加

- `状態=承認`だけ処理する。
- 追加直前にGoogle Placesを再取得し、営業中・自治体・名称・住所・郵便番号を再確認する。
- 既存DBのPlace ID・名称・住所を再照合し、競合は`要再確認`で未追加にする。
- 成功時だけ元シートへ新規1行を追加し、既存行は変更・削除しない。
- `宿泊分類`は自動入力せず、備考に要確認を残す。
- 20件の安全判定自己診断は成功済み（20件成功・0件失敗）。

## コピー版UI実地テスト

`V2ApprovedNewFacilityInsertUiTest.gs`をコピー版Apps Scriptへ追加し、次の順で実行する。

1. `setupHotelDbV2ApprovedNewFacilityUiTest`
2. メニュー⑭ `承認済み新規追加候補を安全に追加`
3. `verifyHotelDbV2ApprovedNewFacilityUiTest`
4. `cleanupHotelDbV2ApprovedNewFacilityUiTest`

セットアップは通常の`新規追加候補`と`修正履歴`をバックアップ名へ退避し、テスト専用データだけを標準名へ置く。テストでは、正常承認1件だけが追加され、未承認・既存Place ID重複・候補名変更は未追加になることを確認する。最後にcleanupで通常データを復元する。

現在は自己診断成功済みで、コピー版UI実地テスト待ち。両方が成功してからマージ判断する。
