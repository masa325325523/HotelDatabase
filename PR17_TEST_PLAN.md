# PR #17 validation plan

1. Apps Scriptへ`V2ApprovedNewFacilityInsert.gs`と`V2ApprovedNewFacilityInsertTests.gs`を追加する。
2. `Code.gs`をPR #17版へ更新する。
3. `runHotelDbV2ApprovedNewFacilityInsertTests`を実行し、20件成功・0件失敗を確認する。
4. 自己診断成功後、コピー版専用の実地テストを追加する。
5. 正常承認だけが1行追加され、未承認・重複・候補変更は未追加になることを確認する。
6. 宿泊分類が自動入力されないこと、備考・Google情報・履歴・監査列を確認する。
7. 実地テスト成功後にのみマージ判断する。
