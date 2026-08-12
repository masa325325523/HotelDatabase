/**
 * PR #14 低スコア要確認の追加安全判定。
 * 元の市区町村と候補住所の自治体が明確に異なる場合だけ、
 * 「要人確認」から「誤候補有力」へ再仕分けする。
 * 元データ・Place ID・状態列は変更しない。
 */

function runHotelDbV2LowScoreReviewTriageWithMunicipalityRefinement() {
  return withHotelDbV2Lock_('低スコア要確認自動仕分け', function() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '低スコア要確認を自動仕分け',
      '「要確認」の低スコア行を、安全側に仕分けします。\n\n' +
      '本体の自動採用基準75点は変更しません。\n' +
      '状態・元データ・Place IDも変更しません。\n' +
      '明確な自治体違いは「誤候補有力」として補助判定します。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2TriageLowScoreReviews_();
    const refined = hotelDbV2RefineLowScoreMunicipalityConflicts_();

    result.wrongLikely += refined.changed;
    result.humanReview = Math.max(0, result.humanReview - refined.changed);

    ui.alert([
      '低スコア要確認の自動仕分け完了',
      '',
      '対象件数: ' + result.total,
      '同一施設有力: ' + result.sameLikely,
      '誤候補有力: ' + result.wrongLikely,
      '要人確認: ' + result.humanReview,
      '対象外: ' + result.outOfScope,
      '自治体違いで再仕分け: ' + refined.changed,
      '',
      '状態・元データ・Place IDは変更していません。'
    ].join('\n'));

    return result;
  });
}

function hotelDbV2RefineLowScoreMunicipalityConflicts_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.REVIEW);
  if (!sheet || sheet.getLastRow() < 2) return { changed: 0 };

  const map = hotelDbV2LowScoreHeaderMap_(sheet);
  hotelDbV2LowScoreValidateHeaders_(map);

  const rowCount = sheet.getLastRow() - 1;
  const values = sheet.getRange(2, 1, rowCount, sheet.getLastColumn()).getDisplayValues();
  const recommendations = [];
  const reasons = [];
  const confidences = [];
  let changed = 0;

  values.forEach(function(row) {
    const currentRecommendation = hotelDbV2Clean_(row[map['推奨判定'] - 1]);
    let recommendation = currentRecommendation;
    let reason = hotelDbV2Clean_(row[map['自動判定理由'] - 1]);
    let confidence = Number(hotelDbV2Clean_(row[map['信頼度'] - 1]) || 0);

    if (currentRecommendation === '要人確認') {
      const municipality = hotelDbV2Clean_(row[map['市区町村'] - 1]);
      const candidateAddress = hotelDbV2Clean_(row[map['候補住所'] - 1]);
      const candidatePlaceId = hotelDbV2Clean_(row[map['候補Place ID'] - 1]);
      const businessStatus = hotelDbV2Clean_(row[map['営業状態'] - 1]);
      const reviewReason = hotelDbV2Clean_(row[map['理由'] - 1]);

      if (
        candidatePlaceId &&
        businessStatus === '営業中' &&
        HOTEL_DB_V2_LOW_SCORE_REASONS.indexOf(reviewReason) !== -1 &&
        hotelDbV2LowScoreHasClearMunicipalityConflict_(municipality, candidateAddress)
      ) {
        recommendation = '誤候補有力';
        reason = '候補住所の自治体が元の市区町村と明確に異なるため、同名・系列名が似ていても別施設候補の可能性が高いです。自動削除・自動反映はしません。';
        confidence = 98;
        changed++;
      }
    }

    recommendations.push([recommendation]);
    reasons.push([reason]);
    confidences.push([confidence]);
  });

  sheet.getRange(2, map['推奨判定'], rowCount, 1).setValues(recommendations);
  sheet.getRange(2, map['自動判定理由'], rowCount, 1).setValues(reasons);
  sheet.getRange(2, map['信頼度'], rowCount, 1).setValues(confidences);

  return { changed: changed };
}

function hotelDbV2LowScoreHasClearMunicipalityConflict_(municipality, candidateAddress) {
  const source = hotelDbV2LowScoreExtractMunicipalityParts_(municipality);
  const candidate = hotelDbV2LowScoreExtractMunicipalityParts_(candidateAddress);

  if (source.city && candidate.city && source.city !== candidate.city) return true;

  if (
    source.city && candidate.city && source.city === candidate.city &&
    source.ward && candidate.ward && source.ward !== candidate.ward
  ) return true;

  if (!source.city && !candidate.city && source.ward && candidate.ward && source.ward !== candidate.ward) {
    return true;
  }

  if (!source.city && !source.ward && source.town && candidate.town && source.town !== candidate.town) {
    return true;
  }

  if (!source.city && !source.ward && source.village && candidate.village && source.village !== candidate.village) {
    return true;
  }

  return false;
}

function hotelDbV2LowScoreExtractMunicipalityParts_(value) {
  let text = hotelDbV2Clean_(value).normalize('NFKC').replace(/\s+/g, '');
  text = text.replace(/^(?:北海道|東京都|大阪府|京都府|.{2,3}県)/, '');

  const cityMatch = text.match(/^(.{1,12}?市)/);
  const city = cityMatch ? cityMatch[1] : '';
  const afterCity = city ? text.slice(city.length) : text;

  const wardMatch = afterCity.match(/^(.{1,12}?区)/);
  const ward = wardMatch ? wardMatch[1] : '';

  let town = '';
  let village = '';
  if (!city && !ward) {
    let local = text.replace(/^.{1,12}?郡/, '');
    const townMatch = local.match(/^(.{1,12}?町)/);
    const villageMatch = local.match(/^(.{1,12}?村)/);
    town = townMatch ? townMatch[1] : '';
    village = villageMatch ? villageMatch[1] : '';
  }

  return {
    city: city,
    ward: ward,
    town: town,
    village: village
  };
}

function runHotelDbV2LowScoreMunicipalityRefinementTests() {
  const cases = [
    ['大阪市北区', '兵庫県神戸市中央区北野町1丁目', true, '大阪→神戸'],
    ['大阪市北区', '大阪府大阪市北区芝田2丁目', false, '同じ大阪市北区'],
    ['大阪市北区', '大阪府大阪市中央区心斎橋筋1丁目', true, '同じ市の別区'],
    ['東京都新宿区', '東京都渋谷区道玄坂1丁目', true, '東京23区の別区'],
    ['鳥取県岩美郡', '鳥取県岩美郡岩美町岩井536', false, '郡のみの元表記は断定しない'],
    ['岩美町', '鳥取県八頭郡八頭町郡家648', true, '町が明確に異なる']
  ];

  const failures = [];
  cases.forEach(function(testCase, index) {
    const actual = hotelDbV2LowScoreHasClearMunicipalityConflict_(testCase[0], testCase[1]);
    if (actual !== testCase[2]) {
      failures.push('例' + (index + 1) + '「' + testCase[3] + '」: 実際=' + actual + ', 期待=' + testCase[2]);
    }
  });

  if (failures.length) {
    throw new Error('自治体違い安全判定テスト失敗\n\n' + failures.join('\n'));
  }

  SpreadsheetApp.getUi().alert([
    '自治体違い安全判定テスト 成功',
    '',
    '成功件数: ' + cases.length + '件',
    '失敗件数: 0件',
    '元データ・Place ID・状態列の変更: なし'
  ].join('\n'));

  return { success: cases.length, failure: 0 };
}
