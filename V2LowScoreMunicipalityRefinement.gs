/**
 * PR #14 低スコア要確認の追加安全判定。
 *
 * 目的:
 * - 明確な自治体違いかつ施設名の一致証拠が弱い場合だけ「誤候補有力」へ寄せる。
 * - 同名・高類似名、または同一自治体内の大きな名称差は、移転・住所誤り・リブランド等を考慮し
 *   「要人確認」へ戻す。
 * - 元データ・Place ID・状態列は変更しない。
 */

function runHotelDbV2LowScoreReviewTriageWithMunicipalityRefinement() {
  return withHotelDbV2Lock_('低スコア要確認自動仕分け', function() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '低スコア要確認を自動仕分け',
      '「要確認」の低スコア行を、安全側に仕分けします。\n\n' +
      '本体の自動採用基準75点は変更しません。\n' +
      '状態・元データ・Place IDも変更しません。\n' +
      '同名・高類似名や同一自治体内の名称変更候補は、人確認へ残します。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2TriageLowScoreReviews_();
    const refined = hotelDbV2RefineLowScoreSafety_();

    result.wrongLikely = Math.max(
      0,
      result.wrongLikely + refined.promotedToWrong - refined.demotedToHuman
    );
    result.humanReview = Math.max(
      0,
      result.humanReview - refined.promotedToWrong + refined.demotedToHuman
    );

    ui.alert([
      '低スコア要確認の自動仕分け完了',
      '',
      '対象件数: ' + result.total,
      '同一施設有力: ' + result.sameLikely,
      '誤候補有力: ' + result.wrongLikely,
      '要人確認: ' + result.humanReview,
      '対象外: ' + result.outOfScope,
      '自治体違いで誤候補へ: ' + refined.promotedToWrong,
      '安全側へ要人確認に戻す: ' + refined.demotedToHuman,
      '',
      '状態・元データ・Place IDは変更していません。'
    ].join('\n'));

    return result;
  });
}

function hotelDbV2RefineLowScoreSafety_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.REVIEW);
  if (!sheet || sheet.getLastRow() < 2) {
    return { promotedToWrong: 0, demotedToHuman: 0 };
  }

  const map = hotelDbV2LowScoreHeaderMap_(sheet);
  hotelDbV2LowScoreValidateHeaders_(map);

  const rowCount = sheet.getLastRow() - 1;
  const values = sheet
    .getRange(2, 1, rowCount, sheet.getLastColumn())
    .getDisplayValues();

  const recommendations = [];
  const reasons = [];
  const confidences = [];
  let promotedToWrong = 0;
  let demotedToHuman = 0;

  values.forEach(function(row) {
    const currentRecommendation = hotelDbV2Clean_(row[map['推奨判定'] - 1]);
    let recommendation = currentRecommendation;
    let reason = hotelDbV2Clean_(row[map['自動判定理由'] - 1]);
    let confidence = Number(hotelDbV2Clean_(row[map['信頼度'] - 1]) || 0);

    const input = hotelDbV2LowScoreInputFromRow_(row, map);
    const decision = hotelDbV2RefineLowScoreDecision_(currentRecommendation, input);

    if (decision.recommendation !== currentRecommendation) {
      recommendation = decision.recommendation;
      reason = decision.reason;
      confidence = decision.confidence;

      if (currentRecommendation === '要人確認' && recommendation === '誤候補有力') {
        promotedToWrong++;
      }
      if (currentRecommendation === '誤候補有力' && recommendation === '要人確認') {
        demotedToHuman++;
      }
    }

    recommendations.push([recommendation]);
    reasons.push([reason]);
    confidences.push([confidence]);
  });

  sheet.getRange(2, map['推奨判定'], rowCount, 1).setValues(recommendations);
  sheet.getRange(2, map['自動判定理由'], rowCount, 1).setValues(reasons);
  sheet.getRange(2, map['信頼度'], rowCount, 1).setValues(confidences);

  return {
    promotedToWrong: promotedToWrong,
    demotedToHuman: demotedToHuman
  };
}

function hotelDbV2RefineLowScoreDecision_(currentRecommendation, input) {
  const data = input || {};
  const sourceName = hotelDbV2NormalizeText_(data.sourceName);
  const candidateName = hotelDbV2NormalizeText_(data.candidateName);
  const nameSimilarity = hotelDbV2SimilarityRatio_(sourceName, candidateName);
  const nameContains = hotelDbV2LowScoreNameContains_(sourceName, candidateName);
  const nameStrong =
    sourceName === candidateName ||
    nameContains ||
    nameSimilarity >= HOTEL_DB_V2_LOW_SCORE_RULES.STRONG_NAME_SIMILARITY;

  const municipalitySame = hotelDbV2LowScoreMunicipalityMatches_(
    data.municipality,
    data.candidateAddress
  );
  const municipalityConflict = hotelDbV2LowScoreHasClearMunicipalityConflict_(
    data.municipality,
    data.candidateAddress
  );

  const isLowScoreRow =
    HOTEL_DB_V2_LOW_SCORE_REASONS.indexOf(hotelDbV2Clean_(data.reason)) !== -1 &&
    Number(data.matchScore || 0) < HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE;

  const candidateUsable =
    !!hotelDbV2Clean_(data.candidatePlaceId) &&
    hotelDbV2Clean_(data.businessStatus) === '営業中';

  if (!isLowScoreRow || !candidateUsable) {
    return {
      recommendation: currentRecommendation,
      reason: '',
      confidence: 0
    };
  }

  // 同名・高類似名なのに自治体が違う場合、元住所・自治体側の誤りや移転の可能性もある。
  // 「誤候補有力」と断定せず、人確認へ戻す。
  if (currentRecommendation === '誤候補有力' && nameStrong) {
    return {
      recommendation: '要人確認',
      reason: '施設名が同一・包含・高類似のため、自治体や住所が異なっていても元データ誤り・移転等を否定できません。誤候補とは断定せず、人が確認します。',
      confidence: 98
    };
  }

  // 同一自治体内で施設名が大きく変わっている低スコア候補は、リブランド・名称変更の可能性がある。
  // 住所一致が弱くても「誤候補」とは断定しない。
  if (currentRecommendation === '誤候補有力' && municipalitySame) {
    return {
      recommendation: '要人確認',
      reason: '候補は元施設と同じ市区町村内です。施設名が大きく異なっていても、リブランド・名称変更・元住所誤りの可能性があるため、人が確認します。',
      confidence: 98
    };
  }

  // 人確認に残っている行のうち、自治体が明確に異なり、施設名の一致証拠も弱い場合だけ誤候補有力へ。
  if (
    currentRecommendation === '要人確認' &&
    municipalityConflict &&
    !nameStrong
  ) {
    return {
      recommendation: '誤候補有力',
      reason: '候補住所の自治体が元の市区町村と明確に異なり、施設名の一致証拠も弱いため、別施設候補の可能性が高いです。自動削除・自動反映はしません。',
      confidence: 98
    };
  }

  return {
    recommendation: currentRecommendation,
    reason: '',
    confidence: 0
  };
}

function hotelDbV2LowScoreHasClearMunicipalityConflict_(municipality, candidateAddress) {
  const source = hotelDbV2LowScoreExtractMunicipalityParts_(municipality);
  const candidate = hotelDbV2LowScoreExtractMunicipalityParts_(candidateAddress);

  if (source.city && candidate.city && source.city !== candidate.city) return true;

  if (
    source.city && candidate.city && source.city === candidate.city &&
    source.ward && candidate.ward && source.ward !== candidate.ward
  ) return true;

  if (
    !source.city && !candidate.city &&
    source.ward && candidate.ward && source.ward !== candidate.ward
  ) {
    return true;
  }

  if (
    !source.city && !source.ward &&
    source.town && candidate.town && source.town !== candidate.town
  ) {
    return true;
  }

  if (
    !source.city && !source.ward &&
    source.village && candidate.village && source.village !== candidate.village
  ) {
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
    const local = text.replace(/^.{1,12}?郡/, '');
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
  const conflictCases = [
    ['大阪市北区', '兵庫県神戸市中央区北野町1丁目', true, '大阪→神戸'],
    ['大阪市北区', '大阪府大阪市北区芝田2丁目', false, '同じ大阪市北区'],
    ['大阪市北区', '大阪府大阪市中央区心斎橋筋1丁目', true, '同じ市の別区'],
    ['東京都新宿区', '東京都渋谷区道玄坂1丁目', true, '東京23区の別区'],
    ['鳥取県岩美郡', '鳥取県岩美郡岩美町岩井536', false, '郡のみの元表記は断定しない'],
    ['岩美町', '鳥取県八頭郡八頭町郡家648', true, '町が明確に異なる']
  ];

  const base = {
    state: '未確認',
    municipality: '大阪市北区',
    sourceAddress: '大阪府大阪市北区神山町8-4',
    sourceName: 'サンプルホテル大阪',
    reason: '一致スコア不足',
    candidateName: '別ブランドホテル',
    candidateAddress: '大阪府大阪市北区神山町8-4',
    candidatePlaceId: 'TEST_PLACE',
    matchScore: 30,
    businessStatus: '営業中',
    detail: '自動採用基準=75'
  };

  function make(overrides) {
    const item = {};
    Object.keys(base).forEach(function(key) { item[key] = base[key]; });
    Object.keys(overrides || {}).forEach(function(key) { item[key] = overrides[key]; });
    return item;
  }

  const decisionCases = [
    {
      name: '同名・自治体違いは人確認へ戻す',
      current: '誤候補有力',
      input: make({
        municipality: '大阪市生野区',
        sourceName: '今里旅館',
        candidateName: '今里旅館',
        candidateAddress: '大阪府大阪市東成区大今里南2丁目5-1',
        matchScore: 58
      }),
      expected: '要人確認'
    },
    {
      name: '高類似ブランド名・別都市は人確認を維持',
      current: '要人確認',
      input: make({
        sourceName: 'ANAクラウンプラザホテル大阪',
        candidateName: 'ANAクラウンプラザホテル神戸',
        candidateAddress: '兵庫県神戸市中央区北野町1丁目',
        matchScore: 35
      }),
      expected: '要人確認'
    },
    {
      name: '同一自治体内の大きな名称差はリブランド考慮で人確認',
      current: '誤候補有力',
      input: make({
        sourceName: '旧ブランドホテル大阪梅田',
        candidateName: 'KOKO HOTEL 大阪梅田',
        municipality: '大阪市北区',
        candidateAddress: '大阪府大阪市北区神山町8-4',
        matchScore: 17
      }),
      expected: '要人確認'
    },
    {
      name: '別自治体・弱い施設名だけ誤候補有力へ',
      current: '要人確認',
      input: make({
        sourceName: '大阪サンプル旅館',
        candidateName: '神戸ゲストハウス',
        candidateAddress: '兵庫県神戸市中央区北野町1丁目',
        matchScore: 40
      }),
      expected: '誤候補有力'
    }
  ];

  const failures = [];

  conflictCases.forEach(function(testCase, index) {
    const actual = hotelDbV2LowScoreHasClearMunicipalityConflict_(
      testCase[0],
      testCase[1]
    );
    if (actual !== testCase[2]) {
      failures.push(
        '自治体例' + (index + 1) + '「' + testCase[3] + '」: ' +
        '実際=' + actual + ', 期待=' + testCase[2]
      );
    }
  });

  decisionCases.forEach(function(testCase, index) {
    const actual = hotelDbV2RefineLowScoreDecision_(
      testCase.current,
      testCase.input
    );
    if (actual.recommendation !== testCase.expected) {
      failures.push(
        '安全例' + (index + 1) + '「' + testCase.name + '」: ' +
        '実際=' + actual.recommendation + ', 期待=' + testCase.expected
      );
    }
  });

  if (failures.length) {
    throw new Error(
      '自治体違い・リブランド安全判定テスト失敗\n\n' + failures.join('\n')
    );
  }

  const successCount = conflictCases.length + decisionCases.length;
  SpreadsheetApp.getUi().alert([
    '自治体違い・リブランド安全判定テスト 成功',
    '',
    '成功件数: ' + successCount + '件',
    '失敗件数: 0件',
    '同名・高類似名: 誤候補へ断定しない',
    '同一自治体内の名称変更候補: 人確認へ残す',
    '元データ・Place ID・状態列の変更: なし'
  ].join('\n'));

  return {
    success: successCount,
    failure: 0,
    sourceAutoChange: false
  };
}
