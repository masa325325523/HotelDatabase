/**
 * Ver2.0 重複候補の安全な自動仕分け。
 *
 * 「重複候補」シートの既存状態は変更せず、右端に
 * 推奨判定・自動判定理由・信頼度を付与する。
 * 元データの削除・更新は一切行わない。
 */

const HOTEL_DB_V2_DUPLICATE_TRIAGE_HEADERS = Object.freeze([
  '推奨判定',
  '自動判定理由',
  '信頼度'
]);

function runHotelDbV2DuplicateTriage() {
  return withHotelDbV2Lock_('重複候補の自動仕分け', function() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '重複候補を自動仕分け',
      '「重複候補」シートを安全条件で分類します。\n\n' +
      '元データは削除・変更しません。\n' +
      '「重複候補」シートの状態列も変更しません。\n' +
      '続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2DuplicateTriage_();

    ui.alert([
      '重複候補の自動仕分け完了',
      '',
      '確認件数: ' + result.scanned,
      '重複濃厚: ' + result.strongDuplicate,
      '要人確認: ' + result.needReview,
      '自動削除: 0',
      '状態列の変更: なし'
    ].join('\n'));

    return result;
  });
}

function hotelDbV2DuplicateTriage_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES);

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      scanned: 0,
      strongDuplicate: 0,
      needReview: 0
    };
  }

  const headerMap = hotelDbV2DuplicateTriageEnsureHeaders_(sheet);
  hotelDbV2DuplicateTriageValidateHeaders_(headerMap);

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const values = sheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getDisplayValues();

  const result = {
    scanned: 0,
    strongDuplicate: 0,
    needReview: 0
  };

  values.forEach(function(row, offset) {
    if (!hotelDbV2DuplicateTriageHasData_(row, headerMap)) return;

    const input = {
      existingDecision: row[headerMap['判定'] - 1],
      name1: row[headerMap['施設名1'] - 1],
      address1: row[headerMap['住所1'] - 1],
      name2: row[headerMap['施設名2'] - 1],
      address2: row[headerMap['住所2'] - 1],
      placeId: row[headerMap['Place ID'] - 1],
      similarity: row[headerMap['類似度'] - 1],
      state: row[headerMap['状態'] - 1]
    };

    const decision = hotelDbV2DuplicateTriageDecision_(input);
    const rowNumber = offset + 2;

    sheet.getRange(rowNumber, headerMap['推奨判定']).setValue(decision.recommendation);
    sheet.getRange(rowNumber, headerMap['自動判定理由']).setValue(decision.reason);
    sheet.getRange(rowNumber, headerMap['信頼度']).setValue(decision.confidence);

    result.scanned++;
    if (decision.recommendation === '重複濃厚') {
      result.strongDuplicate++;
    } else {
      result.needReview++;
    }
  });

  return result;
}

function hotelDbV2DuplicateTriageDecision_(input) {
  const data = input || {};
  const name1 = hotelDbV2NormalizeText_(data.name1);
  const name2 = hotelDbV2NormalizeText_(data.name2);
  const address1 = hotelDbV2NormalizeAddress_(data.address1);
  const address2 = hotelDbV2NormalizeAddress_(data.address2);
  const placeId = hotelDbV2Clean_(data.placeId);

  if (!name1 || !name2 || !address1 || !address2) {
    return hotelDbV2DuplicateTriageReview_(
      '施設名または住所が不足しているため、自動で重複濃厚にはできません。',
      60
    );
  }

  const sameName = name1 === name2;
  const sameAddress = address1 === address2;

  if (sameName && sameAddress) {
    return {
      recommendation: '重複濃厚',
      confidence: placeId ? 99 : 97,
      reason: placeId
        ? '施設名・住所が正規化後も完全一致し、Place IDも同一です。自動削除はせず、重複濃厚として確認対象にします。'
        : '施設名・住所が正規化後も完全一致しています。Place IDがなくても重複の可能性が高いですが、自動削除はしません。'
    };
  }

  if (placeId && !sameAddress) {
    return hotelDbV2DuplicateTriageReview_(
      'Place IDは同一ですが住所が異なります。同一建物の別階・別部屋・別棟などの可能性があるため、人が確認します。',
      98
    );
  }

  if (sameAddress && !sameName) {
    return hotelDbV2DuplicateTriageReview_(
      placeId
        ? 'Place IDと住所は同一ですが施設名が異なります。別邸・別館・館内別施設などの可能性があるため、人が確認します。'
        : '住所は同一ですが施設名が異なります。同一住所に複数施設が存在する可能性があるため、人が確認します。',
      placeId ? 97 : 94
    );
  }

  if (sameName && !sameAddress) {
    return hotelDbV2DuplicateTriageReview_(
      '施設名は同一ですが住所が異なります。移転・支店・別棟・別部屋などの可能性があるため、人が確認します。',
      94
    );
  }

  return hotelDbV2DuplicateTriageReview_(
    '施設名または住所に実質的な差があります。類似度だけでは重複確定できないため、人が確認します。',
    90
  );
}

function hotelDbV2DuplicateTriageReview_(reason, confidence) {
  return {
    recommendation: '要人確認',
    confidence: Number(confidence || 0),
    reason: reason || ''
  };
}

function hotelDbV2DuplicateTriageHasData_(row, headerMap) {
  return !!(
    hotelDbV2Clean_(row[headerMap['施設名1'] - 1]) ||
    hotelDbV2Clean_(row[headerMap['住所1'] - 1]) ||
    hotelDbV2Clean_(row[headerMap['施設名2'] - 1]) ||
    hotelDbV2Clean_(row[headerMap['住所2'] - 1])
  );
}

function hotelDbV2DuplicateTriageEnsureHeaders_(sheet) {
  const existing = sheet
    .getRange(1, 1, 1, Math.max(1, sheet.getLastColumn()))
    .getDisplayValues()[0];
  const map = {};

  existing.forEach(function(value, index) {
    const header = hotelDbV2Clean_(value);
    if (header) map[header] = index + 1;
  });

  HOTEL_DB_V2_DUPLICATE_TRIAGE_HEADERS.forEach(function(header) {
    if (map[header]) return;
    const column = sheet.getLastColumn() + 1;
    sheet.getRange(1, column).setValue(header);
    map[header] = column;
  });

  return map;
}

function hotelDbV2DuplicateTriageValidateHeaders_(headerMap) {
  const required = [
    '判定', '施設名1', '住所1', '施設名2', '住所2',
    'Place ID', '類似度', '状態',
    '推奨判定', '自動判定理由', '信頼度'
  ];

  const missing = required.filter(function(header) {
    return !headerMap[header];
  });

  if (missing.length) {
    throw new Error(
      '「重複候補」シートに必要な列がありません: ' + missing.join(', ')
    );
  }
}

/**
 * 重複候補の安全な自動仕分けに関する自己診断テスト。
 */
function runHotelDbV2DuplicateTriageTests() {
  const cases = [
    {
      name: '施設名・住所・Place IDが完全一致',
      input: {
        name1: '満月荘', address1: '鳥取県鳥取市商栄町108-1',
        name2: '満月荘', address2: '鳥取県鳥取市商栄町108-1',
        placeId: 'same-place', state: '未確認'
      },
      expected: '重複濃厚'
    },
    {
      name: '全角半角と空白の表記差だけ',
      input: {
        name1: 'ホテル 山田', address1: '鳥取県鳥取市栄町２３０',
        name2: 'ホテル山田', address2: '鳥取県鳥取市栄町230',
        placeId: 'same-place', state: '未確認'
      },
      expected: '重複濃厚'
    },
    {
      name: '丁目番地表記差だけ',
      input: {
        name1: 'テスト旅館', address1: '鳥取県鳥取市今町二丁目153番地',
        name2: 'テスト旅館', address2: '鳥取県鳥取市今町2-153',
        placeId: 'same-place', state: '未確認'
      },
      expected: '重複濃厚'
    },
    {
      name: '大字の有無だけ',
      input: {
        name1: '門脇旅館', address1: '鳥取県日野郡江府町大字江尾2064',
        name2: '門脇旅館', address2: '鳥取県日野郡江府町江尾2064',
        placeId: 'same-place', state: '未確認'
      },
      expected: '重複濃厚'
    },
    {
      name: 'Place IDなしでも施設名住所完全一致',
      input: {
        name1: '同一旅館', address1: '鳥取県鳥取市本町1-1',
        name2: '同一旅館', address2: '鳥取県鳥取市本町1-1',
        placeId: '', state: '未確認'
      },
      expected: '重複濃厚'
    },
    {
      name: '同じPlace IDでも901号室と301号室は要人確認',
      input: {
        name1: 'Japan Hinata 1', address1: '名古屋市千種区春岡一丁目4番21号 Fuchsia901',
        name2: 'Japan Hinata 2', address2: '名古屋市千種区春岡一丁目4番21号 Fuchsia301',
        placeId: 'same-place', state: '未確認'
      },
      expected: '要人確認'
    },
    {
      name: '同じPlace IDでも4階と6階は要人確認',
      input: {
        name1: 'おやこホテル', address1: '名古屋市千種区東山通2丁目4番地の1 HARVEY MOTOYAMA 4階',
        name2: 'おやこホテル', address2: '名古屋市千種区東山通2丁目4番地の1 HARVEY MOTOYAMA 6階',
        placeId: 'same-place', state: '未確認'
      },
      expected: '要人確認'
    },
    {
      name: '同じPlace IDでも番地違いは要人確認',
      input: {
        name1: '同名旅館', address1: '鳥取県鳥取市本町1-1',
        name2: '同名旅館', address2: '鳥取県鳥取市本町1-2',
        placeId: 'same-place', state: '未確認'
      },
      expected: '要人確認'
    },
    {
      name: '同住所でも別施設名は要人確認',
      input: {
        name1: '旅館A', address1: '鳥取県鳥取市本町1-1',
        name2: '旅館B', address2: '鳥取県鳥取市本町1-1',
        placeId: '', state: '未確認'
      },
      expected: '要人確認'
    },
    {
      name: '同住所で類似名でも要人確認',
      input: {
        name1: 'おやこホテル', address1: '名古屋市千種区東山通2丁目4番地の1',
        name2: 'おやこホテル 名古屋東山4F', address2: '名古屋市千種区東山通2丁目4番地の1',
        placeId: '', state: '未確認'
      },
      expected: '要人確認'
    },
    {
      name: '同じPlace IDと住所でも別邸表記差は要人確認',
      input: {
        name1: 'わんにゃんリゾートKURAYOSHI', address1: '鳥取県倉吉市余戸谷町1-1',
        name2: 'わんにゃんリゾートKURAYOSHI 別邸余戸谷', address2: '鳥取県倉吉市余戸谷町1-1',
        placeId: 'same-place', state: '未確認'
      },
      expected: '要人確認'
    },
    {
      name: '同施設名でも住所違いは要人確認',
      input: {
        name1: '同名ホテル', address1: '鳥取県米子市明治町125',
        name2: '同名ホテル', address2: '鳥取県米子市明治町175',
        placeId: '', state: '未確認'
      },
      expected: '要人確認'
    },
    {
      name: '施設名1が空欄なら要人確認',
      input: {
        name1: '', address1: '鳥取県鳥取市本町1-1',
        name2: '旅館A', address2: '鳥取県鳥取市本町1-1',
        placeId: 'same-place', state: '未確認'
      },
      expected: '要人確認'
    },
    {
      name: '住所2が空欄なら要人確認',
      input: {
        name1: '旅館A', address1: '鳥取県鳥取市本町1-1',
        name2: '旅館A', address2: '',
        placeId: 'same-place', state: '未確認'
      },
      expected: '要人確認'
    },
    {
      name: 'Place ID一致という既存判定だけでは重複濃厚にしない',
      input: {
        existingDecision: 'Place ID一致',
        name1: '施設A', address1: '名古屋市千種区春岡1-4-21 901号室',
        name2: '施設B', address2: '名古屋市千種区春岡1-4-21 301号室',
        placeId: 'same-place', state: '未確認'
      },
      expected: '要人確認'
    },
    {
      name: '状態が確認済みでも判定関数は入力を変更しない',
      input: {
        name1: '同一旅館', address1: '鳥取県鳥取市本町1-1',
        name2: '同一旅館', address2: '鳥取県鳥取市本町1-1',
        placeId: 'same-place', state: '確認済み'
      },
      expected: '重複濃厚'
    }
  ];

  const failures = [];

  cases.forEach(function(testCase, index) {
    const snapshot = JSON.stringify(testCase.input);
    const decision = hotelDbV2DuplicateTriageDecision_(testCase.input);
    const unchanged = JSON.stringify(testCase.input) === snapshot;

    if (decision.recommendation !== testCase.expected || !unchanged) {
      failures.push(
        '例' + (index + 1) + '「' + testCase.name + '」: ' +
        '推奨判定=' + decision.recommendation +
        '（期待=' + testCase.expected + '）, ' +
        '入力保持=' + unchanged +
        (decision.reason ? ', 理由=' + decision.reason : '')
      );
    }
  });

  if (failures.length) {
    const message = [
      '重複候補安全仕分けテストに失敗しました。',
      '',
      failures.join('\n')
    ].join('\n');
    console.error(message);
    throw new Error(message);
  }

  const successMessage = [
    '重複候補安全仕分けテスト 成功',
    '',
    '成功件数: ' + cases.length + '件',
    '失敗件数: 0件',
    '自動削除: なし',
    '状態列の自動変更: なし'
  ].join('\n');

  console.log(successMessage);
  try {
    SpreadsheetApp.getUi().alert(successMessage);
  } catch (error) {
  }

  return {
    success: true,
    passed: cases.length,
    failed: 0,
    autoDelete: false,
    stateAutoChange: false
  };
}
