/**
 * Ver2.0 施設名の安全な付加表記差を再判定する。
 *
 * 目的:
 * - 同じPlace ID・同じ郵便番号・同じ所在地で、施設名の片方がもう片方を
 *   丸ごと含み、追加部分が説明語・温泉地名・読み仮名・英字併記などに
 *   とどまる場合だけ「却下候補」にする。
 * - 別邸・別館・新館・離れ・ANNEX・駅前・支店・Cafe等、別施設/別業態を
 *   示す可能性がある追加語は必ず要人確認に残す。
 *
 * B列「状態」と元データは変更しない。
 */

const HOTEL_DB_V2_SAFE_NAME_MIN_SCORE = 83;
const HOTEL_DB_V2_SAFE_NAME_MAX_EXTRA_LENGTH = 40;

const HOTEL_DB_V2_SAFE_NAME_RISK_TOKENS = Object.freeze([
  '別邸', '別館', '新館', '本館', '離れ', '別棟', '支店', '本店',
  '駅前', '空港', 'タワー', 'ウイング', 'ウィング', 'ヴィラ',
  'annex', 'branch', 'station', 'airport', 'tower', 'wing', 'villa',
  'east', 'west', 'north', 'south', 'room', 'floor'
]);

const HOTEL_DB_V2_SAFE_NAME_BUSINESS_RISK_TOKENS = Object.freeze([
  'cafe', 'coffee', 'restaurant', 'bar', 'pub', 'shop', 'sushi',
  'カフェ', '喫茶', 'レストラン', '食堂', '寿司', '売店'
]);

function runHotelDbV2SafeFacilityNameTriage() {
  return withHotelDbV2Lock_('施設名表記差の安全な再判定', function() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '施設名表記差を安全に再判定',
      '「修正候補」のうち、同じ所在地で施設名の前後に安全な説明表記が付いただけの候補を再判定します。\n\n' +
      'B列「状態」は変更しません。\n' +
      '元データも変更しません。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2RefineSafeFacilityNameTriage_();

    ui.alert([
      '施設名表記差の再判定完了',
      '',
      '確認件数: ' + result.scanned,
      '却下候補へ変更: ' + result.refined,
      '変更なし: ' + result.unchanged,
      '',
      'B列「状態」は変更していません。'
    ].join('\n'));

    return result;
  });
}

function hotelDbV2RefineSafeFacilityNameTriage_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS);

  if (!sheet || sheet.getLastRow() < 2) {
    return { scanned: 0, refined: 0, unchanged: 0 };
  }

  const headerMap = hotelDbV2TriageEnsureHeaders_(sheet);
  hotelDbV2TriageValidateHeaders_(headerMap);

  const values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .getDisplayValues();

  const recommendationColumn = headerMap['推奨判定'];
  const reasonColumn = headerMap['自動判定理由'];
  const confidenceColumn = headerMap['信頼度'];

  const result = {
    scanned: values.length,
    refined: 0,
    unchanged: 0
  };

  values.forEach(function(row, offset) {
    const currentRecommendation = hotelDbV2Clean_(
      row[recommendationColumn - 1]
    );

    if (currentRecommendation !== '要人確認') {
      result.unchanged++;
      return;
    }

    const input = hotelDbV2TriageInputFromRow_(row, headerMap);
    const decision = hotelDbV2SafeFacilityNameDecision_(input);

    if (!decision.equivalent) {
      result.unchanged++;
      return;
    }

    const rowNumber = offset + 2;
    sheet.getRange(rowNumber, recommendationColumn).setValue('却下候補');
    sheet.getRange(rowNumber, reasonColumn).setValue(decision.reason);
    sheet.getRange(rowNumber, confidenceColumn).setValue(decision.confidence);
    result.refined++;
  });

  return result;
}

function hotelDbV2SafeFacilityNameDecision_(input) {
  const data = input || {};

  if (hotelDbV2Clean_(data.state) === '反映済み') {
    return hotelDbV2SafeFacilityNameNo_('反映済み');
  }

  if (!hotelDbV2Clean_(data.placeId)) {
    return hotelDbV2SafeFacilityNameNo_('Place IDなし');
  }

  if (hotelDbV2Clean_(data.businessStatus) !== '営業中') {
    return hotelDbV2SafeFacilityNameNo_('営業中以外');
  }

  const score = Number(data.matchScore || 0);
  if (score < HOTEL_DB_V2_SAFE_NAME_MIN_SCORE) {
    return hotelDbV2SafeFacilityNameNo_('一致スコア不足');
  }

  const differences = hotelDbV2Clean_(data.differences)
    .split('・')
    .map(hotelDbV2Clean_)
    .filter(Boolean);

  if (differences.indexOf('施設名') === -1) {
    return hotelDbV2SafeFacilityNameNo_('施設名差分なし');
  }

  const unsupported = differences.filter(function(item) {
    return ['施設名', '郵便番号', '市区町村', '住所'].indexOf(item) === -1;
  });
  if (unsupported.length) {
    return hotelDbV2SafeFacilityNameNo_('対象外差分あり');
  }

  const sourcePostal = hotelDbV2NormalizePostalCode_(data.sourcePostalCode);
  const proposedPostal = hotelDbV2NormalizePostalCode_(data.proposedPostalCode);

  if (!sourcePostal || !proposedPostal || sourcePostal !== proposedPostal) {
    return hotelDbV2SafeFacilityNameNo_('郵便番号が実質不一致');
  }

  if (!hotelDbV2SafeNameLocationContextEquivalent_(data)) {
    return hotelDbV2SafeFacilityNameNo_('所在地が一致しない');
  }

  const relation = hotelDbV2SafeNameAffixRelation_(
    data.sourceName,
    data.proposedName
  );

  if (!relation.equivalent) {
    return hotelDbV2SafeFacilityNameNo_(relation.reason || '施設名関係が安全条件外');
  }

  return {
    equivalent: true,
    confidence: differences.length === 1 ? 98 : 97,
    reason:
      'Place ID・郵便番号・所在地が一致し、施設名の差は元名称を丸ごと保持した前後の付加表記だけです。別邸・別館・新館・離れ・駅前・支店・別業態などの危険語もないため、元施設名を維持する却下候補と判定しました。'
  };
}

function hotelDbV2SafeNameLocationContextEquivalent_(data) {
  const sourceMunicipality = hotelDbV2SafeNamePlain_(data.sourceMunicipality);
  const proposedMunicipality = hotelDbV2SafeNamePlain_(data.proposedMunicipality);
  const sourceAddress = hotelDbV2SafeNameCanonicalAddress_(data.sourceAddress);
  const proposedAddress = hotelDbV2SafeNameCanonicalAddress_(data.proposedAddress);

  if (
    sourceMunicipality &&
    proposedMunicipality &&
    hotelDbV2NormalizeText_(sourceMunicipality) ===
      hotelDbV2NormalizeText_(proposedMunicipality)
  ) {
    return Boolean(sourceAddress && sourceAddress === proposedAddress);
  }

  const countyMatch = sourceMunicipality.match(/^(.+?[都道府県])(.+?郡)$/u);
  if (!countyMatch) return false;

  const prefecture = countyMatch[1];
  const townMatch = hotelDbV2Clean_(data.sourceAddress)
    .normalize('NFKC')
    .replace(/\s+/g, '')
    .match(/^(.+?[町村])(.*)$/u);

  if (!townMatch) return false;

  const town = townMatch[1];
  const expectedGoogleMunicipality = prefecture + town;

  if (
    hotelDbV2NormalizeText_(proposedMunicipality) !==
    hotelDbV2NormalizeText_(expectedGoogleMunicipality)
  ) {
    return false;
  }

  const sourceFull = hotelDbV2SafeNameCanonicalAddress_(
    sourceMunicipality + data.sourceAddress
  );
  const proposedFull = hotelDbV2SafeNameCanonicalAddress_(
    prefecture + data.proposedAddress
  );

  return Boolean(sourceFull && sourceFull === proposedFull);
}

function hotelDbV2SafeNameAffixRelation_(sourceName, proposedName) {
  const source = hotelDbV2SafeFacilityNameCompact_(sourceName);
  const proposed = hotelDbV2SafeFacilityNameCompact_(proposedName);

  if (!source || !proposed) {
    return { equivalent: false, reason: '施設名が空欄' };
  }

  if (source === proposed) {
    return { equivalent: true, extra: '' };
  }

  let base = '';
  let extra = '';

  if (proposed.indexOf(source) === 0) {
    base = source;
    extra = proposed.slice(source.length);
  } else if (proposed.lastIndexOf(source) === proposed.length - source.length) {
    base = source;
    extra = proposed.slice(0, proposed.length - source.length);
  } else if (source.indexOf(proposed) === 0) {
    base = proposed;
    extra = source.slice(proposed.length);
  } else if (source.lastIndexOf(proposed) === source.length - proposed.length) {
    base = proposed;
    extra = source.slice(0, source.length - proposed.length);
  } else {
    return { equivalent: false, reason: '片方の施設名がもう片方を前後一致で含まない' };
  }

  if (base.length < 4) {
    return { equivalent: false, reason: '共通施設名が短すぎる' };
  }

  if (!extra) {
    return { equivalent: true, extra: '' };
  }

  if (extra.length > HOTEL_DB_V2_SAFE_NAME_MAX_EXTRA_LENGTH) {
    return { equivalent: false, reason: '追加表記が長すぎる' };
  }

  if (/\d/.test(extra)) {
    return { equivalent: false, reason: '追加表記に数字がある' };
  }

  const risk = HOTEL_DB_V2_SAFE_NAME_RISK_TOKENS.concat(
    HOTEL_DB_V2_SAFE_NAME_BUSINESS_RISK_TOKENS
  ).some(function(token) {
    return extra.indexOf(hotelDbV2SafeFacilityNameCompact_(token)) !== -1;
  });

  if (risk) {
    return { equivalent: false, reason: '追加表記に別施設・別業態の可能性がある語を含む' };
  }

  return { equivalent: true, extra: extra };
}

function hotelDbV2SafeFacilityNameCompact_(value) {
  return hotelDbV2Clean_(value)
    .normalize('NFKC')
    .toLowerCase()
    .replace(/＆/g, '&')
    .replace(/[\s　・･,，.．'’"“”\-ー―‐_/\\()（）\[\]【】〈〉《》<>:：]/g, '');
}

function hotelDbV2SafeNameCanonicalAddress_(value) {
  return hotelDbV2NormalizeAddress_(value)
    .replace(/大字/g, '')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
}

function hotelDbV2SafeNamePlain_(value) {
  return hotelDbV2Clean_(value)
    .normalize('NFKC')
    .replace(/\s+/g, '');
}

function hotelDbV2SafeFacilityNameNo_(reason) {
  return {
    equivalent: false,
    confidence: 0,
    reason: reason || ''
  };
}

/**
 * 自己診断テスト。入力オブジェクトを書き換えないことも確認する。
 */
function runHotelDbV2SafeFacilityNameTriageTests() {
  const base = {
    state: '未確認',
    sourcePostalCode: '680-1442',
    proposedPostalCode: '680-1442',
    sourceMunicipality: '鳥取県鳥取市',
    proposedMunicipality: '鳥取県鳥取市',
    sourceAddress: '吉岡温泉町632-1',
    proposedAddress: '吉岡温泉町632-1',
    sourceName: '摘草の宿 湯菜花',
    proposedName: '湯菜花',
    placeId: 'test-place-id',
    matchScore: 93,
    businessStatus: '営業中',
    differences: '施設名'
  };

  const cases = [
    {
      name: '説明的な宿表記の削除は安全候補',
      input: Object.assign({}, base),
      expected: true
    },
    {
      name: '国民宿舎の説明表記削除は安全候補',
      input: Object.assign({}, base, {
        sourceAddress: '浜坂1390-230',
        proposedAddress: '浜坂1390-230',
        sourceName: '民営国民宿舎 ニュー砂丘荘',
        proposedName: 'ニュー砂丘荘'
      }),
      expected: true
    },
    {
      name: '英字読み併記は安全候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '683-0001',
        proposedPostalCode: '683-0001',
        sourceMunicipality: '鳥取県米子市',
        proposedMunicipality: '鳥取県米子市',
        sourceAddress: '皆生温泉4-6-12',
        proposedAddress: '皆生温泉4-6-12',
        sourceName: 'やど紫苑亭',
        proposedName: 'やど紫苑亭 Yado shiontei'
      }),
      expected: true
    },
    {
      name: '英字B&B併記は安全候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '689-0535',
        proposedPostalCode: '689-0535',
        sourceAddress: '青谷町井手271-1',
        proposedAddress: '青谷町井手271-1',
        sourceName: 'カフェ＆ペンション デルマー',
        proposedName: 'カフェ&ペンション デルマー B&B Delmar pension'
      }),
      expected: true
    },
    {
      name: '温泉宿の説明接頭辞は安全候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '683-0001',
        proposedPostalCode: '683-0001',
        sourceMunicipality: '鳥取県米子市',
        proposedMunicipality: '鳥取県米子市',
        sourceAddress: '皆生温泉4-25-15',
        proposedAddress: '皆生温泉4-25-15',
        sourceName: '松涛園',
        proposedName: 'わんちゃんと泊まる温泉宿 皆生温泉 松涛園'
      }),
      expected: true
    },
    {
      name: '天然温泉等の説明接頭辞は安全候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '684-0004',
        proposedPostalCode: '684-0004',
        sourceMunicipality: '鳥取県境港市',
        proposedMunicipality: '鳥取県境港市',
        sourceAddress: '大正町216',
        proposedAddress: '大正町216',
        sourceName: '天然温泉 夕凪の湯 御宿 野乃境港',
        proposedName: '御宿 野乃境港'
      }),
      expected: true
    },
    {
      name: '郡町村差＋読み仮名併記も所在地一致なら安全候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '680-0415',
        proposedPostalCode: '680-0415',
        sourceMunicipality: '鳥取県八頭郡',
        proposedMunicipality: '鳥取県八頭町',
        sourceAddress: '八頭町下野331',
        proposedAddress: '八頭郡八頭町下野331',
        sourceName: 'OOE VALLEY STAY',
        proposedName: 'OOE VALLEY STAY〈オオエバレーステイ鳥取〉',
        matchScore: 83,
        differences: '市区町村・住所・施設名'
      }),
      expected: true
    },
    {
      name: '郡町村差＋旅館接頭辞は安全候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '682-0715',
        proposedPostalCode: '682-0715',
        sourceMunicipality: '鳥取県東伯郡',
        proposedMunicipality: '鳥取県湯梨浜町',
        sourceAddress: '湯梨浜町はわい温泉4-24',
        proposedAddress: '東伯郡湯梨浜町はわい温泉4-24',
        sourceName: '東郷館',
        proposedName: '旅館 東郷館',
        matchScore: 83,
        differences: '市区町村・住所・施設名'
      }),
      expected: true
    },
    {
      name: '全角郵便番号差解消後＋英字併記は安全候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '689-340１',
        proposedPostalCode: '689-3401',
        sourceMunicipality: '鳥取県米子市',
        proposedMunicipality: '鳥取県米子市',
        sourceAddress: '淀江町今津50-1',
        proposedAddress: '淀江町今津50-1',
        sourceName: '淀江の宿 今津田中家',
        proposedName: '淀江の宿 今津田中家 Yodoe Inn -Imazu Tanakaya-',
        matchScore: 83,
        differences: '郵便番号・施設名'
      }),
      expected: true
    },
    {
      name: 'Cafe追加は別業態の可能性があるため対象外',
      input: Object.assign({}, base, {
        sourcePostalCode: '680-0822',
        proposedPostalCode: '680-0822',
        sourceAddress: '今町2-276',
        proposedAddress: '今町2-276',
        sourceName: 'Drop Inn Tottori',
        proposedName: 'Cafe Drop Inn Tottori'
      }),
      expected: false
    },
    {
      name: '別邸追加は対象外',
      input: Object.assign({}, base, {
        sourcePostalCode: '682-0853',
        proposedPostalCode: '682-0853',
        sourceMunicipality: '鳥取県倉吉市',
        proposedMunicipality: '鳥取県倉吉市',
        sourceAddress: '余戸谷町2991-13',
        proposedAddress: '余戸谷町2991-13',
        sourceName: 'わんにゃんリゾートKURAYOSHI',
        proposedName: 'わんにゃんリゾートKURAYOSHI 別邸余戸谷'
      }),
      expected: false
    },
    {
      name: '離れ追加は対象外',
      input: Object.assign({}, base, {
        sourcePostalCode: '684-0004',
        proposedPostalCode: '684-0004',
        sourceMunicipality: '鳥取県境港市',
        proposedMunicipality: '鳥取県境港市',
        sourceAddress: '大正町53',
        proposedAddress: '大正町53',
        sourceName: '皆玉邸 恵-MEGUMI-',
        proposedName: '皆玉邸 恵-MEGUMI- 全室露天風呂付離れ'
      }),
      expected: false
    },
    {
      name: '所在地番地が違えば対象外',
      input: Object.assign({}, base, {
        sourceAddress: '吉岡温泉町632-1',
        proposedAddress: '吉岡温泉町633-1'
      }),
      expected: false
    },
    {
      name: 'スコア83未満は対象外',
      input: Object.assign({}, base, {
        matchScore: 82
      }),
      expected: false
    }
  ];

  const failures = [];

  cases.forEach(function(testCase, index) {
    const snapshot = JSON.stringify(testCase.input);
    const result = hotelDbV2SafeFacilityNameDecision_(testCase.input);
    const unchanged = JSON.stringify(testCase.input) === snapshot;

    if (result.equivalent !== testCase.expected || !unchanged) {
      failures.push(
        '例' + (index + 1) + '「' + testCase.name + '」: ' +
        '安全判定=' + result.equivalent +
        '（期待=' + testCase.expected + '）, ' +
        '入力保持=' + unchanged +
        (result.reason ? ', 理由=' + result.reason : '')
      );
    }
  });

  if (failures.length) {
    throw new Error(
      '施設名安全表記差テスト失敗\n\n' + failures.join('\n')
    );
  }

  SpreadsheetApp.getUi().alert([
    '施設名安全表記差テスト 成功',
    '',
    '成功件数: ' + cases.length + '件',
    '失敗件数: 0件',
    '状態列の自動変更: なし'
  ].join('\n'));

  return {
    success: cases.length,
    failure: 0,
    stateAutoChange: false
  };
}
