// ============================================================
// 30_price_master_import.gs
// Excelマスタ表からの初期データ取込み
//   ①「基板PCB価格」シート → 基板PCB価格マスタ（原価・公表単価の手動管理表）
//   ②「機種メモ」シート（部品調達計画シートの最新スナップショット部分から
//      機種ごとの主基板/払出基板/D/E/DE基板構成を抽出したもの）
//      → 機種基板構成
//
// 元データ: マスタ表（アミューズメント事業部、遊技機一括管理用アプリ）.xlsx
//   ・基板PCB価格シート … 54行をそのまま取込み
//   ・機種メモシート    … 2029行×128列の四半期スナップショットの繰り返し
//     ブロックから「主基板/払出基板/液晶制御(D)/演出基板(E)/液晶IF基板(DE)」
//     行を持つブロックのみを抽出し、機種コードが後方（＝新しい）のブロックで
//     上書きすることで最新構成のみを残した（118機種分）。
//
// apiSeedMasterPriceData() は既存シートを一度クリアしてこの初期データで
// 上書きするため、実行後に手入力で編集した内容は再実行すると失われる点に注意。
// ============================================================

var MASTER_BOARD_PRICE_SHEET = '基板PCB価格マスタ';
var MODEL_BOARD_COMPOSITION_SHEET = '機種基板構成';

var MASTER_BOARD_PRICE_HEADERS = [
  '基板グループ', '基板コード', '種類', 'VDP', 'PCB原価', '基板メーカー',
  '仕掛原価', '実装費', '仕掛原価(PCB抜き)', '部品表(K10)',
  'PCB公表単価', '仕掛公表単価', '公表単価合計',
  '過去見積リンク1', '過去見積リンク2', '基板販売日', '総販売数',
  'ブランド', '設計課情報', '採用機種',
];

var MODEL_BOARD_COMPOSITION_HEADERS = [
  '機種コード', '主基板', '払出基板', '液晶制御(D)', '演出基板(E)', '液晶IF基板(DE)',
];

var MASTER_BOARD_PRICE_DATA = [["主制御基板", "M1602B", "スロット", "", 389, "板", "", 372, 3127, "FJ+M1602B.xls", "", 5068, 5068, "", "", "", "", "", "", ""], ["", "M1805D", "スロット", "", 475, "板", "", 397, 2399, "JJ+M1805D.xls", "", 4375, 4375, "", "", "", "", "", "", ""], ["", "M2104A", "スロット", "", 491, "板", "", 581, 2592, "", "", 4729, 4729, "", "", "", "", "14-04_基板仕様書_M2104A_220207.pdf", "HB237(FJ+M2104A2)", ""], ["", "M2104A1", "スロット", "", "", "", "", "", "", "", "", "", 0, "", "", "", "", "", "", ""], ["", "M2402A", "スロット", "", 551, "板", "", 470, 2692, "FJ+M2402A.xls", 762, 3716, 4478, "", "", "", "", "", "HB239(FJ+M2402A)", ""], ["", "M2401A", "パチンコ", "", 558, "板", "", 441, 1333, "FJ+M2104A.xls", 731, 1753, 2484, "", "", "", "", "14-04_基板仕様書(FJ+M2401)_240828.pdf", "HB238(FJ+M2401A)", ""], ["", "M2503A", "パチンコ", "", 321, "板", "", 424.4, 1101, "", 407, 1923, 2330, "", "", "", "", "", "HB242(FJ+M2503A)：PFK30", ""], ["液晶制御基板（D）", "D1401C", "パチンコ/スロット", "", 490, "リンク", "", 434, 8571, "", "", 12449, 12449, "", "", "", "", "", "", ""], ["", "D2101A", "パチンコ/スロット", "", 416, "キョウデン", "", 544, 8486, "", 537, 12605, 13142, "", "", "", "", "", "", ""], ["", "D2101A1", "パチンコ/スロット", "", "", "", "", "", "", "", "", "", 0, "", "", "", "", "", "", ""], ["", "D1901B", "", "AG5", "", "", "", "", "", "", "", "", 0, "", "", "", "", "14-04_基板設計仕様書(D1901B)_200114.pdf", "HD185(FJ+D1901B)", ""], ["液晶IF基板（DE）", "DE2502A", "", "", 354, "リンク", "", "-", "", "", 574, 2057, 2631, "", "", "", "", "", "", ""], ["", "DE1802A", "パチンコ", "", 343, "リンク", "", 188, 1089, "FJ+DE1802A.xls", 460, 1858, 2318, "", "", "", "", "", "", ""], ["", "DE2101A", "パチンコ", "", 352, "板", "", 247, 1348, "FJ+DE2101A.xls", "", 2729, 2729, "", "", "", "", "", "", ""], ["", "DE2103A1", "パチンコ", "", "", "", "", "", "", "", "", "", 0, "", "", "", "", "", "", ""], ["", "DE2101A1◎", "パチンコ", "", 399, "リンク", "", 217, 1307, "FJ+DE2101A1.xls", "", 2362, 2362, "https://drive.google.com/file/d/1RKlA-tDzJFdMvjdw5FLmVM3JpXpqEekk/view?usp=drive_link", "https://drive.google.com/file/d/1h_0h6fpsggcxUSel7vhXCUF2wpAthHps/view?usp=drive_link", "", "", "", "HD200(FJ+DE2101A_VerB)", ""], ["", "DE2101A", "パチンコ", "", "", "", "", "", "", "", "", "", 0, "", "", "", "", "", "HD239(FJ+DE2502A)", ""], ["", "DE1607A", "スロット", "", 343, "リンク", "", 175, 984, "FJ+DE1607A.xls", "", 2054, 2054, "https://drive.google.com/file/d/1xxTQsOsj1KSF6Obld0Z_GzeqiNcX-Fk8/view?usp=drive_link", "https://drive.google.com/file/d/17SaNySunlEbcty8khL8R-Ig_QxK1yf5H/view?usp=drive_link", "", "", "", "", ""], ["", "DE1606A", "スロット", "", 343, "リンク", "", 201, 1241, "FJ+DE1606A.xls", "", 2486, 2486, "https://drive.google.com/file/d/1xxTQsOsj1KSF6Obld0Z_GzeqiNcX-Fk8/view?usp=drive_link", "", "", "", "", "", ""], ["", "DE1802A", "スロット", "", "", "リンク", "", "", "", "", "", "", 0, "", "", "", "", "", "", ""], ["演出基板（E）", "E1501C", "スロット", "", 699, "リンク", "", 493, 5472, "FJ+E1501C.xls", "", 7133, 7133, "", "", "", "", "", "", ""], ["", "E1901B", "", "", 556, "板", "", 633, 5019, "FJ+E1901B.xls", "", 7100, 7100, "", "", "", "", "", "", ""], ["", "E1902B", "スロット", "", 742, "板", "", 654, 5244, "FJ+E1902B.xls", "", 7571, 7571, "", "", "", "", "", "", ""], ["", "E2102B", "スロット", "", 887, "板", "", 666, "", "", 1097, 8089, 9186, "", "", "", "", "", "", ""], ["", "E2002B", "パチンコ", "", 556, "板", "", 648, "", "", "", 7244, 7244, "", "", "", "", "14-04_基板設計仕様書(E2002B)_210708.pdf", "HG181(FJ+E2002B)", ""], ["", "E2101B", "パチンコ", "", 887, "板", "", 539, "", "", 1124, 8049, 9173, "", "", "", "", "14-04_基板設計仕様書(E2101B)_211125.pdf", "HG183(FJ+E2101B)", ""], ["", "E2102B", "パチンコ", "", "", "", "", "", "", "", "", "", 0, "", "", "", "", "14-04_基板設計仕様書_FJ+E2102B.pdf", "HG182(FJ+E2102B)", ""], ["", "E2301B", "パチンコ", "", 734, "板", "", 509, "", "", 956, 4942, 5898, "", "", "", "", "14-04_基板設計仕様書(E2301B)_240228.pdf", "HG184(FJ+E2301A)", ""], ["", "E2501B", "パチンコ", "", "", "", "", "", "", "", "", "", 0, "", "", "", "", "14-04_基板設計仕様書(TEST2501)_250619.pdf", "FJ+E2501B", ""], ["", "E2503A", "パチンコ", "", "", "", "", "", "", "", "", "", 0, "", "", "", "", "", "FJ+E2503A", ""], ["", "E2501B", "パチンコ", "AG6R", 699, "板", "", 476, "", "", 974, 4700, 5674, "", "", "", "", "14-04_基板設計仕様書(TEST2501)_250619.pdf", "", ""], ["", "E2503B", "", "", 814, "リンク", "", "-", "", "", 1118, 8449, 9567, "", "", "", "", "", "", ""], ["サブ制御基板（C）", "C1601E", "", "", "", "", "", "", "", "", "", "", 0, "", "", "", "", "", "", ""], ["", "C1901A", "", "", "", "", "", "", "", "", "", "", 0, "", "", "", "", "14-04_(仮)基板仕様書_FJ C1901A_191024.pdf", "", ""], ["", "C2101B", "", "", 378, "板", "", 343, "", "", 445, 1902, 2347, "", "", "", "", "", "", ""], ["", "C2401A", "", "", 749, "板", "", 630.4, "", "", 941, 3624, 4565, "", "", "", "", "", "", ""], ["", "C2202B", "", "", "", "", "", "", "", "", "", "", 0, "", "", "", "", "14-04_FJ+C2202A_基板仕様書_230117.pdf", "HC030(FJ+C2202B)：PF45", ""], ["", "C2501C", "", "", 728, "板", "", 502, "", "", 854, 3169, 4023, "", "", "", "", "", "HC036(FJ+C2501C)：PFK30", ""], ["", "C2502A", "", "", 352, "板", "", 350, "", "", 491, 1926, 2417, "", "", "", "", "", "HC037(FJ+C2502A)：SFK18", ""], ["", "C2501B", "", "", "", "", "", "", "", "", "", "", 0, "", "", "", "", "", "HC035(FJ+C2501B)：PFK30", ""], ["SNB基板", "SNB5163A-00", "", "", "", "", "", "https://drive.google.com/file/d/1lxty11pjp_7TkTUsZEFc4P2DxQVkjbEw/view?usp=drive_link", "", "", "", "", 8146, "", "", "", "", "", "", ""], ["KAM1121A", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["KAM1130A", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["回胴", "SF20", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["スマート回胴", "SFK10", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["スマート回胴", "SFK15", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["スマート回胴", "SFK20", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["ぱちんこ", "PF40", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["スマートぱちんこ", "PF45", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["スマートぱちんこ", "PFK20", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["スマートぱちんこ", "PFK25", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["スマートぱちんこ", "PFK30", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["アンプ基板", "S2501C", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["ジャック基板", "S2502C", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]];

var MODEL_BOARD_COMPOSITION_DATA = [["e:C44", "M2105C", "C2001E", "D2101A", "E2101B", "DE2103A"], ["S:NA07", "M2203A", "", "D1401C", "E1902B", "DE1802A"], ["P:A77", "M2003A1", "C1601E", "D2101A", "E2002B", "DE2103A"], ["P:A76ｻﾌﾞ", "M2003A3", "C1601E", "D2101A", "E1901B", "DE1607A"], ["P:A76甘", "M2003A3", "C1601E", "D2101A", "E1901B", "DE1607A"], ["P+e:A79", "M2003A2", "C1601E", "D1401C", "E1901B", "DE2103A"], ["L:E46", "M2104A1", "C2101B", "D1401C", "E2102B", "DE1802A"], ["P:D57", "M2003A5", "C1601E", "D2101A1", "E2002B", "DE2103A"], ["P:A80", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2103A"], ["P:D55ｻﾌﾞ", "M2003A3", "C1601E", "D2101A", "E2002B", "DE1802A"], ["P:A78ｻﾌﾞ", "M2003A12", "C1601E", "D2101A1", "E2002B", "DE2103A"], ["e:D56", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2101A"], ["P:C45", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2103A"], ["P:A77ｻﾌﾞ", "M2003A1", "C1601E", "D2101A", "E2002B", "DE2103A"], ["P+e:A81", "M2105C2", "C1601E", "D2101A", "E2101B", "DE2101A"], ["L:E47", "M2104A", "C2101B", "D1401C", "E2102B", "DE1802A"], ["P+e:A77ｻﾌﾞ", "M2003A1/M2105C", "C1601E", "D2101A", "E2002B", "DE2103A"], ["P+e:D58", "", "", "D2101A", "E2102B", "DE2101A"], ["P:A81", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2101A"], ["e:A82", "M2105C2", "C2001E", "D2101A", "E2101B", "DE2101A"], ["e:A83", "M2105C3", "C2001E", "D2101A", "E2101B", "DE2101A1"], ["e:D59", "M2401A(M2105C3)", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["P+e:A77甘", "M2003A1", "", "D2101A", "E2002B", "DE2103A"], ["P+e:A79ｻﾌﾞ", "M2003A2/M2105C1", "", "D1401C", "E1901B/E2101B", "DE2103A"], ["P:D57ｻﾌﾞ", "M2003A5", "C1601E", "D2101A1", "E2002B", "DE2103A1"], ["L:E48", "M2104A", "C2101B", "D1401C", "E2102B", "DE1802A"], ["P+e:A80ｻﾌﾞ", "", "", "D2101A1", "E2002B", "DE2101A/2103A"], ["P:D56ｻﾌﾞ", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2101A"], ["L:E49", "M2402A(M2104A2)", "C2101B", "D1401C", "E2102B", "DE1802A"], ["e:A84", "M2401A(M2105C3)", "C2401A(C2001E)", "SNB52163A", "E2301B", "なし"], ["e:A60", "M2105C2", "C2001E", "", "E2301A", ""], ["e:A81ｻﾌﾞ", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2101A"], ["L:E60", "M2402A(M2104A2)", "C2502A(C2101B)", "D2101A", "E2102B", "DE1802A"], ["e:D60", "M2401A(M2105C3)", "C2401A(C2001E)", "SNB52163A", "E2301B", "なし"], ["e:A77ｻﾌﾞ", "M2105C", "C2001E", "D2101A", "E2101B", "DE2103A"], ["P:A79", "M2003A2", "C1601E", "D1401C", "E1901B", "DE2103A"], ["e:A79", "M2105C1", "C2001E", "D1401C", "E2101B", "DE2103A"], ["P:D58", "M2202A", "C2202B", "D2101A", "E2002B", "DE2101A"], ["e:D58", "M2105C2", "C2001B", "D2101A", "E2101B", "DE2101A"], ["P:A77甘", "M2003A1", "C1601E", "D2101A", "E2002B", "DE2101A"], ["e:A77甘", "M2003A1", "C2001E", "D2101A", "E2002B", "DE2103A"], ["P:A80ｻﾌﾞ", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2103A1"], ["e:A80ｻﾌﾞ", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2101A"], ["P:A79ｻﾌﾞ", "M2003A2", "C1601E", "D1401C", "E1901B", "DE2103A"], ["e:A79ｻﾌﾞ", "M2105C1", "C2001E", "D1401C", "E2101B", "DE2103A"], ["P:D57B", "M2003A5", "C1601E", "D2101A1", "E2002B", "DE2103A"], ["P:D56", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2101A"], ["P:A84", "三球M2421", "C1601E", "AXB52163A", "E2401A", "なし"], ["P:A83", "M2003A", "C1601E", "D2101A", "E2002B", "DE2101A1"], ["P:A82", "M2003A", "C1601E", "D2101A", "E2002B", "DE2101A"], ["e:A82ｻﾌﾞ", "M2202A", "C2202B", "D2101A", "E2002B", "DE2101A"], ["P:D58ｻﾌﾞ", "M2202A", "C2202B", "D2101A", "E2002B", "DE2101A"], ["P:D60", "三球M2421", "C1601E", "AXB52163A", "E2401A", "なし"], ["L:E61", "M2402A(M2104A2)", "C2101B", "D1401C", "E2102B", "DE1802A"], ["P:A83ｻﾌﾞ", "M2003A", "C1601E", "D2101A", "E2002B", "DE2101A1"], ["e:A83ｻﾌﾞ", "M2105C3", "C2001E", "D2101A", "E2101B", "DE2101A1"], ["P:A85", "三球", "C1601E", "D2101A", "E2002B", "DE2101A1"], ["e:A85", "M2401A(M2105C3)", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["e:D59ｻﾌﾞ", "M2401A(M2105C3)", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["P:A84サブ", "三球", "C1601E", "AG6R", "E2401A", "なし"], ["e:A84サブ", "M2105C3", "C2001E", "AG6R", "E2301B", "なし"], ["e:A86", "三球", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["L:E62", "M2402A(M2104A2)", "C2101B", "D1401C", "E2102B", "DE1802A"], ["P:D60ｻﾌﾞ", "M2003A6", "C1601E", "AG6R", "E2401A", "なし"], ["e:D60ｻﾌﾞ", "M2401A(M2105C3)", "C2401A(C2001E)", "AG6R", "E2301B", "なし"], ["P+e:A87？", "M2003A6", "C1601E", "AG6R", "E2301A", "なし"], ["P:A79甘", "M2003A2", "C1601E", "D1401C", "E1901B", "DE2103A1"], ["P+e:A87", "M2003A6", "C1601E", "AG6R", "E2301A", "なし"], ["P:D57甘", "M2003A5", "C1601E", "D2101A1", "E2002B", "DE2103A1"], ["P:A81ｻﾌﾞ", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2101A"], ["P:A82ｻﾌﾞ", "M2003A", "C1601E", "D2101A", "E2002B", "DE2101A"], ["P:D57再", "M2003A5", "C1601E", "D2101A1", "E2002B", "DE2103A"], ["P:A80甘", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2103A1"], ["P:D56甘", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2101A"], ["P:A81甘", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2101A"], ["P:A82甘", "M2003A", "C1601E", "D2101A", "E2002B", "DE2101A"], ["P:C45(追加)", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2103A"], ["P:C45(再追加)", "M2003A5", "C1601E", "D2101A", "E2002B", "DE2103A"], ["e:A87", "M2503A", "C2501C", "SNB52163A", "E2501B", "なし"], ["e:D61", "三球(FJ+M2422B)", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["e:A83甘", "M2105C3", "C2001E", "D2101A", "E2101B", "DE2101A1"], ["e:A84ｻﾌﾞ", "M2401A(M2105C3)", "C2401A(C2001E)", "SNB52163A", "E2301B", "なし"], ["P:A83甘", "M2003A", "C1601E", "D2101A", "E2002B", "DE2101A1"], ["L:E63", "M2402A(M2104A2)", "C2502A(C2101B)", "D1401C", "E2102B", "DE1802A"], ["L:E64", "M2402A(M2104A2)", "C2502A(C2101B)", "D2101A", "E2102B", "DE1802A"], ["e:D58ｻﾌﾞ", "M2401A(M2105C3)", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A"], ["e:/P:C46", "M2401A(M2105C3)", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["e:C46", "M2401A(M2105C3)", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["e:D60甘", "M2401A(M2105C3)", "C2401A(C2001E)", "SNB52163A", "E2301B", "なし"], ["e:A88", "M2503A", "C2501C", "D2101A", "E2503B", "DE2502A"], ["e:A85甘", "M2401A(M2105C3)", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["e:A86甘", "三球", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["e:D61甘", "三球", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["e:A87ｻﾌﾞ", "M2503A", "C2501C", "SNB52163A", "E2501B", "なし"], ["e:D58B", "M2401A(M2105C3)", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A"], ["e:D62", "M2503A", "C2501C", "D2101A", "E2503B", "DE2502A"], ["e:D63", "M2503A", "C2501C", "SNB52163A", "E2501B", "なし"], ["L:E66", "M2402A(M2104A2)", "C2502A(C2101B)", "D2101A", "E2102B", "DE1802A"], ["L:※", "三球", "三球", "", "", ""], ["L:E67", "M2402A(M2104A2)", "C2502A(C2101B)", "D2101A", "E2102B", "DE1802A"], ["L:E68", "三球", "三球", "", "", ""], ["e:A89", "M2503A", "C2501C", "D2101A", "E2503B", "DE2502A"], ["L:E65", "M2402A(M2104A2)", "C2502A(C2101B)", "D2101A", "E2102B", "DE1802A"], ["e:A84増台", "M2401A(M2105C3)", "C2401A(C2001E)", "SNB52163A", "E2301B", "なし"], ["e:C47", "M2401A(M2105C3)", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["P:C46B", "M2003A", "C1601E", "D2101A", "E2002B", "DE2101A1"], ["e:A84甘", "M2401A(M2105C3)", "C2401A(C2001E)", "SNB52163A", "E2301B", "なし"], ["e:A84B", "M2401A(M2105C3)", "C2401A(C2001E)", "SNB52163A", "E2301B", "なし"], ["e:A84C", "M2401A(M2105C3)", "C2401A(C2001E)", "SNB52163A", "E2301B", "なし"], ["e:A84再再版", "M2401A(M2105C3)", "C2401A(C2001E)", "SNB52163A", "E2301B", "なし"], ["e:D59B", "M2401A(M2105C3)", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["e:A84再再再版", "M2401A(M2105C3)", "C2401A(C2001E)", "SNB52163A", "E2301B", "なし"], ["C46再販", "M2401A(M2105C3)", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["e:A84再々再々", "M2401A(M2105C3)", "C2401A(C2001E)", "SNB52163A", "E2301B", "なし"], ["e:A84再販×5", "M2401A(M2105C3)", "C2401A(C2001E)", "SNB52163A", "E2301B", "なし"], ["P:C46C", "M2003A", "C1601E", "D2101A", "E2002B", "DE2101A1"], ["e:C46C", "M2401A", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"], ["e:A85B", "M2401A(M2105C3)", "C2401A(C2001E)", "D2101A", "E2101B", "DE2101A1"]];

function _initMasterBoardPriceSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(MASTER_BOARD_PRICE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(MASTER_BOARD_PRICE_SHEET);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 130);
  }
  return sheet;
}

function _initModelBoardCompositionSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(MODEL_BOARD_COMPOSITION_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(MODEL_BOARD_COMPOSITION_SHEET);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 140);
  }
  return sheet;
}

// Excelマスタ表の初期データで両シートを（再）投入する。既存の手入力内容は上書きされる。
function apiSeedMasterPriceData() {
  try {
    var priceSheet = _initMasterBoardPriceSheet();
    priceSheet.clear();
    priceSheet.getRange(1, 1, 1, MASTER_BOARD_PRICE_HEADERS.length).setValues([MASTER_BOARD_PRICE_HEADERS]);
    priceSheet.getRange(1, 1, 1, MASTER_BOARD_PRICE_HEADERS.length).setBackground('#E1F5FE').setFontWeight('bold');
    priceSheet.setFrozenRows(1);
    if (MASTER_BOARD_PRICE_DATA.length) {
      priceSheet.getRange(2, 1, MASTER_BOARD_PRICE_DATA.length, MASTER_BOARD_PRICE_HEADERS.length).setValues(MASTER_BOARD_PRICE_DATA);
    }

    var modelSheet = _initModelBoardCompositionSheet();
    modelSheet.clear();
    modelSheet.getRange(1, 1, 1, MODEL_BOARD_COMPOSITION_HEADERS.length).setValues([MODEL_BOARD_COMPOSITION_HEADERS]);
    modelSheet.getRange(1, 1, 1, MODEL_BOARD_COMPOSITION_HEADERS.length).setBackground('#E8F5E9').setFontWeight('bold');
    modelSheet.setFrozenRows(1);
    if (MODEL_BOARD_COMPOSITION_DATA.length) {
      modelSheet.getRange(2, 1, MODEL_BOARD_COMPOSITION_DATA.length, MODEL_BOARD_COMPOSITION_HEADERS.length).setValues(MODEL_BOARD_COMPOSITION_DATA);
    }

    return { success: true, priceRows: MASTER_BOARD_PRICE_DATA.length, modelRows: MODEL_BOARD_COMPOSITION_DATA.length };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── 基板PCB価格マスタ ──
function apiMasterBoardPriceGet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(MASTER_BOARD_PRICE_SHEET);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, headers: MASTER_BOARD_PRICE_HEADERS, rows: [] };
    var last = sheet.getLastRow();
    var data = sheet.getRange(2, 1, last - 1, MASTER_BOARD_PRICE_HEADERS.length).getValues();
    var rows = data
      .filter(function(r) { return r.some(function(v) { return v !== '' && v !== null; }); })
      .map(function(r, i) {
        var obj = { _row: i + 2 };
        MASTER_BOARD_PRICE_HEADERS.forEach(function(h, ci) { obj[h] = r[ci]; });
        return obj;
      });
    return { success: true, headers: MASTER_BOARD_PRICE_HEADERS, rows: rows };
  } catch (e) { return { success: false, error: e.message }; }
}

// 1セル更新（行番号 _row とヘッダー名を指定）
function apiMasterBoardPriceUpdateCell(payload) {
  try {
    payload = payload || {};
    var row = Number(payload.row || 0);
    var header = String(payload.header || '');
    var colIdx = MASTER_BOARD_PRICE_HEADERS.indexOf(header);
    if (!row || colIdx < 0) return { success: false, error: '不正なパラメータです' };
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(MASTER_BOARD_PRICE_SHEET);
    if (!sheet) return { success: false, error: 'シートがありません' };
    sheet.getRange(row, colIdx + 1).setValue(payload.value);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── 機種基板構成 ──
function apiModelBoardCompositionGet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(MODEL_BOARD_COMPOSITION_SHEET);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, headers: MODEL_BOARD_COMPOSITION_HEADERS, rows: [] };
    var last = sheet.getLastRow();
    var data = sheet.getRange(2, 1, last - 1, MODEL_BOARD_COMPOSITION_HEADERS.length).getValues();
    var rows = data
      .filter(function(r) { return r.some(function(v) { return v !== '' && v !== null; }); })
      .map(function(r, i) {
        var obj = { _row: i + 2 };
        MODEL_BOARD_COMPOSITION_HEADERS.forEach(function(h, ci) { obj[h] = r[ci]; });
        return obj;
      });
    return { success: true, headers: MODEL_BOARD_COMPOSITION_HEADERS, rows: rows };
  } catch (e) { return { success: false, error: e.message }; }
}

function apiModelBoardCompositionUpdateCell(payload) {
  try {
    payload = payload || {};
    var row = Number(payload.row || 0);
    var header = String(payload.header || '');
    var colIdx = MODEL_BOARD_COMPOSITION_HEADERS.indexOf(header);
    if (!row || colIdx < 0) return { success: false, error: '不正なパラメータです' };
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(MODEL_BOARD_COMPOSITION_SHEET);
    if (!sheet) return { success: false, error: 'シートがありません' };
    sheet.getRange(row, colIdx + 1).setValue(payload.value);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function apiModelBoardCompositionAddRow(payload) {
  try {
    var code = String((payload || {}).code || '').trim();
    if (!code) return { success: false, error: '機種コードが必要です' };
    var sheet = _initModelBoardCompositionSheet();
    if (sheet.getLastRow() < 1) {
      sheet.getRange(1, 1, 1, MODEL_BOARD_COMPOSITION_HEADERS.length).setValues([MODEL_BOARD_COMPOSITION_HEADERS]);
    }
    sheet.appendRow([code, '', '', '', '', '']);
    return { success: true, row: sheet.getLastRow() };
  } catch (e) { return { success: false, error: e.message }; }
}
