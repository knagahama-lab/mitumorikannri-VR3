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
// ヘッダー列の後ろに付く内部管理列（表には出さない）
var MASTER_BOARD_PRICE_EXTROW_COL = MASTER_BOARD_PRICE_HEADERS.length + 1; // 元Excel/外部シート上の行番号
var MASTER_BOARD_PRICE_LINKS_COL  = MASTER_BOARD_PRICE_HEADERS.length + 2; // セルごとのリンクURL（JSON）

// 元の値に張られていたハイパーリンク（PCB原価・実装費・PCB公表単価・仕掛公表単価など）。
// 行のインデックスは MASTER_BOARD_PRICE_DATA / MASTER_BOARD_PRICE_SOURCE_ROWS と対応する。
var MASTER_BOARD_PRICE_SOURCE_ROWS = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 49, 50, 51, 52, 54, 55, 56, 57, 58, 60, 61];
var MASTER_BOARD_PRICE_LINKS = ["{\"PCB原価\": \"https://drive.google.com/file/d/11b3-p496S630sllQWL1rsgJjF5C3MQBC/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1Mz1nSz1flqkQxJxcCmoZJErmME2q9Qil/view?usp=drive_link\", \"部品表(K10)\": \"https://docs.google.com/spreadsheets/d/1TVhi9zkS3nwL1oulh1HSKsGHRQ96uPM7/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1uvZXgS6Bey-yNhERG5K8N9e3p8w7ECoo/view?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1rCOuH40rdWhFAvhvPcDQJghQsRt2KyM3/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1y85rPpE9hpbpR_EC19svIrNrVaLnl_kL/view?usp=drive_link\", \"部品表(K10)\": \"https://docs.google.com/spreadsheets/d/1ctIbjZMET6Do88Td8OGQpWZ6kVomwGI0/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1nPt597KNmKgMQqYlzbKBo-07nmhNkcdW/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/15am5Y3mZjpWBb-W1S5r9BgOE9cbDZoh9/view?usp=drive_link\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\", \"ブランド\": \"https://drive.google.com/open?id=1KDAp2vY4ME7ySLQJtDRl1P0NzeJ263dz&usp=drive_copy\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1Es9fC3NA0soY4yiGiiGzAYHmmy6quLIGVRa_eiaJD-k/edit?gid=1607512910\"}", "{}", "{\"PCB原価\": \"https://drive.google.com/file/d/1aktZfgBRBGGOdkBAPJVqRoFsj0mH001l/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1mFC9uMNEf_qKZf3FX9kTBsd0tzxOPhKD/view?usp=drive_link\", \"部品表(K10)\": \"https://docs.google.com/spreadsheets/d/1oDPm1mQY_ED3ViKtwOpowTxO_CZeE55P/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\", \"PCB公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1v1oQkGwJVSjHrM3tD6UpS9gJDxw5UDPv5ruHSFyy_EE/edit?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1kfht4V95itkOsrZv-ZGT3kYHtQ7GuAim/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1P5xQSNWfG02xcXeFj-MxyCMKFK7i5wEV/view?usp=drive_link\", \"部品表(K10)\": \"https://docs.google.com/spreadsheets/d/1zhxCIRYEkUP0bC_CFw-NKd7QgV3iZ2Fe/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\", \"PCB公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"ブランド\": \"https://drive.google.com/open?id=1d76-tQ9xY_X1LJ0o4ziv3sRP3nsLxwBz&usp=drive_copy\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1TVZeGR0ZH7prggd2IxwQHyrUjXy9vOtpH0P93Is_EdM/edit?usp=drive_link\"}", "{\"実装費\": \"https://drive.google.com/file/d/1H2TeXo3MNL_h4Naws0iuS9LBCnxGgasA/view?usp=drive_link\", \"PCB公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1BldU893DlXRwzUbg-759I0fqt7V_KouFb5boAAhQFtc/edit?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/15fYj2DNt3gfHgxLk9DSIwgN2BJXYlA92/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1fO_XQN1YsD7Om8zLw4XgrauU-OlkIAAr/view?usp=drive_link\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1WImxw_o2uh7B-MINjoTm7h3juu5rHiTu/view?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1QIOrV807nls9kA7e-7YIujyx2wgBN4m4/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1-AxtxwWhMeTF44UnLRASYyBHhxzUZto8/view?usp=drive_link\", \"PCB公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\"}", "{}", "{\"ブランド\": \"https://drive.google.com/open?id=1ak3syEToHliY4S9nFbY_kofpgYUIiP01&usp=drive_copy\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1jHIs4rKjcokUTqp6B2TK8XQpgHKj9lTl96Ae4P29mTA/edit?usp=sharing\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1wQdjwgWnvKLn6sSfc6mTkYOu3EiSakDv/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1gGgQsF05y1ognloVtoIomqfrPU35fk10/view?usp=drive_link\", \"PCB公表単価\": \"https://drive.google.com/file/d/1KXqqYnMn4Cj6SUazCjNZpO_BBjcNmXHQ/view?usp=drive_link\", \"仕掛公表単価\": \"https://drive.google.com/file/d/10yT03Rc4qm9h5duXglcOiybpQmoBbh7d/view?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1dsceNf9zVm8wenrV-knVUfzhaMff7CM7/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1T0_jzGe2ZhBdeMZ8DoN2peX5Pj2Ud849/view?usp=drive_link\", \"部品表(K10)\": \"https://docs.google.com/spreadsheets/d/1YJP1AQVo1zRs6lJlah4wRPPEJas0KcuU/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\", \"PCB公表単価\": \"https://drive.google.com/file/d/1G32dimsGMItFCM5PLpUZ-PZURGDlmHW1/view?usp=drive_link\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1YR1Fk1AzjX7w3KKMDxBpWPjDzhwO9ulY/view?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1FaovLmkbXEvg4TW9Su4sZkQb52Xmm9Q1/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/13sgg_G27oijvddUSqbkbqsiPHT8TQMsd/view?usp=drive_link\", \"部品表(K10)\": \"https://docs.google.com/spreadsheets/d/15t2pNezczKJ41GXqCMgA8oV1z0w60iuV/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1Eexd1mQ4GWbXjTbnwueKeHE48_Xht4U5/view?usp=drive_link\"}", "{}", "{\"PCB原価\": \"https://drive.google.com/file/d/13XPtL8TfkkmyjffmtvVo_11AUgj3RCz_/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1YDug0D3ygbZqju45_FCVVfgi0G4dzTKS/view?usp=drive_link\", \"部品表(K10)\": \"https://docs.google.com/spreadsheets/d/1vV5n5l67g5XTLzlDkzkj-EQUYGXeNUDd/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1zrAirdDKR79gMGPhDkmB1Zge_e_dMhRZ/view?usp=sharing\", \"過去見積リンク1\": \"https://drive.google.com/file/d/1RKlA-tDzJFdMvjdw5FLmVM3JpXpqEekk/view?usp=drive_link\", \"過去見積リンク2\": \"https://drive.google.com/file/d/1h_0h6fpsggcxUSel7vhXCUF2wpAthHps/view?usp=drive_link\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1J1up4vF6RgiNwHBii6JZBTCPQcP1RRoMNoqnetNjAF0/edit\"}", "{\"設計課情報\": \"https://docs.google.com/spreadsheets/d/1nTIsjMEFaoV9ymukW73wtOAzO_gLybb1zU3RccawkLM/edit?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1LEhylDbH8zO45QR5aB-x64A5BuK8FJNM/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1vPNFG2wKHJgpddjNKucDGfl6a3TGI-vI/view?usp=drive_link\", \"部品表(K10)\": \"https://docs.google.com/spreadsheets/d/1zPF5_cCLnYwfW1sH2OHt99a-ASHLQirC/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\", \"過去見積リンク1\": \"https://drive.google.com/file/d/1xxTQsOsj1KSF6Obld0Z_GzeqiNcX-Fk8/view?usp=drive_link\", \"過去見積リンク2\": \"https://drive.google.com/file/d/17SaNySunlEbcty8khL8R-Ig_QxK1yf5H/view?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1zWoSYYNQq5XI9eMomTBRfbeIhx72ZSMQ/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1eg-ns8KB37mS5fEDgZtKDtIiGgxwkcqh/view?usp=drive_link\", \"部品表(K10)\": \"https://docs.google.com/spreadsheets/d/1xpWH3TvPO1BjauLdr_t76lNa6FSQ-ctu/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\", \"過去見積リンク1\": \"https://drive.google.com/file/d/1xxTQsOsj1KSF6Obld0Z_GzeqiNcX-Fk8/view?usp=drive_link\"}", "{}", "{\"PCB原価\": \"https://drive.google.com/file/d/14zIB4jPqCdgbScHIkzqOcmW-qRRqnk_d/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1Dhx2tPtQzaPvtU3LAscgPj3H2V4f2wXI/view?usp=drive_link\", \"部品表(K10)\": \"https://docs.google.com/spreadsheets/d/1vw1TBQpzDQUQPmTkwA_ffJUrY6DWANpH/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\", \"仕掛公表単価\": \"https://drive.google.com/file/d/17_QvdIKi3uG0uA6Pi0zgI61NVl-N53F2/view?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1ULqYmRBiab7Vftmz_ILVPGcRwgNmfNtr/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1kvkHSsC8meZsPf2gEf_FfpvCGoPN8VQB/view?usp=drive_link\", \"部品表(K10)\": \"https://docs.google.com/spreadsheets/d/1jHbzfcKOoTqJAWfZ6LLQh9IMQgP-l8__/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1a6r5rmnEPwctg8cw3VGXcqn4uWozTYJx/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1p0UwrV2Gjg2cI5GTT366Lm0R9iD-iKX2/view?usp=drive_link\", \"部品表(K10)\": \"https://docs.google.com/spreadsheets/d/19r8lLSl5RqMRjlhhEsnbiQN7IaVKUgvq/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1oaVjQZaRU9hCF48BD2Hyji29mbUBv9OY/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1nnFHKl_48wQ9onga4K-A0-LJdflJ8jGz/view?usp=drive_link\", \"PCB公表単価\": \"https://drive.google.com/file/d/1r9Ftnt4WyC7JgAkqz8pUju0c0-sP_My-/view?usp=drive_link\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1qlTwXtkQONeJXfKZEmL-a6YpVdyRKlVz/view?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/117exYeAoGwIwIie97S54_eoQTglxYaVd/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1Qqb9R3n123qadCfArEmhDutW3671UWbY/view?usp=drive_link\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\", \"ブランド\": \"https://drive.google.com/open?id=1SLVpGbI1Wl8xZEiHbzyTlhksIblMKLT-&usp=drive_copy\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1_GiYUmBAH7jMIydMcDC4-jUbk4Kf5omNHwezFRM-rtM/edit?usp=sharing\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1V4v8im0vQoXJvPr5cXVm3W-20eK8uQJK/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1iNcm2W4ioiuO3Y7CZvw72Q6yG_HHL2yh/view?usp=drive_link\", \"PCB公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"ブランド\": \"https://drive.google.com/open?id=1C_eUNoQdl2T_-Ph9E-29jmJqFGgzzzvS&usp=drive_copy\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1V3FFrs-FvskrQlBPTINaaRvU6SBJf310V0gRD6zHdMM/edit?usp=sharing\"}", "{\"ブランド\": \"https://drive.google.com/open?id=1fATcvVULFz09TDOLX6JAfC3cA12wbHAf&usp=drive_copy\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1ahX4yavzLwEOsvaTHujLklz2XDynGnZ9_EV0Iy9EElA/edit\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1U7ox1RjVWuHwsXOSdzFjd6Z9oVb6g6ra/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1zfTrUyPQiSCd6kkvsTTu3qyOgRB32Ejk/view?usp=drive_link\", \"PCB公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"ブランド\": \"https://drive.google.com/open?id=1Nsk0yDrZ0Y6v6qck7R0WGvPBp27CBpGr&usp=drive_copy\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1E-5iyLQ0ciHN1KWj24MEcrmStrjaIcvLgylKw6MkoaU/edit\"}", "{\"ブランド\": \"https://drive.google.com/open?id=1Zpnz_dA23Ytx5hHLTLY4-l_nzB_gAYCK&usp=drive_copy\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1oc9OhjGk8Y_Eory_Rjpbv5NqgZPsno0oT-8mZOzrJ0M/edit?usp=drive_link\"}", "{\"設計課情報\": \"https://docs.google.com/spreadsheets/d/1d50PCBb1tqiLFNT2zS_QMBktaXbF9RgwNHpd_gRn0J8/edit?gid=1607512910\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1gxyOQ7r3sA-fFqoY0kX9MGGRU-51ds63/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1QyJwCusb4JpZnA1axycSPm_EoRbI--m9/view?usp=drive_link\", \"PCB公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"ブランド\": \"https://drive.google.com/open?id=1Zpnz_dA23Ytx5hHLTLY4-l_nzB_gAYCK&usp=drive_copy\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1wLvTHiAlv8S49vywYjB38V7m-tzssGNg/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1XGPLeAc_-etEnczNNsw0HjyS2ZBtpLYD/view?usp=drive_link\", \"PCB公表単価\": \"https://drive.google.com/file/d/1Yci0PDwaKkU-KGKS4KMBHWmJzzcx-cSP/view?usp=drive_link\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1F8wCDjvMiRUX6wEXbCDG1vOkYb3T6NlH/view?usp=drive_link\"}", "{}", "{\"ブランド\": \"https://drive.google.com/open?id=1ZgpxtAE9XBD1HIPBbjrQriLeH4tdHXsa&usp=drive_copy\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1IJCf5xVXLY8rlxbwDe9Vnw5-Ijz2cojR/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1BCO1u9mJvlvDEj_tdsM8aZld19JXvXFn/view?usp=drive_link\", \"PCB公表単価\": \"https://drive.google.com/file/d/1r9Ftnt4WyC7JgAkqz8pUju0c0-sP_My-/view?usp=drive_link\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1qlTwXtkQONeJXfKZEmL-a6YpVdyRKlVz/view?usp=drive_link\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1LCfXUQddHT1rB5ijDctqOu1NIelG7-sH/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1P4lBSHPR51GhSx9N6-4-WE4bYLkYAP47/view?usp=drive_link\", \"PCB公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\"}", "{\"ブランド\": \"https://drive.google.com/open?id=1ciWl8sRL5TcUsXkA0mx-r2VlzKoUu7cN&usp=drive_copy\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1lhz181exIl7Tynhn7c3QK0vTP5vdu9rof7hUWzTq6vc/edit\"}", "{\"PCB原価\": \"https://drive.google.com/file/d/1mJLfOXS3QDYMRa5XTFZYT8tYM4r99mZ6/view?usp=drive_link\", \"実装費\": \"https://drive.google.com/file/d/1g763qKIlEAfOG0b8YUTgcGfN__gq4EwG/view?usp=drive_link\", \"PCB公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1k6SkKeJy6TcYcYm0kkD2MqN0TnK749JwRDlFPGs-BCw/edit?gid=1607512910\"}", "{\"PCB公表単価\": \"https://drive.google.com/file/d/1jkcA4Yf3YRFaomKLaGnrJ7iKNfh4E9ds/view?usp=drive_link\", \"仕掛公表単価\": \"https://drive.google.com/file/d/1DcaQt2h94U58VzbUohZe14js9Ho5sTr-/view?usp=drive_link\", \"設計課情報\": \"https://docs.google.com/spreadsheets/d/1rRZjy2-3OFre2n6xHOABftc6ZN7mAq9KLv507igX1F4/edit?gid=1607512910\"}", "{\"設計課情報\": \"https://docs.google.com/spreadsheets/d/1Iu0MTSOaFSRPRnHjVTFCKJjV2EmNuTs3qP86j_OASOU/edit?usp=drive_link\"}", "{\"実装費\": \"https://drive.google.com/file/d/1lxty11pjp_7TkTUsZEFc4P2DxQVkjbEw/view?usp=drive_link\", \"公表単価合計\": \"https://drive.google.com/file/d/1zrAirdDKR79gMGPhDkmB1Zge_e_dMhRZ/view?usp=sharing\"}", "{}", "{}", "{}", "{}", "{}", "{}", "{}", "{}", "{}", "{}", "{}", "{}", "{}"];

// 外部連携先（元Excelの実体であるGoogleスプレッドシート）。編集内容はここへも書き戻す。
// URL: https://docs.google.com/spreadsheets/d/1oQs-gg_vQ7oauwxMOM_7DcPLue7lT0rfwn-6J6HfREk/edit?gid=902539252
var MASTER_PRICE_EXTERNAL_SHEET_ID = '1oQs-gg_vQ7oauwxMOM_7DcPLue7lT0rfwn-6J6HfREk';
var MASTER_PRICE_EXTERNAL_GID = 902539252;

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
    var fullHeaders = MASTER_BOARD_PRICE_HEADERS.concat(['外部行番号', 'リンクJSON']);
    priceSheet.getRange(1, 1, 1, fullHeaders.length).setValues([fullHeaders]);
    priceSheet.getRange(1, 1, 1, fullHeaders.length).setBackground('#E1F5FE').setFontWeight('bold');
    priceSheet.setFrozenRows(1);
    if (MASTER_BOARD_PRICE_DATA.length) {
      priceSheet.getRange(2, 1, MASTER_BOARD_PRICE_DATA.length, MASTER_BOARD_PRICE_HEADERS.length).setValues(MASTER_BOARD_PRICE_DATA);
      var extraCols = MASTER_BOARD_PRICE_DATA.map(function(_, i) {
        return [MASTER_BOARD_PRICE_SOURCE_ROWS[i] || '', MASTER_BOARD_PRICE_LINKS[i] || '{}'];
      });
      priceSheet.getRange(2, MASTER_BOARD_PRICE_EXTROW_COL, extraCols.length, 2).setValues(extraCols);
      // ハイパーリンクが張られていた列は、リンクを保持したままセルに反映する
      MASTER_BOARD_PRICE_DATA.forEach(function(rowVals, i) {
        var links = {};
        try { links = JSON.parse(MASTER_BOARD_PRICE_LINKS[i] || '{}'); } catch (pe) {}
        Object.keys(links).forEach(function(header) {
          var colIdx = MASTER_BOARD_PRICE_HEADERS.indexOf(header);
          if (colIdx < 0) return;
          _setCellValuePreservingLink(priceSheet, i + 2, colIdx + 1, rowVals[colIdx], links[header]);
        });
      });
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
// 通常は金額確認用（読み取り専用）。修正は「詳細」→「修正」の操作からのみ行う想定。
function apiMasterBoardPriceGet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(MASTER_BOARD_PRICE_SHEET);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, headers: MASTER_BOARD_PRICE_HEADERS, rows: [] };
    var last = sheet.getLastRow();
    var data = sheet.getRange(2, 1, last - 1, MASTER_BOARD_PRICE_LINKS_COL).getValues();
    var rows = data
      .filter(function(r) { return r.some(function(v) { return v !== '' && v !== null; }); })
      .map(function(r, i) {
        var obj = { _row: i + 2, _extRow: r[MASTER_BOARD_PRICE_EXTROW_COL - 1] || '', _links: {} };
        try { obj._links = JSON.parse(r[MASTER_BOARD_PRICE_LINKS_COL - 1] || '{}'); } catch (pe) {}
        MASTER_BOARD_PRICE_HEADERS.forEach(function(h, ci) { obj[h] = r[ci]; });
        return obj;
      });
    return { success: true, headers: MASTER_BOARD_PRICE_HEADERS, rows: rows };
  } catch (e) { return { success: false, error: e.message }; }
}

// 値を保ったままハイパーリンクだけ張り直す（リンクが無ければ通常のsetValue）
function _setCellValuePreservingLink(sheet, row, col, value, linkUrl) {
  var range = sheet.getRange(row, col);
  if (linkUrl && value !== '' && value !== null && value !== undefined) {
    var rtv = SpreadsheetApp.newRichTextValue().setText(String(value)).setLinkUrl(linkUrl).build();
    range.setRichTextValue(rtv);
  } else {
    range.setValue(value);
  }
}

// 元Excelの実体である外部スプレッドシートへ同一セルを書き戻す（失敗しても呼び出し元の保存は成功扱いとする）
function _syncMasterBoardPriceToExternal(extRow, header, value, linkUrl) {
  var colIdx = MASTER_BOARD_PRICE_HEADERS.indexOf(header);
  if (colIdx < 0 || !extRow) return false;
  var extSs = SpreadsheetApp.openById(MASTER_PRICE_EXTERNAL_SHEET_ID);
  var sheets = extSs.getSheets();
  var extSheet = null;
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === MASTER_PRICE_EXTERNAL_GID) { extSheet = sheets[i]; break; }
  }
  if (!extSheet) throw new Error('外部シート（gid=' + MASTER_PRICE_EXTERNAL_GID + '）が見つかりません');
  _setCellValuePreservingLink(extSheet, extRow, colIdx + 1, value, linkUrl);
  return true;
}

// 詳細モーダルの「修正」からの一括保存。ローカルシートを更新し、可能であれば外部シートへも同期する。
// payload: { row: ローカル行番号, updates: { ヘッダー名: 値, ... } }
function apiMasterBoardPriceSaveRow(payload) {
  try {
    payload = payload || {};
    var row = Number(payload.row || 0);
    var updates = payload.updates || {};
    if (!row) return { success: false, error: '行が特定できません' };
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(MASTER_BOARD_PRICE_SHEET);
    if (!sheet) return { success: false, error: 'シートがありません' };

    var extRow = Number(sheet.getRange(row, MASTER_BOARD_PRICE_EXTROW_COL).getValue() || 0);
    var links = {};
    try { links = JSON.parse(sheet.getRange(row, MASTER_BOARD_PRICE_LINKS_COL).getValue() || '{}'); } catch (pe) {}

    var externalSynced = false, externalError = '';
    Object.keys(updates).forEach(function(header) {
      var colIdx = MASTER_BOARD_PRICE_HEADERS.indexOf(header);
      if (colIdx < 0) return;
      var value = updates[header];
      var linkUrl = links[header] || null;
      _setCellValuePreservingLink(sheet, row, colIdx + 1, value, linkUrl);
      if (extRow) {
        try {
          _syncMasterBoardPriceToExternal(extRow, header, value, linkUrl);
          externalSynced = true;
        } catch (se) { externalError = se.message; }
      }
    });

    return { success: true, externalRow: extRow || null, externalSynced: externalSynced, externalError: externalError };
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
