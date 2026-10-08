# 住宅着工 景況マップ

国土交通省「建築着工統計調査」（住宅着工統計・建築物着工統計）を、年・月・都道府県別に前年同期比や構成比で集計し、表と都道府県色分け地図で表示するアプリです。

- 公開URL: https://heqiguangda-coder.github.io/housing-starts-map/
- 初期表示は操作確認用の架空データです
- バージョン：`VERSION` に記載（アプリ名の横と「データ」タブに表示）。変更内容は [CHANGELOG.md](CHANGELOG.md)、過去の版は `releases/housing-starts-map_ver.N.N.N.html`

## データ

| 取り込み方 | 内容 |
|---|---|
| 自動更新データ（おすすめ） | `data/` に保存した月次データをボタン1つで読み込み。アプリケーションID不要 |
| 月次Excel | e-Statの月次Excel（表15・16・18、建築物の表4-1・5・6-1・7-1）を読み込み |
| e-Stat API | e-Statのデータベース（2024年12月分まで）から取得。アプリケーションIDが必要 |

### 自動更新の仕組み

GitHub Actions（`.github/workflows/update-data.yml`）が毎日15時台（日本時間）に `scripts/update-data.mjs` を実行します。

1. e-Statの月次一覧ページから、2023年1月分以降の各月のページをたどる
2. 対象の表のExcelを取得し、アプリ本体と同じ読み取り処理で数値を取り出す
3. `data/<表>/<年-月>.json` と `data/index.json` を更新してコミットする

取り込み済みの月はとばし、直近3か月分は修正の有無を確認し直します。実行記録は `data/update.log` に残ります。手動で実行するときは、GitHubの Actions タブで update-data を選び「Run workflow」を押します。

## 出典

国土交通省「建築着工統計調査」（e-Stat）を加工して作成。地図は [dataofjapan/land](https://github.com/dataofjapan/land) の都道府県境界を簡略化して使用。e-StatのアプリケーションIDはブラウザ内にのみ保存され、このリポジトリには含まれません。
