# Changelog

<!-- GAS version は deploy.yml の clasp version で自動作成されます。
     各リリースの GAS バージョン番号は GitHub Actions のジョブサマリーで確認できます。 -->

## [1.6.0] - 2026-09-08
### Added
- 動画シートの D2 に YouTube 動画 ID を記録するようにした
  - WebUI でサムネイル表示と YouTube へのリンクに使う
  - 現在の取得対象（プレイリスト / EXTRA_VIDEO_IDS）に含まれる既存シートは次回実行時に自動補完される
  - 取得対象から外れた動画のシートは `processVideo_` を通らないため D2 は空のまま

## [1.5.2] - 2026-09-06
### Fixed
- `_bookmarks` シートを `PRESERVE_SHEET_NAMES` に追加
  - WebUI のブックマーク（動画選択の保存）が `resetSheets()` で消えないようにする

## [1.5.1] - 2026-09-05
### Fixed
- 動画シートの日時列（A列）の表示書式を毎回 `yyyy/mm/dd hh:mm:ss` に統一するようにした
  - Sheets は 00:00 ちょうどの Date を書いたセルを自動的に「日付のみ」書式にする。`Range.sort()` は値だけを並び替えて書式をセル位置に残すため、その書式のセルに別の時刻の記録が入り込んでいた
  - 結果として Sheets API を `FORMATTED_VALUE` で読む側（playcount-webui）で時刻が 00:00 に潰れ、再生数グラフが単調増加のはずの箇所で山になっていた

## [1.5.0] - 2026-05-26
### Fixed
- データ間引き（サンプリング）の基準を「動画公開日からの経過日数」から「現在時刻からの経過日数」に修正
  - 従来の実装では古い動画の直近データも古いデータとして圧縮されてしまい、1時間前・1日前・1週間前の比較が不正確になっていた
  - 直近30日以内のデータは動画の年齢に関わらず全件保持されるようになった

## [1.4.1] - 2026-05-22
### Fixed
- アップロードプレイリストに未収録の追跡動画（プレイリスト経由追加など）がチャンネル内順位を取得できない問題を修正
  - `computeRanksByChannelGroups_` で追跡動画 ID を `allIds` に補完してから順位計算するように変更

## [1.4.0] - 2026-05-21
### Added
- `main()` 実行後に `_settings` シートの `last_fetched_at` キーを現在時刻（ISO 8601）で更新
  - WebUI がこの値をポーリングして自動的にデータを再取得するための更新マーカー
- `setSetting_(ss, key, value)` ユーティリティ関数を追加（`_settings` シートのキーを find-or-append で更新）

## [1.3.0] - 2026-05-19
### Added
- 再生数比較シートに「順位」列を追加（`チャンネル内順位` シートの最新データから取得）
- `buildLatestRankMap_()` を追加: 順位シート最新行から `{ 動画タイトル: rank }` を返すヘルパー

## [1.2.1] - 2026-05-17
### Changed
- `deploy.yml`: `fetch-depth: 0` を廃止し `fetch-depth: 1` + `fetch-tags: true` に変更（全履歴取得を解消）
- `auto-merge.yml`: `synchronize` トリガーを削除（push毎の不要な起動を抑制）
- `claude-review.yml`: timeout-minutes を 15 → 10 に短縮

## [1.2.0] - 2026-05-16
### Changed
- `CONFIG.PLAYLIST_ID`（単一文字列）を `CONFIG.PLAYLIST_IDS`（配列）に変更し複数プレイリストをサポート
- `fetchPlaylistVideoIds_()` が `PLAYLIST_IDS` の全プレイリストを順に取得するよう変更
- `_settings` シートのキーを `playlist_id` → `playlist_ids`（カンマ区切り複数対応）に変更
- `loadConfig_()` が `playlist_ids` キーを読んで `PLAYLIST_IDS` 配列を上書きするよう変更

## [1.0.5] - 2026-05-13
### Changed
- チャンネル内順位グラフをデータテーブル下の縦積みから行1の横一列並びに変更

## [1.0.4] - 2026-05-13
### Fixed
- `main()` 実行後にシートタブが投稿日昇順に並び替えられない問題を修正（`sortVideoSheetsByPublishDate_` を毎時呼び出すよう変更）
### Added
- 通常PRの自動マージワークフロー（`auto-merge.yml`）を追加

## [1.0.2] - 2026-05-01
### Changed
- `シート1` を `PRESERVE_SHEET_NAMES` から除外
### Added
- `deleteSheet1()` 管理関数を追加

## [1.0.1] - 2026-04-24
### Fixed
- シート並び替え関数 `sortVideoSheetsByPublishDate_` で `PRESERVE_SHEET_NAMES` の定義順が無視されるバグを修正
