# 📘 GA4データ自動更新 設定ガイド

このガイドでは、特定顧客のスプレッドシートをGA4と連携させ、**毎日自動でデータを更新する**ための手順を説明します。

---

## 1. 準備：Google Sheetsアドオンのインストール

1.  対象の顧客スプレッドシートを開きます。
2.  上部メニューの **「拡張機能」** > **「アドオン」** > **「アドオンを取得」** を選択。
3.  検索窓に **「GA4 Reports」** と入力し、公式の `GA4 Reports Builder for Google Analytics` をインストールしてください。

---

## 2. レポートの作成（初期設定）

1.  **「拡張機能」** > **「GA4 Reports builder...」** > **「Create new report」** を選択。
2.  右側のパネルで以下を設定します：
    - **Report Name**: `GA4_Auto_Update`
    - **Property**: 対象顧客のGA4プロパティを選択
    - **Dimensions**: `date`
    - **Metrics**: `sessions`, `activeUsers`, `engagementRate`, `conversions`, `sessionConversionRate`, `advertiserAdCost`
3.  **「Create Report」** をクリックします。
    - `Report Configuration` という新しいシートが作成されます。

---

## 3. 自動更新のスケジュール設定

1.  **「拡張機能」** > **「GA4 Reports builder...」** > **「Schedule Reports」** を選択。
2.  **「Enable reports to run automatically」** にチェックを入れます。
3.  更新頻度（おすすめ：Every day, 4am-5am）を設定し、**「Save」** を押します。

これで、毎朝自動的に最新データがスプレッドシートに書き込まれます。

---

## 4. ダッシュボードへの反映方法

本ダッシュボードは、アドオンが作成したシート名（通常は `GA4_Auto_Update`）を探して自動的に読み込みます。

### ハイブリッド運用の注意点
- **自動化する顧客**: 上記設定を1度行えば、それ以降は入力を忘れてもダッシュボードに最新値が出ます。
- **手動の顧客**: これまで通り `GA4` タブに手動で貼り付けてください。

---
© 2026 マーケティング・ダッシュボード開発チーム
