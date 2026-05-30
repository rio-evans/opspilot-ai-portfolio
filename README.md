# OpsPilot AI

バックオフィス向けのAI活用ポートフォリオアプリです。社内依頼を貼り付けると、疑似AIが業務分類、優先度、担当、期限、次アクションを推定し、ワークフローボードへ登録します。

## Features

- 依頼本文からタスクを自動登録
- 優先度、担当、期限、次アクションの提案
- カンバン形式のステータス管理
- 部門フィルタと検索
- KPIと改善インサイトの表示
- Microsoft 365連携デモ
  - Outlookメール風の依頼をタスク化
  - Planner登録用のGraph APIペイロードを生成
  - Entra ID Client ID / Planner Plan ID の入力欄
- CSVレポート出力
- LocalStorageによるデータ保存

## Demo Scenario

1. 社内チャットやメール本文を貼り付ける
2. AIが依頼を構造化して、担当と次アクションを提案する
3. カンバンで処理状況を可視化する
4. Microsoft 365連携でOutlook依頼の取り込みとPlanner登録案を見せる
5. 繰り返し業務を自動化候補として説明する

## Local Preview

このアプリは静的HTML/CSS/JavaScriptだけで動きます。APIキーやサーバーは不要です。

```bash
python -m http.server 4173
```

ブラウザで `http://localhost:4173/index.html` を開きます。

## Deployment

ポートフォリオ用途では、次の構成がおすすめです。

- GitHub: ソースコードとREADMEを公開
- Vercel または Netlify: 本番URLを発行
- 独自ドメイン: 必要に応じて `demo.your-domain.com` などを設定

このリポジトリはビルド不要です。公開ディレクトリはルート、ビルドコマンドは空で問題ありません。

詳しい手順は [DEPLOY.md](DEPLOY.md) を参照してください。

## Microsoft 365 Integration Plan

現在は認証なしで見せられるデモモードです。実運用化する場合は、Microsoft Entra IDでアプリ登録を行い、MSAL.jsでサインイン後、Microsoft Graph APIを呼び出します。

代表的な連携先:

- Outlook: `GET /me/messages`
- Planner: `POST /planner/tasks`
- Excel / SharePoint: 業務データの集計やレポート生成
- Teams: チャネル投稿やメンションから依頼を取り込み
