# Deploy Guide

OpsPilot AIをポートフォリオとして一番きれいに公開するなら、**GitHub + Vercel** がおすすめです。理由は、URLがきれいで、更新が簡単で、将来ReactやAPI連携に拡張しても移行しやすいからです。

## Recommended: GitHub + Vercel

1. GitHubで新しいリポジトリを作成します。
   - Repository name: `opspilot-ai`
   - Visibility: Public または Private
2. このフォルダのファイルをリポジトリへアップロードします。
3. Vercelで `Add New Project` を選び、GitHubリポジトリを接続します。
4. Project Settingsは次のようにします。
   - Framework Preset: `Other`
   - Build Command: 空欄
   - Output Directory: `./`
5. Deployを押すと公開URLが発行されます。

## Alternative: Netlify

Netlifyも静的サイトに向いています。GitHub連携でも、フォルダをドラッグ&ドロップでも公開できます。

設定:

- Base directory: 空欄
- Build command: 空欄
- Publish directory: `.`

`netlify.toml` を含めているので、GitHub連携時もそのまま公開できます。

## Alternative: GitHub Pages

GitHubだけで完結したい場合はGitHub Pagesでも公開できます。

1. GitHubリポジトリの `Settings` を開きます。
2. `Pages` を選びます。
3. Sourceを `Deploy from a branch` にします。
4. Branchを `main`、folderを `/root` にします。
5. 数分後に `https://<username>.github.io/opspilot-ai/` で公開されます。

## Before Sharing

公開前に確認すること:

- `index.html` がトップページとして開ける
- 「Outlookから取込」でタスクが追加される
- 「Planner登録案を作成」でJSONが表示される
- READMEにデモの目的が書かれている
- 実運用のMicrosoft 365認証情報をコードに入れていない

## Portfolio Pitch

デモ紹介文の例:

> OpsPilot AIは、メールやチャットで届くバックオフィス依頼をAIで分類し、優先度、担当、期限、次アクションを提案するワークフロー管理デモです。Microsoft 365連携を想定し、Outlook依頼の取り込みとPlanner登録案の生成までを体験できます。
