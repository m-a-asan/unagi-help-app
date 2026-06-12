# うなぎ屋 ヘルプ人員調整システム

複数店舗間でヘルプスタッフの調整をするWebアプリです。

## 機能

- 店舗管理者：ヘルプ依頼の作成・他店舗の依頼に応答
- 全体シフト管理者：応答の承認・却下・店舗管理
- 部分承認対応：複数店舗の応答を承認でき、承認済み人数の合計が必要人数に達した時点で依頼が充足（例：3名必要に対し2店舗から1名＋2名）
- スタッフ種別：ホール / 焼き / キッチン
- 最大20店舗対応

## 構成

- `index.html` … フロントエンド（GitHub Pages などの静的ホスティングで公開）
- `Code.gs` … バックエンド（Google Apps Script + Google スプレッドシート）

フロントエンドは JSONP（GET リクエスト）で GAS と通信します。

## セットアップ

### 1. スプレッドシートの準備

Google スプレッドシートを新規作成し、以下の3つのシートを作成して、
それぞれ1行目に次のヘッダーを入力します。

| シート名 | 1行目のヘッダー（列順は自由） |
| --- | --- |
| `Stores` | `id`, `name` |
| `Requests` | `id`, `storeId`, `date`, `time`, `staffType`, `count`, `note`, `status`, `createdAt` |
| `Responses` | `requestId`, `storeId`, `count`, `names`, `status` |

`Stores` シートに最低1店舗（例: `1`, `本店`）を入力してください。

### 2. Google Apps Script のデプロイ

1. スプレッドシートのメニューから「拡張機能 → Apps Script」を開く
2. `Code.gs` の内容を貼り付けて保存
3. 「デプロイ → 新しいデプロイ → ウェブアプリ」を選択
   - 次のユーザーとして実行：**自分**
   - アクセスできるユーザー：**全員**
4. 発行された ウェブアプリURL（`https://script.google.com/macros/s/.../exec`）をコピー

### 3. フロントエンドの設定

`index.html` 内の `API_URL` を手順2でコピーしたURLに書き換えます。

```js
const API_URL = 'https://script.google.com/macros/s/＜あなたのデプロイID＞/exec';
```

## GitHub Pages でのデプロイ

1. GitHubに新しいリポジトリを作成
2. このフォルダをプッシュ
3. Settings → Pages → Source を「main branch」に設定
4. 数分後に `https://<username>.github.io/<repo-name>/` で公開されます

## 注意事項

- 認証機能はありません。URLを知っている人は誰でも操作できるため、店舗間の内部利用を想定しています。
- `Code.gs` を更新した場合は、再度「デプロイ → デプロイを管理」から新しいバージョンを発行してください。
