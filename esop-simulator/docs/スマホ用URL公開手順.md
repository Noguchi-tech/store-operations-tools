# スマホで使えるURLの作り方（会員登録なし）

このツールは1枚の `index.html` だけで動くため、公開リポジトリに置けば
**ログイン・会員登録なし**でスマホからそのまま使えます。方法は2つあります。

---

## ① いますぐ使える簡易URL（設定不要）

現在のブランチのファイルを、`htmlpreview` というサービス経由で表示します。
リポジトリが公開（public）なので、設定なしで今すぐ開けます。

```
https://htmlpreview.github.io/?https://raw.githubusercontent.com/Noguchi-tech/store-operations-tools/claude/esop-simulator-tool-i5fnd0/esop-simulator/index.html
```

- スマホのブラウザにこのURLを貼るだけで動きます。会員登録は不要です。
- 注意：このURLは作業用ブランチ（`claude/esop-simulator-tool-i5fnd0`）を指しています。
  ブランチを削除すると開けなくなります。**正式運用には下の②をおすすめします。**

---

## ② 正式な公開URL（GitHub Pages・おすすめ）

きれいで覚えやすい固定URLになります。一度設定すれば、以後はファイル更新が自動反映されます。
設定はリポジトリの管理者（オーナー）操作が1回だけ必要です。

### 手順

1. このブランチの内容を `main` に取り込む（マージ）。
   - GitHubの「Pull requests」から作業ブランチ → `main` のPRを作成して `Merge` するだけです。
   - （ご希望であれば、こちらでPRを作成します。お申し付けください。）
2. リポジトリの **Settings → Pages** を開く。
3. **Source** を「Deploy from a branch」にする。
4. **Branch** を `main`、フォルダを `/ (root)` にして **Save**。
5. 1〜2分待つと、Pagesが有効になります。

### 公開URL（②設定後）

```
https://noguchi-tech.github.io/store-operations-tools/esop-simulator/
```

- ログイン不要・スマホ対応。社内外問わず、このURLを共有するだけで使えます。
- ファイルを更新して `main` に反映すると、URLの中身も自動で新しくなります。

---

## 対面説明でのコツ

- 上記URLから **QRコード** を作っておくと、面談相手がスマホで一瞬で開けます。
  （「QRコード 作成」で出る無料サイトでURLを貼ればOK。アプリ登録は不要です。）
- ホーム画面に追加すると、アプリのように起動できます。

---

## まとめ

| 方法 | URL | 設定 | 安定性 |
|---|---|---|---|
| ① 簡易（htmlpreview） | 上記の長いURL | 不要・今すぐ | ブランチ依存（暫定向き） |
| ② GitHub Pages | `https://noguchi-tech.github.io/store-operations-tools/esop-simulator/` | 1回だけ管理者操作 | 安定（正式運用向き） |
