# 台本工房

ライブ・イベントの**進行台本（Word）**と**プロンプター（PowerPoint）**を、AIに作ってもらうための依頼書づくりツールです。

**ツールを開く → https://isumiaki.github.io/daihon-kobo/**

**使い方ガイド → https://isumiaki.github.io/daihon-kobo/guide.html**

## 使い方

1. ツールに公演情報・キャストと色・曲と歌詞・セットリスト・部署別のキュー（PA／VJ／照明／舞台）を入れる
2. 「コピー」または「.mdで保存」で依頼書を書き出す
3. 依頼書を使っているAI（Claude・ChatGPT・Geminiなど）に渡す
4. AIが出す台本のたたき台を確認してOKを出すと、Wordの台本とPPTXのプロンプターが書き出される

入力内容とライブラリ（登録した曲・キャスト）は、使っているブラウザの中にだけ保存されます。サーバーには送られません。

## 入っているファイル

| ファイル | 内容 |
|---|---|
| `index.html` | ツール本体 |
| `guide.html` | 使い方ガイド（Web版） |
| `台本工房_使い方ガイド.docx` | 使い方ガイド（Word版） |
| `台本プロンプター作成仕様書.md` | AIに渡す仕様書（ツールが書き出す依頼書にも同じものが入る） |
| `build_script.py` | 台本とプロンプターを書き出すプログラム（仕様書の付録Aと同じ） |
| `sample/` | 架空の公演で作った見本データと出力例 |

## 書き出しプログラムを自分で動かす場合

```
pip install python-docx python-pptx
python build_script.py sample/sample.json
```
