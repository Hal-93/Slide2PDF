# Slide2PDF
<div style="display:flex; align-items:center; justify-content:space-between; gap:16px;">
  <div>
    <p>Googleスライドの埋め込みビューや通常ビューから、ワンクリックで全スライドをPDF化するChrome拡張機能です。</p>
  </div>
  <div>
    <img width="128" height="128" alt="icon" src="https://github.com/user-attachments/assets/0d29e79e-c6ae-4eb4-af67-66d915c5e990" />
  </div>
</div>

## 機能
<img width="210" height="38" alt="button" src="https://github.com/user-attachments/assets/f2020d87-5206-41e3-96e5-d387f7a38961" /><br>

- コントロールバーにダウンロードボタンが表示されるようになります
- ダウンロードボタンを押すことで自動でスライドをキャプチャしPDFとしてダウンロードします
- 埋め込みiframe内でも動作します

## 使い方
[リリース一覧](https://github.com/Hal-93/Slide2PDF/releases)より最新版をダウンロードしてください。

1. Chromeの拡張機能管理画面で「パッケージ化されていない拡張機能を読み込む」からダウンロードし、解凍したフォルダを指定して読み込む
2. Googleスライド（通常URL/埋め込みURLどちらでも可）を開く
3. 画面下部のコントロールバー内「オプション（︙）」の右隣に青い丸アイコンが出るのでクリック
4. ダウンロード完了後、自動で開始時のスライドに戻り、PDFが保存される

## 注意点
- スライド数が多い場合は処理に時間がかかります。完了まで操作せずお待ちください
- 途中でキャプチャが進行しなくなった場合、ページをリロードして再度ダウンロードしてください
