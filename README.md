
<!-- logo -->
<p align="center">
    <img src="https://raw.githubusercontent.com/JRC-Tojo/Rakubim/refs/heads/main/%E3%82%89%E3%81%8F%E3%81%B3%E3%82%80%E8%A1%A8%E7%B4%99.png" alt="らくびむロゴ" height="350"/></a>
</p>

<h1 align="center"> らくびむ </h1>

<h4 align="center">
    電子納品における納品ファイル名のフォーマットツール
</h4>

<p align="center">
    <!-- 実装言語 -->
    <img src="https://img.shields.io/badge/Language-VBA-purple.svg" alt="実装言語" />
    <!-- 最新バージョン -->
    <img src="https://img.shields.io/github/release/JRC-Tojo/Rakubim.svg" alt="最新バージョン" />
    <!-- ライセンス -->
    <img src="https://img.shields.io/badge/LICENCE-MIT-green.svg" alt="ライセンス" />
</p>



## 概要

「らくびむ」は、電子納品作業の効率化を目的としたExcelベースの業務改善ツールです。Excelシートに記載された図面情報と、指定されたフォルダ内にあるCADやPDFファイルを照合し、独自の自然言語処理アルゴリズムを用いて図面名とファイル名を高い精度でマッチングさせます。最終的に、所定の命名規則に従ってファイル名を自動でリネームし、出力します。

これにより、手作業によるファイル名の変更や整理といった煩雑な作業を大幅に削減し、ヒューマンエラーを防ぎます。



## 初めて使う方へ

1.  `らくびむ_VerX.X.X.xlsb` を開きます。
2.  Excelシートに必要な図面情報（図面番号、図面名など）を入力します。
3.  処理対象となるCADファイル（`.dxf`, `.dwg`など）とPDFファイルを、`CAD`, `PDF`フォルダに配置します。
4.  各シートに配置されたボタンをクリックして、以下の作業を順に実行できます。
    1. ファイルマッチング
    2. 更新ファイル名の確認
    3. リネームの実行
5.  処理が完了すると、リネームされたファイルが指定の出力先フォルダに出力されます。

> 社内向けの解説文書は[社内フォルダ](https://jrcjregroup.app.box.com/folder/348645297166)を参照してください。



## 開発サポート

本プロジェクトは新機能の提案や問題点の修正コードの投稿を歓迎しています。

コードを作成した際にはPull Requestを作成する形でぜひご投稿ください！

開発者向けの詳細情報は[開発者向け文書](./CONTRIBUTE.md)を参照してください。



## ライセンス

本プロジェクトのコードは商用・非商用の別にかかわらず自由に改変してご自身のプロジェクトに組み込むことが可能です。

使用前の事前相談等は一切不要ですが、本ライセンスは著作権の放棄ではないためご自身のプロジェクトのクレジット等で本プロジェクトの著作権表示を行う必要があります。

ただし、本プロジェクトがライブラリとして利用している以下のコードについては配布元のライセンスが適用されます。

- [`vba-files\Class\src\libs\SequenceMatcher`](https://github.com/CPkobo/vba-sequencematcher): 文字列の類似度を計測する
- [`vba-files\Class\src\libs\VBAtest`](https://github.com/vba-tools/vba-test): VBA環境上で単体テストを記述する


