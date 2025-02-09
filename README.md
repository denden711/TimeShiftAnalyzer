# TimeShiftAnalyzer

## 概要
**TimeShiftAnalyzer**は、複数のCSVファイルの時間と電圧データを処理し、基準となるファイルを元に他のデータを時間軸上でシフトさせて解析するツールです。指定された時間範囲内でゼロ電圧点を基準にし、データをプロットおよびExcelファイルに出力します。

## 特徴
- 複数のCSVファイルの時間データを基準に合わせてシフト
- ユーザーが指定した時間範囲内でのゼロ点調整
- シフト後の結果をExcelファイルに保存（科学的記数法で20桁表示）
- グラフの自動プロット
- GUIを使ってドラッグ&ドロップで簡単にファイルを選択可能

## システム要件
- Python 3.x
- Windows, macOS, Linux

## 依存関係
以下のPythonライブラリを使用しています。インストールされていない場合は、`pip`コマンドを使ってインストールしてください。

```bash
pip install pandas matplotlib tkinterdnd2 openpyxl
```

## インストール

1. **Python環境のセットアップ**  
   Python 3.xがインストールされていることを確認してください。インストールされていない場合は、[Python公式サイト](https://www.python.org/)からダウンロードしてください。

2. **必要なライブラリのインストール**  
   ターミナルやコマンドプロンプトで次のコマンドを実行し、必要なライブラリをインストールしてください。

   ```bash
   pip install pandas matplotlib tkinterdnd2 openpyxl
   ```

3. **TimeShiftAnalyzerのダウンロード**  
   プロジェクトフォルダにプログラムファイルを配置してください。

## 使い方

1. **プログラムの起動**  
   `TimeShiftAnalyzer.py`を実行して、GUIを起動します。

   ```bash
   python TimeShiftAnalyzer.py
   ```

2. **基準ファイルの選択**  
   GUI上の「Select Baseline CSV」ボタンをクリックし、基準となるCSVファイルを選択するか、基準ファイルをウィンドウにドラッグ&ドロップします。

3. **シフト対象ファイルの選択**  
   シフトさせたい複数のCSVファイルをGUIのリストボックスにドラッグ&ドロップします。

4. **時間範囲の指定**  
   「Enter Time Range」フィールドにシフトの基準となる時間範囲を入力します。デフォルト値は`5e-05`から`8e-05`です。

5. **実行**  
   「Execute Shift and Plot」ボタンを押すと、時間シフトが計算され、結果がプロットされます。結果はExcelファイルに保存されます。

## 出力結果
- **Excelファイル**: シフトされたデータが科学的記数法（20桁）で保存されます。ファイル名、ゼロ点の時間、シフト値が記載されます。
- **グラフ**: 基準データとシフトされたデータがグラフとしてプロットされ、タイムシフトの様子を視覚的に確認できます。

## 注意事項
- 各CSVファイルの18列目に時間データ、5列目に電圧データが含まれていることを確認してください。
- 時間範囲が適切に設定されていない場合、エラーメッセージが表示されます。再度正しい範囲を入力してください。

## トラブルシューティング
- **ファイルが読み込めない場合**: ファイルパスに特殊文字が含まれているか、ファイル形式が異なる可能性があります。ファイルパスを再確認し、CSVファイル形式を使用していることを確認してください。
- **時間範囲のエラー**: 時間範囲の入力が数値でない場合、エラーメッセージが表示されます。正しい形式で数値を入力してください。

## ライセンス
このプログラムはMITライセンスの下で提供されています。