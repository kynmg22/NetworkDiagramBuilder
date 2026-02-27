# ネットワーク構成図ビルダー

FortiGate / ネットワーク機器の接続関係をGUIで設定して  
draw.io (.drawio) ファイルとして出力するWindowsアプリです。

---

## ビルド方法

### 必要なもの
- **Visual Studio 2022** または **.NET 8 SDK**
  - https://dotnet.microsoft.com/download/dotnet/8.0

### 単体EXE作成（コマンドプロンプト）

```
build.bat
```

または手動で:

```cmd
dotnet publish NetworkDiagramBuilder.csproj ^
  -c Release ^
  -r win-x64 ^
  --self-contained true ^
  -p:PublishSingleFile=true ^
  -p:IncludeNativeLibrariesForSelfExtract=true ^
  -o ./publish
```

→ `publish\NetworkDiagramBuilder.exe` が生成されます（約50MB）

---

## メモファイルのフォーマット

```
【構成機器】
機器名　型番　IPアドレス
FW      FG-70G  192.168.1.1
SW-L2   GS-748  192.168.1.2
AP      FAP-231 192.168.1.10
【構成以上】

【以下メモ欄】
・FWはWAN側でISPと接続
・AP は SW-L2 のポート8に接続
【メモ欄以上】
```

- 区切りは **タブ**・**全角スペース**・**2文字以上の半角スペース** のいずれでもOK
- IPアドレスは省略可能

---

## 操作の流れ

1. **メモ読込**
   - テキストファイルをドラッグ&ドロップ、またはテキストを直接貼り付け
   - 「解析して接続設定へ」をクリック

2. **接続設定**
   - 「＋ 接続ブロックを追加」で接続関係を定義
   - 親機器を選択 → 接続先を追加（複数可）
   - 「接続なし（端末）」= それ以上ぶら下がりなし
   - 一度使った機器は選択肢から消えるが、プルダウン下部の  
     「配置済み機器（二重接続用）」から再選択可能

3. **XML生成**
   - `.drawio` として保存 → draw.io で開いて編集可能

---

## 機器スタイルの自動判定

機器名・型番に以下の文字列が含まれると Cisco シェイプが自動選択されます:

| キーワード | シェイプ |
|-----------|---------|
| fortigate / fg- / fw / firewall / asa | ファイアウォール |
| sw / switch / hub / catalyst | L2スイッチ |
| router / rt / rtx / yamaha | ルーター |
| ap / wifi / wireless | アクセスポイント |
| nas / server / srv | サーバー |
| pc / client / desktop | PC |
| modem / onu / olt | モデム |
