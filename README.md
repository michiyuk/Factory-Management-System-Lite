📦 Factory Management System – Lite Edition
Excel VBA-based Factory Operations Automation Suite
（工場向け Excel 自動化システム・ライト版）

📘 概要
このリポジトリは、工場の 発注管理・在庫管理・作業管理・設備管理 を Excel VBA で自動化するための
「ライト版テンプレート」 です。

実際の商用版は複数ファイル連携・マスタ参照・メール自動化・設備レイアウト連携などを含みますが、
このリポジトリでは 設計思想・構造・サンプルコード を公開しています。

🧠 設計思想（Architecture）
本システムは Excel をフロントエンドとした
業務基幹システム（ERP / MES / CMMS） を想定しています。

コード
┌──────────────────────────────┐
│        Factory Management System        │
├──────────────────────────────┤
│ ① 発注管理（Order Management）         │
│ ② 在庫管理（Inventory Control）        │
│ ③ 作業管理（OH / 組立）               │
│ ④ 設備台帳（Asset Ledger）             │
│ ⑤ 設備レイアウト（Digital Twin）       │
│ ⑥ 保全履歴（Maintenance History）      │
└──────────────────────────────┘
🧩 モジュール構造（Lite版）
コード
/src
├─ Core
│   ├─ Mod_Utils.bas
│   ├─ Mod_Config.bas
│   └─ Mod_Logger.bas
│
├─ Order
│   ├─ Mod_OrderParser.bas
│   ├─ Mod_OrderWriter.bas
│   └─ Mod_OutlookDraft.bas
│
├─ Work
│   ├─ Mod_RowJudge.bas
│   ├─ Mod_RowButton.bas
│   └─ Mod_ProgressFlag.bas
│
├─ Highlight
│   ├─ Mod_HighlightController.bas
│   ├─ Mod_ColorManager.bas
│   └─ Mod_LabelManager.bas
│
└─ Web
    └─ Mod_WebSearch.bas
🧱 クラス構造（Lite版）
コード
/classes
├─ clsAppEvents.cls
├─ clsSheetEvents.cls
└─ clsConfig.cls
🧪 ダミーコード（安全なサンプル）
Mod_RowJudge.bas（抜粋）
vb
Option Explicit

' 行の状態を判定するサンプル（実際のロジックは非公開）
Public Function RowStatus(ByVal ws As Worksheet, ByVal r As Long) As String

    Dim maker As String
    maker = Trim$(ws.Cells(r, "C").Value)

    ' 空欄 → 内部確認
    If maker = "" Then
        RowStatus = "InternalCheck"
        Exit Function
    End If

    ' 特定文字を含む場合 → Web検索
    If InStr(maker, "TEST") > 0 Then
        RowStatus = "WebSearch"
        Exit Function
    End If

    ' それ以外 → 見積依頼
    RowStatus = "EstimateDraft"
End Function
Mod_WebSearch.bas（抜粋）
vb
Public Sub WebSearchLite(ByVal ws As Worksheet, ByVal r As Long)
    Dim q As String
    q = Trim$(ws.Cells(r, "C").Value)

    If q = "" Then Exit Sub

    Dim url As String
    url = "https://www.bing.com/search?q=" & q

    ThisWorkbook.FollowHyperlink url
End Sub
Mod_HighlightController.bas（抜粋）
vb
Public Sub HighlightSelection(ByVal ws As Worksheet, ByVal target As Range)
    Dim v As Variant
    v = target.Value

    Dim c As Range
    For Each c In ws.UsedRange
        If CStr(c.Value) = CStr(v) Then
            c.Interior.Color = RGB(204, 255, 204)
        End If
    Next c
End Sub
Mod_LabelManager.bas（抜粋）
vb
Public Sub ShowMatchLabel(ByVal ws As Worksheet, ByVal target As Range, ByVal count As Long)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes("shpMatchCount")
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddLabel(msoTextOrientationHorizontal, 0, 0, 120, 20)
        shp.Name = "shpMatchCount"
    End If

    shp.TextFrame.Characters.Text = "一致数: " & count
    shp.Left = target.Left + 100
    shp.Top = target.Top + 80
    shp.Visible = msoTrue

End Sub
🖼 画面イメージ（構成図）
コード
┌──────────────────────────────┐
│   作業リスト（Work List）     │
├───────────────┬──────────────┤
│  A:日付        │  F〜J:工程フラグ │
│  B:機器名      │  K:工数          │
│  C:メーカー    │  L:担当者        │
│  D:商品名      │  I:ボタン        │
└───────────────┴──────────────┘
📝 使い方（Lite版）
Excel を開く

「作業リスト」シートにデータを入力

C列にメーカー名を入力すると、

InternalCheck

WebSearch

EstimateDraft
のいずれかのボタンが自動生成されます

ボタンを押すとダミー処理が実行されます

⚙ 設定シート例（Lite版）
ini
[Config]
OrderFilePath = C:\Dummy\OrderList.xlsx
InventoryFilePath = C:\Dummy\Inventory.xlsx
OutlookTo = test@example.com
OutlookCC = cc@example.com
※ この設定シートは Lite 版のため実際には動作しません。
商用版では複数ファイル連携・商社マスタ・設備台帳・Outlook 自動送信など
高度な設定項目が追加されます。

🔐 商用版について
商用版では以下を含みます：

発注管理（Outlook 自動送信）

在庫管理（自動転記）

作業管理（達成判定・完了リスト）

設備レイアウト（図形検索）

設備台帳（リンク連携）

保全履歴（自動集計）

商社マスタ連携

金額自動分割

タスクスケジューラ連携

📩 商用版・カスタム依頼
GitHub の Issues または ココナラ からお問い合わせください。
