# <a name="dialog-boxes-in-office-add-ins"></a>Office アドインのダイアログ ボックス
 
ダイアログ ボックスは、作業中の Office アプリケーション ウインドウの手前に浮動するサーフェスです。ダイアログ ボックスを使用すれば、作業ウィンドウで直接開くことができないサインイン ページ、ユーザーによるアクションを確認するための要求、作業ウィンドウ内で再生すると小さすぎるビデオの表示などのタスクのために追加の画面領域を提供できます。

**例:ダイアログ ボックス**

![ダイアログ ボックスの一般的なレイアウトを表示する画像の例](../images/overview_withApp_dialog.png)

### <a name="best-practices"></a>ベスト プラクティス

|**使用可能**|**使用不可**|
|:-----|:--------|
|<ul><li>アドイン名および現在のタスクを含む説明的なタイトルが含まれます。</li></ul>|<ul><li>タイトルには会社名を追加しません。</li></ul>|
||<ul><li>シナリオで必要な場合を除き、ダイアログ ボックスを開きません。</li></ul>|

## <a name="implementation"></a>実装

ダイアログ ボックスを実装するサンプルについては、GitHub の「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。

## <a name="additional-resources"></a>その他のリソース

- [UX パターンのサンプル](https://office.visualstudio.com/DefaultCollection/OC/_git/GettingStarted-FabricReact)
- [GitHub の開発リソース](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Dialog オブジェクト](https://dev.office.com/reference/add-ins/shared/officeui.dialog)


