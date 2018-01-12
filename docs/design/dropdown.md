# <a name="dropdown-component-in-office-ui-fabric"></a>Office UI Fabric のドロップダウン コンポーネント

ドロップダウンは、ドロップダウン ボタンをクリックすると表示される、オプションのリストです。 ドロップダウン リストまたはメニューを使用して、UI デザインを簡素化します。また、ユーザーが UI 内で選択する必要がある場合にも使用します。 リストを折りたたむと、選択した項目が表示されます。 選択した項目を変更するには、ユーザーはリストを開いて、新しい値を選択します。
  
#### <a name="example-drop-down-in-a-task-pane"></a>例:作業ウィンドウ内のドロップダウン

<br/>

![ドロップダウンが表示された画像](../../images/overview_withApp_dropdown.png)

<br/>

## <a name="best-practices"></a>ベスト プラクティス

|**するべきこと**|**してはいけないこと**|
|:------------|:--------------|
|既定の選択済みオプションが、他のオプションよりも選択される可能性が高い場合は、ドロップダウンを使用します。 対照的に、ChoiceGroup やラジオ ボタンはすべての選択肢を表示します。そのため、すべてのオプションを同じように強調します。|すべてのオプションが同じように選択される可能性がある場合は、ドロップダウンは使用しないでください。|
|1 つのフィールドに折りたたむことができる、複数の選択肢がある場合は、ドロップダウンを使用します。 長い項目リストがある場合や、画面のスペースが制限されている場合にもドロップダウンを使用します。|選択肢が 2 つ未満の場合は、ドロップダウンを使用しないでください。 代わりに、チェック ボックスを使用します。|
|ドロップダウン リストでは、短い文や単語を使用します。| |

## <a name="variants"></a>バリアント

|**バリエーション**|**説明**|**例**|
|:------------|:--------------|:----------|
|**基本的な未制御のドロップダウン**|多くのオプションが選択可能な場合に使用します。|![基本的な未制御のドロップダウンの画像](../../images/dropdownUncontrolled.png)<br/>|
|**defaultSelectedKey を使用して無効化された未制御のドロップダウン**|ドロップダウンが無効にされた状態です。|![defaultSelectedKey を使用して無効化された未制御のドロップダウンの画像](../../images/dropdownDisabled.png)<br/>|
|**制御されたドロップダウン**|既定の選択済み項目が UI の別の場所の影響を受け、ドロップダウン内の選択された項目を保持する必要がある場合に使用します。|![制御されたドロップダウンの画像](../../images/dropdownControlled.png)<br/>|

## <a name="implementation"></a>実装

詳細については、「[ドロップダウン](https://dev.office.com/fabric#/components/dropdown)」と「[Fabric React のコード サンプルの使用にあたって](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)」を参照してください。

## <a name="additional-resources"></a>その他のリソース

- [UX 設計パターン](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office アドインの Office UI Fabric](office-ui-fabric.md)
