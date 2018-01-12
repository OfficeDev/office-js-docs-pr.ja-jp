# <a name="toggle-component-in-office-ui-fabric"></a>Office UI Fabric の切り替えコンポーネント

切り替えは、オンまたはオフにするなどの物理的なスイッチを表しています。 相互排他的な 2 つのオプション (オンまたはオフなど) を表示するには、切り替えを使用します。オプションを選択すると、すぐにアクションが実行されます。
  
#### <a name="example-toggle-in-a-task-pane"></a>例:作業ウィンドウ内の切り替え


![切り替えが表示された画像](../../images/overview_withApp_toggle.png)

<br/>

## <a name="best-practices"></a>ベスト プラクティス

|**するべきこと**|**してはいけないこと**|
|:------------|:--------------|
|変更を即座に適用する場合は、バイナリ設定に切り替えを使用します。<br/><br/>![切り替えでするべきことの例](../../images/toggleDo.png)<br/>|変更が有効になる前にユーザーが追加の手順を実行する必要がある場合は、切り替えを使用しないでください。<br/><br/>![切り替えでしてはいけないことの例](../../images/toggleDont.png)<br/>|
|設定に使用する特定のラベルがある場合は、**[オン]** と **[オフ]** のラベルのみを交換します。 反対のバイナリを表す短いラベル (3 から 4 文字) を使用します。| |

## <a name="variants"></a>バリアント

|**バリエーション**|**説明**|**例**|
|:------------|:--------------|:----------|
|**有効でチェック済み**|切り替え状態がアクティブな場合に使用します。|![有効でチェック済みの画像](../../images/toggleEnabledOn.png)<br/>|
|**有効で未チェック**|切り替え状態が非アクティブな場合に使用します。|![有効で未チェックの画像](../../images/toggleEnabledOff.png)<br/>|
|**無効でチェック済み**|アクティブな状態を変更できない場合に使用します。|![無効でチェック済みの画像](../../images/toggleDisabledOn.png)<br/>|
|**無効で未チェック**|非アクティブの状態を変更できない場合に使用します。|![無効で未チェックの画像](../../images/toggleDisabledOff.png)<br/>|

## <a name="implementation"></a>実装

詳細については、「[切り替え](https://dev.office.com/fabric#/components/toggle)」と「[Fabric React のコード サンプルの使用にあたって](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)」を参照してください。

## <a name="additional-resources"></a>その他のリソース

- [UX 設計パターン](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office アドインの Office UI Fabric](office-ui-fabric.md)
