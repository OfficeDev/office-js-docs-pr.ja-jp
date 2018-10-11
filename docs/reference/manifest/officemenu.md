# <a name="officemenu-element"></a>OfficeMenu 要素

Office のコンテキスト メニューに追加されるコントロールのコレクションを定義します。Word、Excel、PowerPoint、OneNote アドインに適用されます。

## <a name="attributes"></a>属性

| 属性            | 必須 | 説明                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | はい      | 定義されている OfficeMenu の種類。|

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [コントロール](#control)    | はい |  1 つ以上のコントロール オブジェクトのコレクション。  |

## <a name="xsitype"></a>xsi:type

この Office アドインを追加する Office クライアント アプリケーションの組み込みメニューを指定します。

- `ContextMenuText` - テキストが選択され、その選択されたテキストのコンテキスト メニューをユーザーが開いたときに (右クリック)、コンテキスト メニューに項目が表示されます。Word、Excel、PowerPoint、OneNote に適用されます。
- `ContextMenuCell` - ユーザーがスプレッドシートのセルのコンテキスト メニューを開くと (右クリック)、コンテキスト メニューに項目が表示されます。Excel に適用されます。 

## <a name="control"></a>コントロール

各 **OfficeMenu** 要素には、1 つ以上の[メニュー](control.md#menu-dropdown-button-controls)コントロールが必要です。 

## <a name="example"></a>例

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
