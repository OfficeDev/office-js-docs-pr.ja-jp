# <a name="customtab-element"></a>CustomTab 要素

リボン上で、アドイン コマンドに使用するタブとグループを指定します。これは既定のタブ (**[ホーム]**、**[メッセージ]**、または **[会議]** のいずれか) か、アドインで定義されたカスタム タブになります。

カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、一つのカスタム タブに制限されています。

**id** 属性はマニフェスト内で一意でなければなりません。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [グループ](group.md)      | はい |  コマンドのグループを定義します。  |
|  [ラベル](#label-tab)      | はい |  CustomTab または Group のラベル。  |
|  [コントロール](control.md)    | はい |  一つ以上のコントロール オブジェクトのコレクション。  |

### <a name="group"></a>グループ

必須です。[Group 要素](group.md)を参照してください。

### <a name="label-tab"></a>ラベル(タブ)

必須。カスタム タブのラベルです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。


## <a name="customtab-example"></a>CustomTab の例

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```