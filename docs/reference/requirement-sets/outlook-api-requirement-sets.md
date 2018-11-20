# <a name="outlook-javascript-api-requirement-sets"></a>Outlook JavaScript API の要件セット

Outlook アドインでは、[マニフェスト](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)で [Requirements](/office/dev/add-ins/reference/manifest/requirements) 要素を使用し、必要な API のバージョンを宣言します。 Outlook アドインには、`Name` 属性が `Mailbox` に設定され、`MinVersion` 属性がアドインのシナリオをサポートする最小 API 要件セットに設定された [Set](/office/dev/add-ins/reference/manifest/set) 要素が常に含まれます。

たとえば、次のマニフェストのスニペットは、最小要件セットの 1.1 を表します。

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

すべての Outlook API は `Mailbox` [要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)に属しています。 `Mailbox` 要件のセットにはバージョンがあります。リリースされる新しい API の各セットは、新しいバージョンのセットに属しています。 すべての Outlook クライアントが最新の API セットをサポートしているわけではありません。しかし Outlook クライアントが要件セットのサポートを宣言する場合は、その要件セットの API すべてがサポートされています。

マニフェストに要件セットの最小バージョンを設定することで、アドインが表示される Outlook クライアントをコントロールできます。クライアントが最小要件セットをサポートしない場合、アドインはロードされません。たとえば、要件セットのバージョン 1.3 が指定されている場合、1.3 以上をサポートしていない Outlook クライアントには表示されません。

## <a name="using-apis-from-later-requirement-sets"></a>後続の要件セットからの API の使用

要件セットを設定しても、アドインで使用できる API は制限されません。 たとえば、アドインでは要件セット 1.1 が指定されていますが、1.3 をサポートしている Outlook クライアントで実行されている場合、アドインでは要件セット 1.3 の API を使用できます。

より新しい API を使用するには、開発者は標準の JavaScript の技法を使用することで、新しい API の有無を確認できます。

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

このようなチェックは、マニフェストで指定された要件セット バージョンに存在する API には必要ありません。

## <a name="choosing-a-minimum-requirement-set"></a>最小要件セットの選択

開発者は、アドインを使用するために必要な、シナリオで必須の API のセットが含まれている初期の要件セットを使用する必要があります。

## <a name="clients"></a>クライアント

以下のクライアントは、Outlook のアドインをサポートしています。

| クライアント | サポートされる API の要件セット |
| --- | --- |
| Windows 用 Outlook 2019 | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6)、[1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2019 for Mac | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6)、[1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Windows 用 Outlook 2016 (クイック実行) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6)、[1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Windows 用 Outlook 2016 (MSI) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook 2016 for Mac | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Windows 用 Outlook 2013 | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook for iPhone | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook for Android | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook on the web (Office 365 および Outlook.com) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook Web App (オンプレミスの Exchange 2013) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1) |
| Outlook Web App (オンプレミスの Exchange 2016) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3) |
| Outlook Web App (オンプレミスの Exchange 2019) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |

> [!NOTE]
> Outlook 2013 での 1.3 のサポートは、「[2015 年 12 月 8 日付、Outlook 2013 用更新プログラム (KB3114349)](https://support.microsoft.com/kb/3114349)」の一部として追加されました。 Outlook 2013 での 1.4 のサポートは、「[MS16-107: Outlook 2013 セキュリティ更新プログラムについて 2016 年 9 月 13 日](https://support.microsoft.com/help/3118280)」の一部として追加されました。
