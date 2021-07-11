Outlookは、主に Mailbox オブジェクトを介して公開される API を[使用](/javascript/api/outlook/office.mailbox)します。 Outlook アドイン専用のオブジェクトおよびメンバー (たとえば、[Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) オブジェクトなど) にアクセスするには、次のコード行に示すように、[Context](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) オブジェクトの **mailbox** プロパティを使用して、**Mailbox** オブジェクトにアクセスします。

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

さらに、Outlookアドインは次のオブジェクトを使用できます。

-  **Office** オブジェクト: 初期化に使用します。

-  **Context** オブジェクト: コンテンツおよび表示言語のプロパティへのアクセスに使用します。

-  **RoamingSettings** オブジェクト: アドインがインストールされているユーザーのメールボックスに Outlook アドイン固有のカスタム設定を保存する際に使用します。

JavaScript API のOutlookについては[、「Outlook」を参照してください](../outlook/outlook-add-ins-overview.md)。