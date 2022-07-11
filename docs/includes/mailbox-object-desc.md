Outlook アドインでは、主に [Mailbox](/javascript/api/outlook/office.mailbox) オブジェクトにより公開されている API のサブセットを使用します。 [Item](/javascript/api/outlook/office.item) オブジェクトなど、Outlook アドインで特に使用するオブジェクトとメンバーにアクセスするには、次のコード行に示すように、**Context** オブジェクトの [メールボックス](/javascript/api/office/office.context#office-office-context-mailbox-member) プロパティを使用して **Mailbox** オブジェクトにアクセスします。

```js
// Access the Item object.
const item = Office.context.mailbox.item;
```

さらに、Outlook アドインでは、次のオブジェクトを使用できます。

- **Office** オブジェクト: 初期化に使用します。

- **Context** オブジェクト: コンテンツおよび表示言語のプロパティへのアクセスに使用します。

- **RoamingSettings** オブジェクト: アドインがインストールされているユーザーのメールボックスに Outlook アドイン固有のカスタム設定を保存する際に使用します。

Outlook アドインでの JavaScript の使用については、「[Outlook アドイン](../outlook/outlook-add-ins-overview.md)」を参照してください。
