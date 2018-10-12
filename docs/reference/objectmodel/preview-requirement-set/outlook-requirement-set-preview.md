# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook アドイン API 要件セットのプレビュー

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> 注:このドキュメントは、[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)  の**プレビュー** 用です。 この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。 アドイン マニフェストでこの要件を指定しないでください。 この要件のセットに導入されているメソッドとプロパティは、使用前に可用性を個別にテストする必要があります。

要件セットのプレビューには、[要件セット 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) のすべての機能が含まれています。

## <a name="features-in-preview"></a>プレビューの機能

次の機能はプレビュー段階です。

- [SharedProperties](/javascript/api/outlook/office.sharedproperties) - には、共有フォルダー、予定表、またはメールボックス内のメッセージ、予定、またはアイテムのプロパティを表す新しいオブジェクトが追加されます。
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-) - 1 つの有効な値 `allowEvent` を持つディクショナリである、新しいオプション パラメーター `options`。この値は、イベントの実行をキャンセルするために使用されます。
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - は、添付ファイルの base64 エンコードをメッセージまたは予定を新しいメソッドを追加します。
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) - アドインが[操作可能メッセージによってアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)されると渡される初期化データを返す新しい機能が追加されました。
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback) - は、メッセージ アイテムまたは予定の sharedProperties を表すオブジェクトを取得する新しいメソッドを追加します。
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) - Microsoft Graph API の[アクセス トークンの取得](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)をアドインに対して許可する、`getAccessTokenAsync` へのアクセスが追加されました。
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) - は、代理アクセス権を指定する新しいビット フラグ列挙を追加します。
- [Office.EventType](/javascript/api/office/office.eventtype) をサポートするように OfficeThemeChanged のイベントの追加を `OfficeThemeChanged` のエントリです。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](https://docs.microsoft.com/outlook/add-ins/quick-start)