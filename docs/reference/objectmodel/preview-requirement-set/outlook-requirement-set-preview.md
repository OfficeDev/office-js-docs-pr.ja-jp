---
title: Outlook アドイン API 要件セットのプレビュー
description: ''
ms.date: 03/07/2019
localization_priority: Priority
ms.openlocfilehash: b1a3f5c675b2bcb43003ad15b3358e3febd80260
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512861"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook アドイン API 要件セットのプレビュー

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)の**プレビュー**用です。 この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。 アドイン マニフェストでこの要件を指定しないでください。 この要件のセットに導入されているメソッドとプロパティは、使用前に可用性を個別にテストする必要があります。 また、場合によっては [Office Insider プログラム](https://products.office.com/office-insider)に参加する必要もあります。

要件セットのプレビューには、[要件セット 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) のすべての機能が含まれています。

## <a name="features-in-preview"></a>プレビューの機能

次の機能はプレビュー段階です。

### <a name="add-in-commands"></a>アドイン コマンド

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[Event.completed](/javascript/api/office/office.addincommands.event#completed-options-)

1 つの有効な値 `allowEvent` を持つディクショナリである、新しいオプション パラメーター `options` が追加されました。 この値は、イベントの実行をキャンセルするために使用されます。

**使用できる場所**: Outlook on the web (クラシック)

### <a name="attachments"></a>添付ファイル

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[AttachmentContent](/javascript/api/outlook/office.attachmentcontent)

添付ファイルのコンテンツを表す新しいオブジェクトが追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

メッセージまたは予定に base 64 エンコード文字列として表されるファイルを添付する新しい方法が追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent)

特定の添付ファイルのコンテンツを取得する新しい方法が追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>[Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails)

新規作成モードでアイテムの添付ファイルを取得する新しい方法が追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

添付ファイルのコンテンツに適用されるフォーマットを特定する新しい列挙型が追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

アイテムから添付ファイルが追加されたか、または削除されたかどうかを特定する新しい列挙型が追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.AttachmentsChanged](/javascript/api/office/office.eventtype)

`AttachmentsChanged` イベントが `Item` に追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

### <a name="delegate-access"></a>代理人アクセス

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[SharedProperties](/javascript/api/outlook/office.sharedproperties)

共有フォルダー、予定表、メールボックスの中の予定やメッセージ アイテムのプロパティを表す新しいオブジェクトが追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

予定やメッセージ アイテムの sharedProperties を表すオブジェクトを取得する新しい方法が追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

代理人のアクセス権を指定する新しいビット フラグ列挙型が追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[SupportsSharedFolders マニフェスト要素](../../manifest/supportssharedfolders.md)

[DesktopFormFactor](../../manifest/desktopformfactor.md) マニフェスト要素に子要素が追加されました。 代理人のシナリオでアドインが使用できるかどうかを定義します。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

### <a name="enhanced-location"></a>強化された場所

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)

予定の場所のセットを表す新しいオブジェクトが追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[LocationDetails](/javascript/api/outlook/office.locationdetails)

場所を表す新しいオブジェクトが追加されました。 読み取り専用です。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[LocationIdentifier](/javascript/api/outlook/office.locationidentifier)

場所の ID を表す新しいオブジェクトが追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation)

予定の場所のセットを表す新しいプロパティが追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype)

予定の場所の種類を指定する新しい列挙型が追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.EnhancedLocationsChanged](/javascript/api/office/office.eventtype)

`EnhancedLocationsChanged` イベントが `Item` に追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

### <a name="integration-with-actionable-messages"></a>操作可能なメッセージとの統合

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

アドインが[操作可能メッセージによってアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。

**使用できる場所**: Office 2019 for Windows (Office 365 サブスクリプション)、Outlook on the web (クラシック)

### <a name="internet-headers"></a>インターネット ヘッダー

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[InternetHeaders](/javascript/api/outlook/office.internetheaders)

メッセージ アイテムのインターネット ヘッダーを表す新しいオブジェクトが追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a>[Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders)

メッセージ アイテムのインターネット ヘッダーを表す新しいプロパティが追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

### <a name="office-theme"></a>Office テーマ

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[Office.context.mailbox.officeTheme](/javascript/api/office/office.officetheme)

Office テーマを取得する機能が追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

`OfficeThemeChanged` イベントが `Mailbox` に追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)

### <a name="sso"></a>SSO

#### <a name="officecontextauthgetaccesstokenasynchttpsdocsmicrosoftcomofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

Microsoft Graph API の[アクセス トークンの取得](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)をアドインに対して許可する、`getAccessTokenAsync` へのアクセスが追加されました。

**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)、Mac 用 Outlook 2019、Outlook on the web (Office 365 および Outlook.com)、Outlook on the web (クラシック)

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](https://docs.microsoft.com/outlook/add-ins/quick-start)
