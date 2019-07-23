---
title: Outlook アドイン API 要件セットのプレビュー
description: ''
ms.date: 07/18/2019
localization_priority: Priority
ms.openlocfilehash: f2030ca40a83694c1c1c92ca4b956a9294cb3e0a
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804991"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook アドイン API 要件セットのプレビュー

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)の**プレビュー**用です。 この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。 アドイン マニフェストでこの要件を指定しないでください。 この要件のセットに導入されているメソッドとプロパティは、使用前に可用性を個別にテストする必要があります。 また、場合によっては [Office Insider プログラム](https://products.office.com/office-insider)に参加する必要もあります。

要件セットのプレビューには、[要件セット 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) のすべての機能が含まれています。

## <a name="features-in-preview"></a>プレビューの機能

次の機能はプレビュー段階です。

### <a name="attachments"></a>添付ファイル

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[AttachmentContent](/javascript/api/outlook/office.attachmentcontent)

添付ファイルのコンテンツを表す新しいオブジェクトが追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

メッセージまたは予定に base 64 エンコード文字列として表されるファイルを添付する新しい方法が追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

特定の添付ファイルのコンテンツを取得する新しい方法が追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

新規作成モードでアイテムの添付ファイルを取得する新しい方法が追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

添付ファイルのコンテンツに適用されるフォーマットを特定する新しい列挙型が追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

アイテムから添付ファイルが追加されたか、または削除されたかどうかを特定する新しい列挙型が追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.AttachmentsChanged](/javascript/api/office/office.eventtype)

`AttachmentsChanged` イベントが `Item` に追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

---

### <a name="block-on-send"></a>送信のブロック

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[Event.completed](/javascript/api/office/office.addincommands.event#completed-options-)

1 つの有効な値 `allowEvent` を持つディクショナリである、新しいオプション パラメーター `options` が追加されました。 この値は、イベントの実行をキャンセルするために使用されます。

**使用できる場所**: Outlook on the web (クラシック)

---

### <a name="categories"></a>カテゴリ

Outlook では、ユーザーはカテゴリを使用してメッセージと予定を色分けしてグループ化できます。 ユーザーは自分のメールボックスのマスター リストにカテゴリを定義します。 その後、アイテムに 1 つ以上のカテゴリを適用できます。

> [!NOTE]
> この機能は Outlook on iOS または Android ではサポートされていません。

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[Categories](/javascript/api/outlook/office.categories)

項目カテゴリを表す新しいオブジェクトが追加されました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[CategoryDetails](/javascript/api/outlook/office.categorydetails)

カテゴリの詳細 (名前とそれに関連付けられた色) を表す新しいオブジェクトが追加されました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[MasterCategories](/javascript/api/outlook/office.mastercategories)

メールボックスのカテゴリ マスター リストを表す新しいオブジェクトが追加されました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[Office.context.mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)

メールボックスのカテゴリ マスター リストを表す新しいプロパティが追加されました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[Office.context.mailbox.item.categories](/javascript/api/outlook/office.item#categories)

アイテムのカテゴリのセットを表す新しいプロパティが追加されました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[Office.MailboxEnums.CategoryColor](/javascript/api/outlook/office.mailboxenums.categorycolor)

カテゴリに関連付ける使用可能な色を指定する新しい列挙が追加されました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

---

### <a name="delegate-access"></a>代理人アクセス

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[SharedProperties](/javascript/api/outlook/office.sharedproperties)

共有フォルダー、予定表、メールボックスの中の予定やメッセージ アイテムのプロパティを表す新しいオブジェクトが追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[Office.context.mailbox.item.getItemIdAsync](office.context.mailbox.item.md#getitemidasyncoptions-callback)

保存済みの予定またはメッセージ アイテムの ID を取得する新しいメソッドが追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

予定やメッセージ アイテムの sharedProperties を表すオブジェクトを取得する新しい方法が追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

代理人のアクセス権を指定する新しいビット フラグ列挙型が追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[SupportsSharedFolders マニフェスト要素](../../manifest/supportssharedfolders.md)

[DesktopFormFactor](../../manifest/desktopformfactor.md) マニフェスト要素に子要素が追加されました。 代理人のシナリオでアドインが使用できるかどうかを定義します。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

---

### <a name="enhanced-location"></a>強化された場所

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)

予定の場所のセットを表す新しいオブジェクトが追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[LocationDetails](/javascript/api/outlook/office.locationdetails)

場所を表す新しいオブジェクトが追加されました。 読み取り専用です。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[LocationIdentifier](/javascript/api/outlook/office.locationidentifier)

場所の ID を表す新しいオブジェクトが追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

予定の場所のセットを表す新しいプロパティが追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype)

予定の場所の種類を指定する新しい列挙型が追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.EnhancedLocationsChanged](/javascript/api/office/office.eventtype)

`EnhancedLocationsChanged` イベントが `Item` に追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)

---

### <a name="integration-with-actionable-messages"></a>操作可能なメッセージとの統合

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (クラシック)

---

### <a name="internet-headers"></a>インターネット ヘッダー

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[InternetHeaders](/javascript/api/outlook/office.internetheaders)

メッセージ アイテムのカスタム インターネット ヘッダーを表す新しいオブジェクトが追加されました。

**使用できる場所**: Windows 用 Outlook (Office 365 サブスクリプションに接続している場合)

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheaders)

メッセージ アイテムのカスタム インターネット ヘッダーを表す新しいプロパティが追加されました。

**使用できる場所**: Windows 用 Outlook (Office 365 サブスクリプションに接続している場合)

---

### <a name="office-theme"></a>Office テーマ

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Office テーマを取得する機能が追加されました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

`OfficeThemeChanged` イベントが `Mailbox` に追加されました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

---

### <a name="sso"></a>SSO

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[Office.context.auth.getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

Microsoft Graph API の[アクセス トークンの取得](/outlook/add-ins/authenticate-a-user-with-an-sso-token)をアドインに対して許可する、`getAccessTokenAsync` へのアクセスが追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac、Outlook on the web (モダン)、Outlook on the web (クラシック)

## <a name="see-also"></a>関連項目

- [Outlook アドイン](/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](/outlook/add-ins/quick-start)
- [要求セットとサポートされているクライアント](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
