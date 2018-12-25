---
title: Outlook アドイン API 要件セットのプレビュー
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: e1ed6cae6ac3753f420763b63de0d05283a8fac5
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433664"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook アドイン API 要件セットのプレビュー

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)の**プレビュー**用です。 この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。 アドイン マニフェストでこの要件を指定しないでください。 この要件のセットに導入されているメソッドとプロパティは、使用前に可用性を個別にテストする必要があります。

要件セットのプレビューには、[要件セット 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) のすべての機能が含まれています。

## <a name="features-in-preview"></a>プレビューの機能

次の機能はプレビュー段階です。

- [AttachmentContent](/javascript/api/outlook/office.attachmentcontent) - 添付ファイルのコンテンツを表す新しいオブジェクトが追加されました。
- [InternetHeaders](/javascript/api/outlook/office.internetheaders) - メッセージ アイテムのインターネット ヘッダーを表す新しいオブジェクトが追加されました。
- [SharedProperties](/javascript/api/outlook/office.sharedproperties) - 共有フォルダー、予定表、メールボックスの中の予定やメッセージ アイテムのプロパティを表す新しいオブジェクトが追加されました。
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-) - 1 つの有効な値 `allowEvent` を持つディクショナリである、新しいオプション パラメーター `options`。この値は、イベントの実行をキャンセルするために使用されます。
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - メッセージまたは予定に base 64 エンコード文字列として表されるファイルを添付する新しい方法が追加されました。
- [Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) - 特定の添付ファイルのコンテンツを取得する新しい方法が追加されました。
- [Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) - 作成モードで、アイテムの添付ファイルを取得する新しい方法が追加されました。
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) - アドインが[操作可能メッセージによってアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)されると渡される初期化データを返す新しい機能が追加されました。
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback) - 予定やメッセージ アイテムの sharedProperties を表すオブジェクトを取得する新しい方法が追加されました。
- [Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) - メッセージ アイテムのインターネット ヘッダーを表す新しいプロパティが追加されました。
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) - Microsoft Graph API の[アクセス トークンの取得](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)をアドインに対して許可する、`getAccessTokenAsync` へのアクセスが追加されました。
- [Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) - 添付ファイルのコンテンツに適用されるフォーマットを特定する新しい列挙型が追加されました。
- [Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus) - アイテムから添付ファイルが追加されたか、または削除されたかどうかを特定する新しい列挙型が追加されました。
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) - 代理人のアクセス権を指定する新しいビット フラグ列挙型が追加されました。
- [Office.EventType](/javascript/api/office/office.eventtype) - AttachmentsChanged イベントおよび OfficeThemeChanged イベントを、それぞれに `AttachmentsChanged` エントリと `OfficeThemeChanged` エントリを追加することによりサポートするように変更されました。
- [SupportsSharedFolders manifest element](../../manifest/supportssharedfolders.md) - [DesktopFormFactor](../../manifest/desktopformfactor.md) マニフェスト要素に子要素が追加されました。 代理人のシナリオでアドインが使用できるかどうかを定義します。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](https://docs.microsoft.com/outlook/add-ins/quick-start)