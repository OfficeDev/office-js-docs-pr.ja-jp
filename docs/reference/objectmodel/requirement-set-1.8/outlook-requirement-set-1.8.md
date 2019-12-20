---
title: Outlook アドイン API 要件セット 1.8
description: ''
ms.date: 12/17/2019
localization_priority: Priority
ms.openlocfilehash: 1a542a0b00bd04674c4da4c45909ff03a152983c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814160"
---
# <a name="outlook-add-in-api-requirement-set-18"></a>Outlook アドイン API 要件セット 1.8

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

## <a name="whats-new-in-18"></a>1.8 の新機能

要件セット 1.8 には、[要件セット 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) のすべての機能が含まれています。 次の機能が追加されました。

- 添付ファイル、カテゴリ、代理人アクセス、拡張された場所、インターネット ヘッダー、および送信ブロック機能用の新しい API が追加されました。
- Event.completed にオプションの `options` パラメーターが追加されました。
- AttachmentsChanged および EnhancedLocationsChanged イベントのサポートが追加されました。

### <a name="change-log"></a>変更ログ

- [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8) が追加されました: 添付ファイルのコンテンツを表す新しいオブジェクトを追加します。
- [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.8) が追加されました: 項目カテゴリを表す新しいオブジェクトを追加します。
- [CategoryDetails](/javascript/api/outlook/office.categorydetails?view=outlook-js-1.8) が追加されました: カテゴリの詳細 (名前とそれに関連付けられた色) を表す新しいオブジェクトを追加します。
- [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) が追加されました: 予定の場所のセットを表す新しいオブジェクトを追加します。
- [InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8) が追加されました: メッセージ アイテムのインターネット ヘッダーを表す新しいオブジェクトを追加します。 新規作成モードのみです。
- [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8) が追加されました: 場所を表す新しいオブジェクトを追加します。 読み取り専用です。
- [LocationIdentifier](/javascript/api/outlook/office.locationidentifier?view=outlook-js-1.8) が追加されました: 場所の ID を表す新しいオブジェクトを追加します。
- [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8) が追加されました: メールボックスのカテゴリ マスター リストを表す新しいオブジェクトを追加します。
- [SharedProperties](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) が追加されました: 共有フォルダー、予定表、メールボックスの中の予定やメッセージ アイテムのプロパティを表す新しいオブジェクトを追加します。
- [SupportsSharedFolders マニフェスト要素](../../manifest/supportssharedfolders.md) が追加されました: [DesktopFormFactor](../../manifest/desktopformfactor.md) マニフェスト要素に子要素を追加します。 代理人のシナリオでアドインが使用できるかどうかを定義します。
- [Office.context.mailbox.masterCategories](office.context.mailbox.md#properties) が追加されました: メールボックスのカテゴリ マスター リストを表す新しいプロパティを追加します。
- [Office.context.mailbox.item.categories](office.context.mailbox.item.md#properties) が追加されました: アイテムのカテゴリのセットを表す新しいプロパティを追加します。
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#methods) が追加されました: メッセージまたは予定に Base 64 エンコード文字列として表されるファイルを添付する新しい方法を追加します。
- [Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#properties) が追加されました: 予定の場所のセットを表す新しいプロパティを追加します。
- [Office.context.mailbox.item.getAllInternetHeadersAsync](office.context.mailbox.item.md#methods) が追加されました: メッセージ アイテムのすべてのインターネット ヘッダーを取得する新しいメソッドを追加します。 閲覧モードのみ。
- [Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#methods) が追加されました: 特定の添付ファイルのコンテンツを取得する新しい方法を追加します。
- [Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#methods) が追加されました: 作成モードで、アイテムの添付ファイルを取得する新しい方法を追加します。
- [Office.context.mailbox.item.getItemIdAsync](office.context.mailbox.item.md#methods) が追加されました: 保存済みの予定またはメッセージ アイテムの ID を取得する新しい方法を追加します。
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#methods) が追加されました: 予定やメッセージ アイテムの sharedProperties を表すオブジェクトを取得する新しい方法を追加します。
- [Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#properties) が追加されました: メッセージ アイテムのインターネット ヘッダーを表す新しいプロパティを追加します。 新規作成モードのみです。
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-) が変更されました: 1 つの有効な値 `allowEvent` を持つディクショナリである、新しいオプション `options` パラメーターを追加します 。 この値は、イベントの実行をキャンセルするために使用されます。
- [Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8) が追加されました: 添付ファイルのコンテンツに適用される書式を特定する新しい列挙型を追加します。
- [Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus?view=outlook-js-1.8) が追加されました: アイテムから添付ファイルが追加されたか、または削除されたかどうかを特定する新しい列挙型を追加します。
- [Office.MailboxEnums.CategoryColor](/javascript/api/outlook/office.mailboxenums.categorycolor?view=outlook-js-1.8) が追加されました: カテゴリに関連付ける使用可能な色を指定する新しい列挙を追加します。
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions?view=outlook-js-1.8) が追加されました: 代理人のアクセス権を指定する新しいビット フラグ列挙型を追加します。
- [Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype?view=outlook-js-1.8) が追加されました: 予定の場所の種類を指定する新しい列挙型を追加します。
- [Office.EventType](/javascript/api/office/office.eventtype) が変更されました: `AttachmentsChanged` と `EnhancedLocationsChanged` のイベントにサポートを追加します。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](/outlook/add-ins/quick-start)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
