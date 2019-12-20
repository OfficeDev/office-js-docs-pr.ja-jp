---
title: Outlook アドイン API 要件セット 1.5
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 1a12156feb7a03e596e521650a757fe7198b4d76
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814746"
---
# <a name="outlook-add-in-api-requirement-set-15"></a>Outlook アドイン API 要件セット 1.5

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)のためのものです。

## <a name="whats-new-in-15"></a>1.5 の新機能

要件セット 1.5 には、[要件セット 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) のすべての機能が含まれています。次の機能が追加されました。

- [ピン留め可能な作業ウィンドウ](/outlook/add-ins/pinnable-taskpane)のサポートが追加されました。
- [REST API](/outlook/add-ins/use-rest-api) の呼び出しのサポートが追加されました。
- インラインで添付ファイルにマークを付ける機能が追加されました。
- 作業ウィンドウまたはダイアログを閉じる機能が追加されました。

### <a name="change-log"></a>変更ログ

- [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods) が追加されました。サポートされているイベントのイベント ハンドラーを追加します。
- 追加された、サポートされているイベントの種類のイベントハンドラを削除[し](office.context.mailbox.md#methods)ました。
- [Office.EventType](office.md#eventtype-string) が追加されました。イベント ハンドラーに関連付けられているイベントを指定し、ItemChanged イベントのサポートが含まれるようになります。
- [Office.context.mailbox.restUrl](office.context.mailbox.md#properties) が追加されました。この電子メール アカウントの REST エンドポイントの URL を取得します。
- [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods) が変更されました。このメソッドの新しい署名付きの新しいバージョン (`getCallbackTokenAsync([options], callback)`) が追加されました。元のバージョンは引き続き使用でき、変更されていません。
- [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) が追加されました。
- [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods) が変更されました。`isInline` と呼ばれる `options` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。
- [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) が変更されました。`formData.attachments` と呼ばれる `isInline` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。
- [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) が変更されました。`isInline` と呼ばれる `formData.attachments` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](/outlook/add-ins/quick-start)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
