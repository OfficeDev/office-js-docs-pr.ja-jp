---
title: Outlook アドイン API 要件セット 1.2
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 898e768dfc1828ba44f29e9da5c4baa61de186cb
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902096"
---
# <a name="outlook-add-in-api-requirement-set-12"></a>Outlook アドイン API 要件セット 1.2

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)のためのものです。 

## <a name="whats-new-in-12"></a>1.2 の新機能

要件セット 1.2 には、[要件セット 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) のすべての機能が含まれています。アドインを使用して、メッセージの件名または本文内のいずれかで、ユーザーのカーソル位置にテキストを挿入する機能が追加されました。

### <a name="change-log"></a>変更ログ

- [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string) が追加されました: メッセージの件名または本文から、選択されたデータを非同期的に返します。
- [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback) が追加されました。メッセージの本文または件名に非同期的にデータを挿入します。
- [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback) が変更されました。`attachments` パラメーターに `formData` プロパティが追加されました。
- [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback) が変更されました。`formData` パラメーターに `attachments` プロパティが追加されました。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](/outlook/add-ins/quick-start)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
