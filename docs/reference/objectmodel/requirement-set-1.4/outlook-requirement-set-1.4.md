---
title: Outlook アドイン API 要件セット 1.4
description: メールボックス API 1.4 の一部Outlook JavaScript API および Office JavaScript API 用に導入された機能と API。
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: e9e39f3682748498dec38708ee61568d8335b02a
ms.sourcegitcommit: efd0966f6400c8e685017ce0c8c016a2cbab0d5c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/08/2021
ms.locfileid: "60237617"
---
# <a name="outlook-add-in-api-requirement-set-14"></a>Outlook アドイン API 要件セット 1.4

Office Outlook JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。

## <a name="whats-new-in-14"></a>1.4 の新機能

要件セット 1.4 には、要件セット [1.3 のすべての機能が含まれています](../requirement-set-1.3/outlook-requirement-set-1.3.md)。 名前空間へのアクセスが追加されました。

### <a name="change-log"></a>変更ログ

- [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#displayDialogAsync_startAddress__options__callback_): アプリケーション内のダイアログ ボックスをOfficeしました。
- [Office.context.ui.messageParent](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#messageParent_message__messageOptions_) が追加されました。メッセージをダイアログ ボックスからその親/オープナー ページに配信します。
- [Dialog](/javascript/api/office/office.dialog?view=outlook-js-1.4&preserve-view=true) オブジェクトが追加されました。このオブジェクトは、[`displayDialogAsync`](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#displayDialogAsync_startAddress__options__callback_) メソッドが呼び出されたときに返されます。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
