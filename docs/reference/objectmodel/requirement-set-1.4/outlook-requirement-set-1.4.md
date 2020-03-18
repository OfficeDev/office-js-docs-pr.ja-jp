---
title: Outlook アドイン API 要件セット 1.4
description: Outlook アドインおよび Office JavaScript Api for the Mailbox API 1.4 の一部として導入された機能と Api。
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 5cad049096f18dee925d3ac2ad8802e0047a0ef7
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720044"
---
# <a name="outlook-add-in-api-requirement-set-14"></a>Outlook アドイン API 要件セット 1.4

Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。

## <a name="whats-new-in-14"></a>1.4 の新機能

要件セット 1.4 には、[要件セット 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) のすべての機能が含まれています。`Office.ui` 名前空間へのアクセスが追加されました。

### <a name="change-log"></a>変更ログ

- [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) が追加されました。Office ホストでダイアログ ボックスを表示します。
- [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-) が追加されました。メッセージをダイアログ ボックスからその親/オープナー ページに配信します。
- [Dialog](/javascript/api/office/office.dialog) オブジェクトが追加されました。このオブジェクトは、[`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) メソッドが呼び出されたときに返されます。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](../../../outlook/outlook-add-ins-overview.md)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](../../../quickstarts/outlook-quickstart.md)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
