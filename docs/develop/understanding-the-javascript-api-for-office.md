---
title: Office JavaScript API について
description: Office JavaScript API の概要
ms.date: 02/27/2020
localization_priority: Priority
ms.openlocfilehash: 28aac00cd801019d95b2d4b487bfdf8e2187af95
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324700"
---
# <a name="understanding-the-office-javascript-api"></a>Office JavaScript API について

Office アドインでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメント内のコンテンツを操作できます。

## <a name="accessing-the-office-javascript-api-library"></a>Office JavaScript API ライブラリへのアクセス

[!include[information about accessing the Office JS API library](../includes/office-js-access-library.md)]

## <a name="api-models"></a>API モデル

[!include[information about the Office JS API models](../includes/office-js-api-models.md)]

## <a name="api-requirement-sets"></a>API 要件セット

[!include[information about the Office JS API requirement sets](../includes/office-js-requirement-sets.md)]

> [!NOTE]
> AppSource にアドインを[公開](../publish/publish.md)し、Office エクスペリエンスで利用できるようにする予定がある場合は、[AppSource の検証ポリシー](/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[セクション 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)のページを参照してください)。 

## <a name="see-also"></a>関連項目

- [Office JavaScript API リファレンス](../reference/javascript-api-for-office.md)
- [DOM とランタイム環境を読み込む](loading-the-dom-and-runtime-environment.md)
- [Office JavaScript API ライブラリの参照](referencing-the-javascript-api-for-office-library-from-its-cdn.md)
- [Office アドインを初期化する](initialize-add-in.md)
