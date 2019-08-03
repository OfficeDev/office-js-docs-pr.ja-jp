---
title: OneNote JavaScript API の要件セット
description: ''
ms.date: 07/17/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 3a1e5133b36af612156fb272651f1775e916a0fe
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064873"
---
# <a name="onenote-javascript-api-requirement-sets"></a>OneNote JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

次の表は、OneNote の要件セット、それらの要件セットをサポートする Office ホスト アプリケーション、ビルド バージョンまたは一般提供開始日の一覧です。

|  要件セット  |  Web 上の Office |
|:-----|:-----|
| [OneNoteApi 1.1](/javascript/api/onenote?view=onenote-js-1.1)  | 2016 年 9 月 |  

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="onenote-javascript-api-11"></a>OneNote JavaScript API 1.1

OneNote JavaScript API 1.1 は、API の最初のバージョンです。 API について詳しくは、「[OneNote の JavaScript API のプログラミングの概要](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)」をご覧ください。

## <a name="runtime-requirement-support-check"></a>ランタイム要件のサポートのチェック

実行時に、アドインは、次の手順に従って、特定のホストが API 要件セットをサポートしているかどうかを確認できます。

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>マニフェストに基づく要件のサポートのチェック

アドインマニフェスト`Requirements`の要素を使用して、アドインが使用する必要がある重要な要件セットまたは API メンバーを指定します。 Office ホストまたはプラットフォームが、 `Requirements`要素で指定されている要件セットや API メンバーをサポートしていない場合、アドインはそのホストまたはプラットフォームでは実行されず、アドインには表示されません。

OneNoteApi 要件セット、バージョン 1.1 をサポートするすべての Office ホスト アプリケーションで読み込まれるアドインのコード例を以下に示します。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [OneNote JavaScript API リファレンスドキュメント](/javascript/api/onenote)
- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
